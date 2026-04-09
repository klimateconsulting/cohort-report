"""
50001 Ready Cohort Program - HTML Report Generator

This script generates the EERE 50001 Ready Cohort Program Update HTML report.
It reads from the live Excel file and allows user-defined date ranges.

Usage:
    python generate_cohort_report.py --start 2020-07-01 --end 2025-09-30
    python generate_cohort_report.py  # Uses default date range
"""

import pandas as pd
import folium
import plotly.express as px
from datetime import datetime
import plotly.offline as po
import argparse
import os

# =============================================================================
# CONFIGURATION
# =============================================================================

def get_latest_closed_quarter():
    """
    Calculate the latest closed quarter based on the current date.
    Returns (start_date, end_date) as strings in YYYY-MM-DD format.

    Quarters:
    - Q1: Jan 1 - Mar 31
    - Q2: Apr 1 - Jun 30
    - Q3: Jul 1 - Sep 30
    - Q4: Oct 1 - Dec 31
    """
    today = datetime.now()
    current_month = today.month
    current_year = today.year

    # Determine which quarter just closed
    if current_month >= 1 and current_month <= 3:
        # We're in Q1, so last closed quarter is Q4 of previous year
        start = datetime(current_year - 1, 10, 1)
        end = datetime(current_year - 1, 12, 31)
    elif current_month >= 4 and current_month <= 6:
        # We're in Q2, so last closed quarter is Q1
        start = datetime(current_year, 1, 1)
        end = datetime(current_year, 3, 31)
    elif current_month >= 7 and current_month <= 9:
        # We're in Q3, so last closed quarter is Q2
        start = datetime(current_year, 4, 1)
        end = datetime(current_year, 6, 30)
    else:
        # We're in Q4, so last closed quarter is Q3
        start = datetime(current_year, 7, 1)
        end = datetime(current_year, 9, 30)

    return start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d')

# Color mapping for offices
COLOR_MAPPING = {
    'BTO': '#94C155',
    'FEMP': '#185E25',
    'ITO': '#0C2E60',
    'SCEP': '#4FA6DC',
}

# =============================================================================
# DATA LOADING
# =============================================================================

def load_data_from_excel(excel_path):
    """Load data from the live Excel file."""
    df = pd.read_excel(excel_path)
    return df

def convert_date(date_str):
    """Convert date string to datetime."""
    try:
        return datetime.strptime(str(date_str), '%m/%d/%Y')
    except (ValueError, TypeError):
        return None

def process_data(df, start_date, end_date):
    """Process and filter the dataframe."""
    # Make a copy to avoid SettingWithCopyWarning
    df = df.copy()

    # Convert dates if they're not already datetime
    # Excel files typically have dates as datetime64 already
    if df['Start'].dtype == 'object':
        df['Start'] = df['Start'].apply(convert_date)
    if df['Planned End Date'].dtype == 'object':
        df['Planned End Date'] = df['Planned End Date'].apply(convert_date)

    # Convert start_date to pandas Timestamp for comparison
    start_date_ts = pd.Timestamp(start_date)

    # Filter the dataframe based on the user-defined date range
    df = df[(df['Planned End Date'] >= start_date_ts)].copy()

    # Ensure latitude and longitude are of correct type
    df['Latitude'] = pd.to_numeric(df['Latitude'], errors='coerce')
    df['Longitude'] = pd.to_numeric(df['Longitude'], errors='coerce')

    # Fill missing values with placeholders
    df['City'] = df['City'].fillna('Unknown City')
    df['State'] = df['State'].fillna('Unknown State')

    return df

# =============================================================================
# MAP CREATION
# =============================================================================

def create_map(df, output_path):
    """Create the participation map using Folium."""
    # Aggregate cohorts by office and location
    aggregated_df = df.groupby(
        ['Office', 'City', 'State', 'COUNTRY', 'Latitude', 'Longitude']
    ).size().reset_index(name='Count')

    # Initialize the map centered around the US and Hawaii
    map_center = [36.100137, -5.666518]
    m = folium.Map(location=map_center, zoom_start=2)

    # Create a feature group for each office
    feature_groups = {office: folium.FeatureGroup(name=office) for office in COLOR_MAPPING.keys()}

    # Add points to the corresponding feature group
    for index, row in aggregated_df.iterrows():
        popup_text = (
            f"Office: {row['Office']}<br>"
            f"City: {row['City']}<br>"
            f"State: {row['State']}<br>"
            f"Country: {row['COUNTRY']}<br>"
            f"Count: {row['Count']}<br>"
            f"Location: {row['Latitude']}, {row['Longitude']}"
        )

        iframe = folium.IFrame(html=popup_text, width=200, height=100)
        popup = folium.Popup(iframe, max_width=200)

        folium.CircleMarker(
            location=[row['Latitude'], row['Longitude']],
            radius=5 + row['Count'],  # Adjust radius based on count
            popup=popup,
            color=COLOR_MAPPING.get(row['Office'], 'gray'),
            fill=True,
            fill_color=COLOR_MAPPING.get(row['Office'], 'gray'),
            fill_opacity=0.7
        ).add_to(feature_groups.get(row['Office'], feature_groups.get('BTO')))

    # Add feature groups to the map
    for office, feature_group in feature_groups.items():
        feature_group.add_to(m)

    # Add custom legend HTML
    legend_html = '''
    <div id="maplegend" class="maplegend">
    <div class="legend-title">Office Layers</div>
    <div class="legend-scale">
      <ul class="legend-labels">
        <li><span style="background-color:#94C155;"></span>BTO</li>
        <li><span style="background-color:#185E25;"></span>FEMP</li>
        <li><span style="background-color:#0C2E60;"></span>ITO</li>
        <li><span style="background-color:#4FA6DC;"></span>SCEP</li>
      </ul>
    </div>
    </div>
    <style type="text/css">
      .maplegend {
        position: absolute;
        z-index:9999;
        bottom: 50px;
        left: 50px;
        border:2px solid grey;
        background-color:white;
        opacity: 0.8;
        padding: 10px;
        font-size:12px;
        color: #000;
      }
      .maplegend .legend-title {
        text-align: left;
        margin-bottom: 5px;
        font-weight: bold;
        font-size: 90%;
      }
      .maplegend .legend-scale ul {
        margin: 0;
        padding: 0;
        list-style: none;
      }
      .maplegend .legend-scale ul li {
        list-style: none;
        margin-left: 0;
        line-height: 18px;
        margin-bottom: 2px;
      }
      .maplegend ul.legend-labels li span {
        display: block;
        float: left;
        height: 16px;
        width: 16px;
        margin-right: 5px;
        margin-left: 0;
        border: 1px solid #999;
      }
    </style>
    '''

    # Add the legend to the map
    m.get_root().html.add_child(folium.Element(legend_html))

    # Add layer control to the map
    folium.LayerControl().add_to(m)

    # Save the map to an HTML file
    m.save(output_path)

    return m

# =============================================================================
# TIMELINE CREATION
# =============================================================================

def create_timeline(df):
    """Create the timeline chart using Plotly."""
    # Aggregate by 'Cohort Name'
    aggregated_df = df.groupby(['Office', 'Cohort Name', 'Coach']).agg({
        'Start': 'min',
        'Planned End Date': 'max'
    }).reset_index()

    # Create unique y-axis labels by combining Office and Cohort Name
    aggregated_df['Unique Cohort Name'] = aggregated_df['Office'] + ': ' + aggregated_df['Cohort Name']

    # Sort by 'Start' date
    aggregated_df = aggregated_df.sort_values(by=['Office', 'Start'])

    # Format the Start and End dates as strings without time
    aggregated_df['Start_str'] = aggregated_df['Start'].dt.strftime('%Y-%m-%d')
    aggregated_df['End_str'] = aggregated_df['Planned End Date'].dt.strftime('%Y-%m-%d')

    # Create the timeline bar chart
    fig = px.timeline(
        aggregated_df,
        x_start="Start",
        x_end="Planned End Date",
        y="Unique Cohort Name",
        color="Office",
        color_discrete_map=COLOR_MAPPING,
        title='Cohort Timelines by Office',
        height=800,
        custom_data=aggregated_df[['Office', 'Cohort Name', 'Coach', 'Start_str', 'End_str']]
    )

    # Update the layout to include Office, Cohort Name, Coach, Start, and End in hover data
    fig.update_traces(
        hovertemplate="<br>".join([
            "Office: %{customdata[0]}",
            "Cohort Name: %{customdata[1]}",
            "Coach: %{customdata[2]}",
            "Start: %{customdata[3]}",
            "End: %{customdata[4]}"
        ]),
    )

    # Update the layout
    fig.update_layout(
        xaxis_title='Date',
        yaxis_title='Cohort Name',
        legend_title='Office',
        hovermode='closest'
    )

    # Return as div
    timeline_div = po.plot(fig, include_plotlyjs=False, output_type='div')
    return timeline_div

# =============================================================================
# SUMMARY TABLE CREATION
# =============================================================================

def create_summary_table(df):
    """Create the summary statistics table."""
    # Define a function to categorize implementation status
    def categorize_implementation(status):
        if status == 'Yes':
            return 'Implementation Track Sites'
        elif status == 'No':
            return 'Training Only Sites'
        elif status == 'Dropout':
            return 'Dropout Sites'
        else:
            return None

    # Apply the function to categorize the implementation status
    df_copy = df.copy()
    df_copy['Implementation Category'] = df_copy['Implementation'].apply(categorize_implementation)

    # Group by Office and Implementation Category, and count the occurrences
    category_counts = df_copy.groupby(['Office', 'Implementation Category']).size().unstack(fill_value=0)

    # Sum the total participants for each office
    total_participants = df_copy.groupby('Office')['Participants'].sum()

    # Count the number of recognized sites
    recognized_sites = df_copy[df_copy['Date of Recognition'].notna()].groupby('Office').size()

    # Combine the category counts and total participants into a summary DataFrame
    summary_df = category_counts.join(total_participants).join(recognized_sites.rename('Recognized Sites')).reset_index()

    # Fill any missing categories with zeros
    summary_df = summary_df.fillna(0)

    # Calculate the % of Implementation Track Sites that are recognized
    summary_df['% of Implementation Track Sites Recognized'] = (
        summary_df['Recognized Sites'] / summary_df['Implementation Track Sites']
    ) * 100
    summary_df['% of Implementation Track Sites Recognized'] = summary_df[
        '% of Implementation Track Sites Recognized'
    ].map("{:.0f}%".format)

    # List of required columns
    required_columns = [
        'Office', 'Dropout Sites', 'Implementation Track Sites', 'Training Only Sites',
        'Participants', 'Recognized Sites', '% of Implementation Track Sites Recognized'
    ]

    # Add missing columns with empty values
    for column in required_columns:
        if column not in summary_df.columns:
            summary_df[column] = 0

    # Ensure the columns are in the correct order
    summary_df = summary_df[required_columns]

    # Rename columns for better readability
    summary_df.rename(columns={'Participants': 'Total Participants'}, inplace=True)

    # Reorder columns to match the desired output
    summary_df = summary_df[[
        'Office', 'Implementation Track Sites', 'Training Only Sites', 'Dropout Sites',
        'Total Participants', 'Recognized Sites', '% of Implementation Track Sites Recognized'
    ]]

    # Add a Total row
    total_row = summary_df.sum(numeric_only=True)
    total_row['Office'] = 'Total'
    total_row['% of Implementation Track Sites Recognized'] = "{:.0f}%".format(
        (total_row['Recognized Sites'] / total_row['Implementation Track Sites']) * 100
    )

    # Convert total_row to a DataFrame
    total_row_df = pd.DataFrame([total_row])

    # Concatenate the total_row_df to summary_df
    summary_df = pd.concat([summary_df, total_row_df], ignore_index=True)

    # Convert all numeric columns to integers
    summary_df = summary_df.astype({
        'Implementation Track Sites': 'int',
        'Training Only Sites': 'int',
        'Dropout Sites': 'int',
        'Total Participants': 'int',
        'Recognized Sites': 'int'
    })

    return summary_df

# =============================================================================
# COMPLETED COHORTS TABLE CREATION
# =============================================================================

def create_completed_cohorts_table(df):
    """Create the completed cohorts and recognition status table."""
    df_copy = df.copy()

    # Handle Actual End Date - may already be datetime from Excel
    if 'Actual End Date' in df_copy.columns:
        if df_copy['Actual End Date'].dtype == 'object':
            df_copy['End'] = df_copy['Actual End Date'].apply(convert_date)
        else:
            df_copy['End'] = df_copy['Actual End Date']
    else:
        df_copy['End'] = None

    # Filter to only rows with valid End dates
    df_with_end = df_copy[df_copy['End'].notna()].copy()

    if len(df_with_end) == 0:
        # No completed cohorts
        completed_cohorts_df = pd.DataFrame(columns=[
            'Office', 'Cohort Name', 'Coach', 'Actual End Date',
            'Completed Program', 'Ready Recognition Status'
        ])
    else:
        # Aggregate by cohort
        completed_cohorts_df = df_with_end.groupby(['Office', 'Cohort Name', 'Coach', 'End']).agg({
            'Completed': lambda x: (x == 'Yes').sum(),
            'Likelihood of Recognition': lambda x: ', '.join([
                f"{x.value_counts()[i]} {i}" for i in x.value_counts().index
            ])
        }).reset_index()

        completed_cohorts_df.columns = [
            'Office', 'Cohort Name', 'Coach', 'Actual End Date',
            'Completed Program', 'Ready Recognition Status'
        ]

    return completed_cohorts_df

# =============================================================================
# HTML REPORT GENERATION
# =============================================================================

def create_html_report(df, start_date, end_date, map_html, timeline_div,
                       summary_table_html, completed_cohorts_table_html):
    """Create the combined HTML report (static version)."""
    html_template = f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>EERE 50001 Ready Cohort Program Update</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f9f9f9;
        }}
        h1 {{
            text-align: center;
            color: white;
            background-color: #94C155;
            padding: 20px;
        }}
        h2 {{
            text-align: center;
            color: #333;
            margin-top: 40px;
        }}
        p {{
            text-align: center;
            color: #666;
        }}
        .iframe-container {{
            width: 80%;
            height: 600px;
            margin: 20px auto;
            border: 1px solid #ddd;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            clear: both;
        }}
        .iframe-content {{
            width: 100%;
            height: 100%;
            border: 0;
        }}
        .table-container {{
            width: 80%;
            margin: 40px auto;
            clear: both;
        }}
        .table {{
            width: 100%;
            margin: auto;
            border-collapse: collapse;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }}
        .table th, .table td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
        }}
        .table th {{
            padding-top: 12px;
            padding-bottom: 12px;
            background-color: #4CAF50;
            color: white;
        }}
        .table tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        .table tr:hover {{
            background-color: #ddd;
        }}
    </style>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
</head>
<body>
    <h1>50001 Ready Cohort Program Update</h1>
    <h2>Report for Active Cohorts Between {start_date.strftime('%B %d, %Y')} and {end_date.strftime('%B %d, %Y')}</h2>


    <h2>Map of Participating Sites in 50001 Ready Cohorts</h2>
    <div class="iframe-container">
        <iframe class="iframe-content" srcdoc="{map_html.replace('"', '&quot;')}"></iframe>
    </div>

    <h2>Timelines for Cohorts</h2>
    <div class="iframe-container">
        {timeline_div}
    </div>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <h2>Number of Sites and Participants Engaged in the Program</h2>
    <div class="table-container">
        {summary_table_html}
    </div>

    <h2>Completed Cohorts and Ready Recognition Status</h2>
    <div class="table-container">
        {completed_cohorts_table_html}
    </div>

</body>
</html>
'''
    return html_template


def create_interactive_html_report(df_original, start_date, end_date):
    """
    Create an interactive HTML report with date pickers that dynamically
    filter all visualizations and tables.
    """
    import json

    # Prepare data for JSON embedding
    # Convert dataframe to JSON-serializable format
    df_json = df_original.copy()

    # Convert datetime columns to ISO strings for JSON
    date_columns = ['Start', 'Planned End Date', 'Actual End Date', 'Date of Recognition',
                    'Date of last contact?', 'Anticipated Recognition Date?']
    for col in date_columns:
        if col in df_json.columns:
            df_json[col] = df_json[col].apply(
                lambda x: x.isoformat() if pd.notna(x) and hasattr(x, 'isoformat') else None
            )

    # Convert to JSON string
    data_json = df_json.to_json(orient='records')

    # Color mapping as JSON
    color_mapping_json = json.dumps(COLOR_MAPPING)

    # Calculate the minimum start date from the data (oldest timestamp available)
    min_start_date = df_original['Start'].min()
    if pd.notna(min_start_date):
        min_start_date_str = pd.Timestamp(min_start_date).strftime('%Y-%m-%d')
    else:
        min_start_date_str = start_date.strftime('%Y-%m-%d')

    html_template = f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>EERE 50001 Ready Cohort Program Update - Interactive</title>

    <!-- Leaflet CSS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />

    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f9f9f9;
        }}
        h1 {{
            text-align: center;
            color: white;
            background-color: #94C155;
            padding: 20px;
            margin: 0;
        }}
        h2 {{
            text-align: center;
            color: #333;
            margin-top: 40px;
        }}
        .date-filter-container {{
            background-color: #e8f5e9;
            padding: 20px;
            text-align: center;
            border-bottom: 2px solid #94C155;
        }}
        .date-filter-container label {{
            font-weight: bold;
            margin-right: 10px;
            color: #333;
        }}
        .date-filter-container input[type="date"] {{
            padding: 8px 12px;
            font-size: 16px;
            border: 1px solid #ccc;
            border-radius: 4px;
            margin-right: 20px;
        }}
        .date-filter-container button {{
            padding: 10px 25px;
            font-size: 16px;
            background-color: #94C155;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-weight: bold;
        }}
        .date-filter-container button:hover {{
            background-color: #7da844;
        }}
        #date-range-display {{
            text-align: center;
            font-size: 18px;
            color: #333;
            margin: 20px 0;
        }}
        #record-count {{
            text-align: center;
            font-size: 14px;
            color: #666;
            margin-bottom: 20px;
        }}
        .viz-row {{
            display: flex;
            justify-content: center;
            gap: 30px;
            margin: 20px auto;
            width: 90%;
        }}
        .viz-wrapper {{
            text-align: center;
            width: 42%;
        }}
        .viz-wrapper h3 {{
            color: #333;
            margin-bottom: 10px;
        }}
        .viz-container {{
            width: 100%;
            aspect-ratio: 1 / 1;
            border: 1px solid #ddd;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }}
        #map {{
            width: 100%;
            height: 100%;
        }}
        #timeline {{
            width: 100%;
            height: 100%;
        }}
        .table-container {{
            width: 80%;
            margin: 40px auto;
            overflow-x: auto;
        }}
        .table {{
            width: 100%;
            margin: auto;
            border-collapse: collapse;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }}
        .table th, .table td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
        }}
        .table th {{
            padding-top: 12px;
            padding-bottom: 12px;
            background-color: #4CAF50;
            color: white;
        }}
        .table tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        .table tr:hover {{
            background-color: #ddd;
        }}
        .legend {{
            padding: 10px;
            background: white;
            border-radius: 5px;
            box-shadow: 0 0 15px rgba(0,0,0,0.2);
        }}
        .legend-title {{
            font-weight: bold;
            margin-bottom: 8px;
        }}
        .legend-item {{
            display: flex;
            align-items: center;
            margin: 4px 0;
        }}
        .legend-color {{
            width: 16px;
            height: 16px;
            border-radius: 50%;
            margin-right: 8px;
            border: 1px solid #999;
        }}

        /* Mobile overlay */
        #mobile-overlay {{
            display: none;
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0,0,0,0.85);
            z-index: 10000;
            justify-content: center;
            align-items: center;
            text-align: center;
            padding: 30px;
        }}
        #mobile-overlay .mobile-message {{
            background: white;
            border-radius: 12px;
            padding: 40px 30px;
            max-width: 360px;
        }}
        #mobile-overlay .mobile-message h2 {{
            color: #333;
            margin: 0 0 16px 0;
            font-size: 20px;
        }}
        #mobile-overlay .mobile-message p {{
            color: #666;
            margin: 0;
            font-size: 15px;
            line-height: 1.5;
        }}
        .download-btn {{
            display: inline-block;
            margin-top: 10px;
            padding: 8px 18px;
            font-size: 13px;
            background-color: #0C2E60;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-weight: bold;
        }}
        .download-btn:hover {{
            background-color: #1a4a8a;
        }}
        .download-btn:disabled {{
            opacity: 0.6;
            cursor: not-allowed;
        }}
        @media (max-width: 768px) {{
            #mobile-overlay {{ display: flex; }}
        }}
    </style>

    <!-- Leaflet JS -->
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <!-- Plotly JS -->
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <!-- html2canvas for map screenshot -->
    <script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
</head>
<body>
    <div id="mobile-overlay">
        <div class="mobile-message">
            <h2>Desktop Recommended</h2>
            <p>This dashboard contains interactive charts and maps that are best viewed on a desktop or laptop computer. Please open this page on a larger screen for the full experience.</p>
        </div>
    </div>
    <h1>50001 Ready Cohort Program Update</h1>

    <!-- Fetch Latest Data Button -->
    <div style="text-align: center; padding: 10px; background-color: #f0f0f0; border-bottom: 1px solid #ddd;">
        <button id="fetch-btn" onclick="fetchLatestData()" style="padding: 8px 20px; font-size: 14px; background-color: #0C2E60; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">
            Fetch Latest Data
        </button>
        <span id="fetch-status" style="margin-left: 10px; font-size: 14px; color: #666;"></span>
    </div>

    <script>
        function fetchLatestData() {{
            const btn = document.getElementById('fetch-btn');
            const status = document.getElementById('fetch-status');
            btn.disabled = true;
            btn.style.opacity = '0.6';
            status.textContent = 'Triggering report update...';

            fetch('https://cohort-update.arianagha.workers.dev', {{ method: 'POST' }})
                .then(r => r.json())
                .then(data => {{
                    if (data.success) {{
                        status.textContent = 'Update triggered! Page will refresh in ~2 minutes.';
                        status.style.color = '#2e7d32';
                        setTimeout(() => window.location.reload(), 120000);
                    }} else {{
                        status.textContent = 'Error: ' + data.message;
                        status.style.color = '#c62828';
                        btn.disabled = false;
                        btn.style.opacity = '1';
                    }}
                }})
                .catch(err => {{
                    status.textContent = 'Error: ' + err.message;
                    status.style.color = '#c62828';
                    btn.disabled = false;
                    btn.style.opacity = '1';
                }});
        }}
    </script>

    <!-- Date Filter Controls -->
    <div class="date-filter-container">
        <label for="start-date">Start Date:</label>
        <input type="date" id="start-date" value="{start_date.strftime('%Y-%m-%d')}" min="{min_start_date_str}">

        <label for="end-date">End Date:</label>
        <input type="date" id="end-date" value="{end_date.strftime('%Y-%m-%d')}">

        <button onclick="applyDateFilter()">Apply Filter</button>
    </div>

    <div id="date-range-display">
        Report for Active Cohorts Between <span id="display-start">{start_date.strftime('%B %d, %Y')}</span>
        and <span id="display-end">{end_date.strftime('%B %d, %Y')}</span>
    </div>
    <div id="record-count">Showing <span id="filtered-count">0</span> records</div>

    <div class="viz-row">
        <div class="viz-wrapper">
            <h3>Map of Participating Sites</h3>
            <div class="viz-container">
                <div id="map"></div>
            </div>
            <button class="download-btn" onclick="downloadMap()">Download Map as PNG</button>
        </div>
        <div class="viz-wrapper">
            <h3>Timelines for Cohorts</h3>
            <div class="viz-container">
                <div id="timeline"></div>
            </div>
            <button class="download-btn" onclick="downloadTimeline()">Download Timeline as PNG</button>
        </div>
    </div>

    <h2>Number of Sites and Participants Engaged in the Program</h2>
    <div class="table-container">
        <table class="table" id="summary-table">
            <thead></thead>
            <tbody></tbody>
        </table>
    </div>

    <h2>Completed Cohorts and Ready Recognition Status</h2>
    <div class="table-container">
        <table class="table" id="completed-table">
            <thead></thead>
            <tbody></tbody>
        </table>
    </div>

    <script>
        // Embedded data
        const rawData = {data_json};
        const colorMapping = {color_mapping_json};

        // Map instance
        let map = null;
        let markersLayer = null;

        // Initialize on page load
        document.addEventListener('DOMContentLoaded', function() {{
            initializeMap();
            applyDateFilter();
        }});

        function initializeMap() {{
            map = L.map('map', {{ backgroundColor: '#ffffff' }}).setView([36.1, -95.7], 4);
            // Use CartoDB Positron (no labels): white oceans, light gray land
            L.tileLayer('https://{{s}}.basemaps.cartocdn.com/light_nolabels/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                attribution: '&copy; OpenStreetMap contributors &copy; CARTO',
                maxZoom: 19
            }}).addTo(map);

            markersLayer = L.layerGroup().addTo(map);

            // Add legend
            const legend = L.control({{position: 'bottomleft'}});
            legend.onAdd = function(map) {{
                const div = L.DomUtil.create('div', 'legend');
                div.innerHTML = '<div class="legend-title">Office</div>';
                for (const [office, color] of Object.entries(colorMapping)) {{
                    div.innerHTML += `<div class="legend-item"><div class="legend-color" style="background:${{color}}"></div>${{office}}</div>`;
                }}
                return div;
            }};
            legend.addTo(map);
        }}

        function parseDate(dateStr) {{
            if (!dateStr) return null;
            return new Date(dateStr);
        }}

        function formatDate(date) {{
            const options = {{ year: 'numeric', month: 'long', day: 'numeric', timeZone: 'UTC' }};
            return date.toLocaleDateString('en-US', options);
        }}

        function formatDateShort(date) {{
            if (!date) return '';
            const d = new Date(date);
            return d.toISOString().split('T')[0];
        }}

        function parseDateInput(dateStr) {{
            // Parse YYYY-MM-DD string as UTC to avoid timezone shifts
            const [year, month, day] = dateStr.split('-').map(Number);
            return new Date(Date.UTC(year, month - 1, day));
        }}

        function applyDateFilter() {{
            const startInput = document.getElementById('start-date').value;
            const endInput = document.getElementById('end-date').value;

            const startDate = parseDateInput(startInput);
            const endDate = parseDateInput(endInput);

            // Update display
            document.getElementById('display-start').textContent = formatDate(startDate);
            document.getElementById('display-end').textContent = formatDate(endDate);

            // Filter data: Planned End Date >= startDate
            const filteredData = rawData.filter(row => {{
                const plannedEnd = parseDate(row['Planned End Date']);
                return plannedEnd && plannedEnd >= startDate;
            }});

            document.getElementById('filtered-count').textContent = filteredData.length;

            // Update all visualizations
            updateMap(filteredData);
            updateTimeline(filteredData);
            updateSummaryTable(filteredData);
            updateCompletedTable(filteredData);
        }}

        function updateMap(data) {{
            markersLayer.clearLayers();

            // Aggregate by location
            const locationMap = {{}};
            data.forEach(row => {{
                const key = `${{row.Office}}|${{row.City}}|${{row.State}}|${{row.COUNTRY}}|${{row.Latitude}}|${{row.Longitude}}`;
                if (!locationMap[key]) {{
                    locationMap[key] = {{
                        office: row.Office,
                        city: row.City,
                        state: row.State,
                        country: row.COUNTRY,
                        lat: row.Latitude,
                        lng: row.Longitude,
                        count: 0
                    }};
                }}
                locationMap[key].count++;
            }});

            // Add markers
            Object.values(locationMap).forEach(loc => {{
                if (loc.lat && loc.lng) {{
                    const color = colorMapping[loc.office] || 'gray';
                    const radius = 5 + loc.count;

                    const marker = L.circleMarker([loc.lat, loc.lng], {{
                        radius: radius,
                        fillColor: color,
                        color: color,
                        weight: 1,
                        opacity: 1,
                        fillOpacity: 0.7
                    }});

                    marker.bindPopup(`
                        <b>Office:</b> ${{loc.office}}<br>
                        <b>City:</b> ${{loc.city}}<br>
                        <b>State:</b> ${{loc.state}}<br>
                        <b>Country:</b> ${{loc.country}}<br>
                        <b>Count:</b> ${{loc.count}}
                    `);

                    markersLayer.addLayer(marker);
                }}
            }});
        }}

        function updateTimeline(data) {{
            // Aggregate by cohort
            const cohortMap = {{}};
            data.forEach(row => {{
                const key = `${{row.Office}}|${{row['Cohort Name']}}|${{row.Coach}}`;
                if (!cohortMap[key]) {{
                    cohortMap[key] = {{
                        office: row.Office,
                        cohortName: row['Cohort Name'],
                        coach: row.Coach,
                        start: row.Start,
                        end: row['Planned End Date']
                    }};
                }} else {{
                    // Get min start and max end
                    if (row.Start && (!cohortMap[key].start || row.Start < cohortMap[key].start)) {{
                        cohortMap[key].start = row.Start;
                    }}
                    if (row['Planned End Date'] && (!cohortMap[key].end || row['Planned End Date'] > cohortMap[key].end)) {{
                        cohortMap[key].end = row['Planned End Date'];
                    }}
                }}
            }});

            // Convert to array and sort
            const cohorts = Object.values(cohortMap)
                .filter(c => c.start && c.end)
                .sort((a, b) => {{
                    if (a.office !== b.office) return a.office.localeCompare(b.office);
                    return new Date(a.start) - new Date(b.start);
                }});

            // Create Plotly timeline data
            const traces = {{}};
            cohorts.forEach(c => {{
                if (!traces[c.office]) {{
                    traces[c.office] = {{
                        x: [],
                        y: [],
                        base: [],
                        type: 'bar',
                        orientation: 'h',
                        name: c.office,
                        marker: {{ color: colorMapping[c.office] }},
                        hovertemplate: []
                    }};
                }}

                const startDate = new Date(c.start);
                const endDate = new Date(c.end);
                const duration = endDate - startDate;

                traces[c.office].base.push(startDate);
                traces[c.office].x.push(duration);
                traces[c.office].y.push(`${{c.office}}: ${{c.cohortName}}`);
                traces[c.office].hovertemplate.push(
                    `<b>${{c.cohortName}}</b><br>` +
                    `Office: ${{c.office}}<br>` +
                    `Coach: ${{c.coach}}<br>` +
                    `Start: ${{formatDateShort(c.start)}}<br>` +
                    `End: ${{formatDateShort(c.end)}}<extra></extra>`
                );
            }});

            const plotData = Object.values(traces);

            const layout = {{
                title: 'Cohort Timelines by Office',
                barmode: 'overlay',
                xaxis: {{ type: 'date' }},
                yaxis: {{ automargin: true }},
                legend: {{ title: {{ text: 'Office' }} }},
                margin: {{ l: 200 }},
                autosize: true
            }};

            const config = {{ responsive: true }};

            Plotly.newPlot('timeline', plotData, layout, config);
        }}

        function updateSummaryTable(data) {{
            // Calculate summary statistics
            const summary = {{}};

            data.forEach(row => {{
                const office = row.Office;
                if (!summary[office]) {{
                    summary[office] = {{
                        impl: 0,
                        training: 0,
                        dropout: 0,
                        participants: 0,
                        recognized: 0
                    }};
                }}

                if (row.Implementation === 'Yes') summary[office].impl++;
                else if (row.Implementation === 'No') summary[office].training++;
                else if (row.Implementation === 'Dropout') summary[office].dropout++;

                summary[office].participants += (row.Participants || 0);
                if (row['Date of Recognition']) summary[office].recognized++;
            }});

            // Build table HTML
            let thead = `<tr>
                <th>Office</th>
                <th>Implementation Track Sites</th>
                <th>Training Only Sites</th>
                <th>Dropout Sites</th>
                <th>Total Participants</th>
                <th>Recognized Sites</th>
                <th>% of Implementation Track Sites Recognized</th>
            </tr>`;

            let tbody = '';
            let totals = {{ impl: 0, training: 0, dropout: 0, participants: 0, recognized: 0 }};

            Object.keys(summary).sort().forEach(office => {{
                const s = summary[office];
                const pct = s.impl > 0 ? Math.round((s.recognized / s.impl) * 100) : 0;
                tbody += `<tr>
                    <td>${{office}}</td>
                    <td>${{s.impl}}</td>
                    <td>${{s.training}}</td>
                    <td>${{s.dropout}}</td>
                    <td>${{s.participants}}</td>
                    <td>${{s.recognized}}</td>
                    <td>${{pct}}%</td>
                </tr>`;

                totals.impl += s.impl;
                totals.training += s.training;
                totals.dropout += s.dropout;
                totals.participants += s.participants;
                totals.recognized += s.recognized;
            }});

            const totalPct = totals.impl > 0 ? Math.round((totals.recognized / totals.impl) * 100) : 0;
            tbody += `<tr style="font-weight: bold;">
                <td>Total</td>
                <td>${{totals.impl}}</td>
                <td>${{totals.training}}</td>
                <td>${{totals.dropout}}</td>
                <td>${{totals.participants}}</td>
                <td>${{totals.recognized}}</td>
                <td>${{totalPct}}%</td>
            </tr>`;

            document.querySelector('#summary-table thead').innerHTML = thead;
            document.querySelector('#summary-table tbody').innerHTML = tbody;
        }}

        function updateCompletedTable(data) {{
            // Group by cohort with actual end date
            const cohortMap = {{}};
            data.forEach(row => {{
                if (!row['Actual End Date']) return;

                const key = `${{row.Office}}|${{row['Cohort Name']}}|${{row.Coach}}|${{row['Actual End Date']}}`;
                if (!cohortMap[key]) {{
                    cohortMap[key] = {{
                        office: row.Office,
                        cohortName: row['Cohort Name'],
                        coach: row.Coach,
                        endDate: row['Actual End Date'],
                        completed: 0,
                        recognition: {{}}
                    }};
                }}

                if (row.Completed === 'Yes') cohortMap[key].completed++;

                const likelihood = row['Likelihood of Recognition'] || 'Unknown';
                cohortMap[key].recognition[likelihood] = (cohortMap[key].recognition[likelihood] || 0) + 1;
            }});

            // Build table
            let thead = `<tr>
                <th>Office</th>
                <th>Cohort Name</th>
                <th>Coach</th>
                <th>Actual End Date</th>
                <th>Completed Program</th>
                <th>Ready Recognition Status</th>
            </tr>`;

            // Sort by end date descending
            const sorted = Object.values(cohortMap).sort((a, b) =>
                new Date(b.endDate) - new Date(a.endDate)
            );

            let tbody = '';
            sorted.forEach(c => {{
                const recognitionStr = Object.entries(c.recognition)
                    .map(([k, v]) => `${{v}} ${{k}}`)
                    .join(', ');

                tbody += `<tr>
                    <td>${{c.office}}</td>
                    <td>${{c.cohortName}}</td>
                    <td>${{c.coach}}</td>
                    <td>${{formatDateShort(c.endDate)}}</td>
                    <td>${{c.completed}}</td>
                    <td>${{recognitionStr}}</td>
                </tr>`;
            }});

            document.querySelector('#completed-table thead').innerHTML = thead;
            document.querySelector('#completed-table tbody').innerHTML = tbody || '<tr><td colspan="6">No completed cohorts in this date range</td></tr>';
        }}

        function downloadMap() {{
            const btn = event.target;
            btn.disabled = true;
            btn.textContent = 'Generating...';

            const mapContainer = document.getElementById('map');

            html2canvas(mapContainer, {{
                useCORS: true,
                allowTaint: true,
                scale: 3,
                backgroundColor: '#ffffff'
            }}).then(canvas => {{
                const link = document.createElement('a');
                link.download = 'cohort_site_map.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
                btn.disabled = false;
                btn.textContent = 'Download Map as PNG';
            }}).catch(err => {{
                alert('Map download failed: ' + err.message);
                btn.disabled = false;
                btn.textContent = 'Download Map as PNG';
            }});
        }}

        function downloadTimeline() {{
            Plotly.downloadImage('timeline', {{
                format: 'png',
                width: 1600,
                height: 900,
                scale: 3,
                filename: 'cohort_timeline'
            }});
        }}
    </script>
</body>
</html>
'''
    return html_template

# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    # Get the latest closed quarter for default dates
    default_start, default_end = get_latest_closed_quarter()

    parser = argparse.ArgumentParser(
        description='Generate EERE 50001 Ready Cohort Program Update HTML Report'
    )
    parser.add_argument(
        '--start', type=str, default=default_start,
        help=f'Start date (YYYY-MM-DD format). Default: {default_start} (latest closed quarter)'
    )
    parser.add_argument(
        '--end', type=str, default=default_end,
        help=f'End date (YYYY-MM-DD format). Default: {default_end} (latest closed quarter)'
    )
    parser.add_argument(
        '--input', type=str, default='../Input/cohorts_data_preprocessed.xlsx',
        help='Path to input Excel file (default: ../Input/cohorts_data_preprocessed.xlsx)'
    )
    parser.add_argument(
        '--output-dir', type=str, default='../Output',
        help='Output directory for generated files'
    )
    parser.add_argument(
        '--interactive', action='store_true',
        help='Generate interactive HTML report with dynamic date filtering'
    )

    args = parser.parse_args()

    # Parse dates
    start_date = datetime.strptime(args.start, '%Y-%m-%d')
    end_date = datetime.strptime(args.end, '%Y-%m-%d')

    # Create output directory if needed
    os.makedirs(args.output_dir, exist_ok=True)

    # Date range string for filenames
    date_str = f"{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}"

    print("=" * 60)
    print("50001 Ready Cohort Program - HTML Report Generator")
    print("=" * 60)
    print(f"Date range: {start_date.strftime('%B %d, %Y')} to {end_date.strftime('%B %d, %Y')}")
    print(f"Input file: {args.input}")
    print(f"Output directory: {args.output_dir}")
    print(f"Report type: {'Interactive' if args.interactive else 'Static'}")
    print("=" * 60)

    # Load data from Excel
    print("\nLoading data from Excel...")
    df = load_data_from_excel(args.input)
    print(f"Loaded {len(df)} records from Excel")

    # Process data
    print("Processing data...")
    df = process_data(df, start_date, end_date)
    print(f"Filtered to {len(df)} records")

    # Create map
    print("\nCreating map...")
    map_path = os.path.join(args.output_dir, f'cohorts_bubble_map_{date_str}.html')
    m = create_map(df, map_path)
    print(f"Map saved to: {map_path}")

    # Read map HTML for embedding
    with open(map_path, 'r') as f:
        map_html = f.read()

    # Create timeline
    print("Creating timeline...")
    timeline_div = create_timeline(df)

    # Create summary table
    print("Creating summary table...")
    summary_df = create_summary_table(df)
    summary_csv_path = os.path.join(args.output_dir, f'summary_data_{date_str}.csv')
    summary_df.to_csv(summary_csv_path, index=False)
    print(f"Summary data saved to: {summary_csv_path}")
    summary_table_html = summary_df.to_html(index=False, classes='table table-striped', justify='center')

    # Create completed cohorts table
    print("Creating completed cohorts table...")
    completed_cohorts_df = create_completed_cohorts_table(df)
    completed_csv_path = os.path.join(args.output_dir, f'completed_cohorts_data_{date_str}.csv')
    completed_cohorts_df.to_csv(completed_csv_path, index=False)
    print(f"Completed cohorts data saved to: {completed_csv_path}")
    completed_cohorts_table_html = completed_cohorts_df.sort_values(
        'Actual End Date', ascending=False
    ).to_html(index=False, classes='table table-striped', justify='center')

    # Create combined HTML report
    if args.interactive:
        # Generate interactive report with dynamic date filtering
        print("\nCreating interactive HTML report...")
        # Load original data again for interactive report (needs all data embedded)
        df_original = load_data_from_excel(args.input)
        html_content = create_interactive_html_report(df_original, start_date, end_date)
        report_path = os.path.join(args.output_dir, f'eere_50001_ready_cohort_program_update_interactive_{date_str}.html')
    else:
        # Generate static report
        print("\nCreating static HTML report...")
        html_content = create_html_report(
            df, start_date, end_date, map_html, timeline_div,
            summary_table_html, completed_cohorts_table_html
        )
        report_path = os.path.join(args.output_dir, f'eere_50001_ready_cohort_program_update_{date_str}.html')

    with open(report_path, 'w') as f:
        f.write(html_content)
    print(f"HTML report saved to: {report_path}")

    print("\n" + "=" * 60)
    print("Report generation complete!")
    print("=" * 60)
    print("\nOutput files:")
    print(f"  - Map:              {map_path}")
    print(f"  - Summary CSV:      {summary_csv_path}")
    print(f"  - Completed CSV:    {completed_csv_path}")
    print(f"  - Combined Report:  {report_path}")
    if args.interactive:
        print("\nNote: Interactive report allows dynamic date filtering in the browser.")


if __name__ == '__main__':
    main()
