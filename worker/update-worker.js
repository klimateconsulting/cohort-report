/**
 * Cloudflare Worker — proxies the "Update" button click
 * to trigger the GitHub Actions workflow_dispatch.
 *
 * Environment variables (set as Worker secrets):
 *   GITHUB_PAT  — GitHub Personal Access Token with repo/workflow scope
 */

const GITHUB_OWNER = 'klimateconsulting';
const GITHUB_REPO = 'cohort-report';
const WORKFLOW_FILE = 'update-report.yml';

export default {
  async fetch(request, env) {
    // Handle CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: {
          'Access-Control-Allow-Origin': 'https://cohorts.klimateconsulting.com',
          'Access-Control-Allow-Methods': 'POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type',
        },
      });
    }

    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405 });
    }

    try {
      const githubResponse = await fetch(
        `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/actions/workflows/${WORKFLOW_FILE}/dispatches`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${env.GITHUB_PAT}`,
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': 'cohort-report-worker',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ ref: 'main' }),
        }
      );

      if (githubResponse.status === 204) {
        return new Response(
          JSON.stringify({ success: true, message: 'Report update triggered. Page will refresh in ~2 minutes.' }),
          {
            headers: {
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': 'https://cohorts.klimateconsulting.com',
            },
          }
        );
      } else {
        const errorText = await githubResponse.text();
        return new Response(
          JSON.stringify({ success: false, message: `GitHub API error: ${githubResponse.status}` }),
          {
            status: 502,
            headers: {
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': 'https://cohorts.klimateconsulting.com',
            },
          }
        );
      }
    } catch (err) {
      return new Response(
        JSON.stringify({ success: false, message: err.message }),
        {
          status: 500,
          headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': 'https://cohorts.klimateconsulting.com',
          },
        }
      );
    }
  },
};
