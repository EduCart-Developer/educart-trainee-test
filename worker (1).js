// Cloudflare Worker — Educart BDM Assessment Proxy
//
// Forwards GET ?action=... and POST { type: ... } requests through to the
// Google Apps Script web app, adding CORS headers so the static frontend can
// call us cross-origin.
//
// After redeploying google-apps-script (3).js as a NEW VERSION, copy the new
// /exec URL here and redeploy this worker. The /exec URL changes per version.

const GAS_URL = 'https://script.google.com/macros/s/AKfycbzuMFjSr5novapEcuQ9WIeAd4CD-HglYIYfQcO83KpwtQXUxvC4qmJ6aMQHsdu7bAcm/exec';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export default {
  async fetch(request) {
    // Handle preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }

    try {
      let gasRes;

      if (request.method === 'GET') {
        // Forward query string to GAS doGet
        const url = new URL(request.url);
        const gasUrl = GAS_URL + url.search;
        gasRes = await fetch(gasUrl, {
          method: 'GET',
          redirect: 'follow',
        });
      } else if (request.method === 'POST') {
        // Forward JSON body to GAS doPost
        const body = await request.text();
        gasRes = await fetch(GAS_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body,
          redirect: 'follow',
        });
      } else {
        return new Response('Method not allowed', { status: 405, headers: CORS });
      }

      const text = await gasRes.text();
      return new Response(text, {
        status: 200,
        headers: { ...CORS, 'Content-Type': 'application/json' },
      });
    } catch (err) {
      return new Response(
        JSON.stringify({ status: 'error', message: err.message }),
        { status: 500, headers: { ...CORS, 'Content-Type': 'application/json' } }
      );
    }
  },
};
