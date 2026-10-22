/* global process, setInterval */
// Preloaded with `node --require` by the end-to-end tests so tenant discovery
// never reaches Microsoft. FAKE_FETCH controls the answer:
//   tenant (default)  a work or school tenant (FAKE_TENANT_ID)
//   consumer          the personal Microsoft account tenant
//   missing           a domain Microsoft does not know (HTTP 400)
//   forbid            fail the test process if anything fetches at all
// Each requested URL is appended to FAKE_FETCH_LOG when it is set.
const fs = require('fs');

const CONSUMER = '9188040d-6c67-4c5b-b112-36a304b66dad';

globalThis.fetch = async (url) => {
  if (process.env.FAKE_FETCH_LOG) fs.appendFileSync(process.env.FAKE_FETCH_LOG, url + '\n');
  const mode = process.env.FAKE_FETCH || 'tenant';
  if (mode === 'forbid') throw new Error(`Unexpected fetch of ${url}`);
  if (mode === 'missing') return { ok: false, status: 400, json: async () => ({}) };
  const tenant = mode === 'consumer' ? CONSUMER : process.env.FAKE_TENANT_ID;
  return {
    ok: true,
    status: 200,
    json: async () => ({ issuer: `https://login.microsoftonline.com/${tenant}/v2.0` }),
  };
};

// FAKE_KEEPALIVE holds the event loop open, the way a pending device-code
// sign-in's polling does in the real server.
if (process.env.FAKE_KEEPALIVE) setInterval(() => undefined, 60_000);
