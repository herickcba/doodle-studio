/* ============================================================
   Shared helper for the Gemini serverless functions.
   The API key lives ONLY here, server-side, read from the
   GEMINI_API_KEY environment variable — never in the client or repo.
   Files starting with "_" are not routed by Vercel.
   ============================================================ */
'use strict';

const BASE = 'https://generativelanguage.googleapis.com/v1beta/models';

// Model ids are overridable via env so we can fix them without a code change.
const MODELS = {
  nano2: process.env.GEMINI_MODEL_NANO2 || 'gemini-3.1-flash-image',
  pro: process.env.GEMINI_MODEL_PRO || 'gemini-3-pro-image',
  text: process.env.GEMINI_MODEL_TEXT || 'gemini-2.5-flash',
};

function resolveImageModel(key) {
  return key === 'pro' ? MODELS.pro : MODELS.nano2;   // default: nano banana 2
}

function getKey() {
  const k = process.env.GEMINI_API_KEY;
  if (!k) throw new Error('GEMINI_API_KEY não configurada no servidor.');
  return k;
}

async function callGemini(model, body) {
  const resp = await fetch(`${BASE}/${model}:generateContent`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': getKey() },
    body: JSON.stringify(body),
  });
  const text = await resp.text();
  let json;
  try { json = JSON.parse(text); } catch (_) { json = null; }
  if (!resp.ok) {
    const msg = (json && json.error && json.error.message) || text || ('HTTP ' + resp.status);
    const err = new Error(msg);
    err.status = resp.status;
    throw err;
  }
  return json;
}

// Pull the first inline image out of a generateContent response.
function extractImage(json) {
  const parts = (json && json.candidates && json.candidates[0]
    && json.candidates[0].content && json.candidates[0].content.parts) || [];
  for (const p of parts) {
    if (p.inlineData && p.inlineData.data) {
      return { mimeType: p.inlineData.mimeType || 'image/png', data: p.inlineData.data };
    }
  }
  return null;
}

// Pull concatenated text out of a generateContent response.
function extractText(json) {
  const parts = (json && json.candidates && json.candidates[0]
    && json.candidates[0].content && json.candidates[0].content.parts) || [];
  return parts.filter((p) => typeof p.text === 'string').map((p) => p.text).join('');
}

// Minimal JSON body reader (Vercel parses application/json into req.body,
// but be defensive for other setups).
async function readJson(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  if (typeof req.body === 'string' && req.body) {
    try { return JSON.parse(req.body); }
    catch (e) { console.warn('[api] corpo string não-JSON:', e.message); }
  }
  return await new Promise((resolve) => {
    let data = '';
    req.on('data', (c) => { data += c; });
    req.on('end', () => {
      try { resolve(JSON.parse(data || '{}')); }
      catch (e) { console.warn('[api] corpo não-JSON (%d bytes):', data.length, e.message); resolve({}); }
    });
    req.on('error', (e) => { console.warn('[api] erro lendo corpo:', e.message); resolve({}); });
  });
}

function fail(res, status, message) {
  res.status(status).json({ error: message });
}

module.exports = { MODELS, resolveImageModel, callGemini, extractImage, extractText, readJson, fail };
