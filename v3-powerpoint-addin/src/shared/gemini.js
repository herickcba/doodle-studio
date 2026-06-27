/* ============================================================
   Client-side Gemini (BYO key) — TEST MODE.
   When the user pastes their own key in the panel, the add-in calls
   the Gemini API directly from the browser with that key (stored in
   localStorage). This avoids any server setup while testing. For
   production, leave the key empty and the module uses our serverless
   functions (key stays server-side). See image-gen.js dispatch.
   ============================================================ */
(function (global) {
  'use strict';
  const BASE = 'https://generativelanguage.googleapis.com/v1beta/models';
  const KEY_LS = 'doodle.geminiKey';

  const MODELS = {
    nano2: 'gemini-3.1-flash-image',
    pro: 'gemini-3-pro-image',
    text: 'gemini-2.5-flash',
  };
  function imageModel(m) { return m === 'pro' ? MODELS.pro : MODELS.nano2; }

  function getKey() { try { return localStorage.getItem(KEY_LS) || ''; } catch (_) { return ''; } }
  function setKey(k) { try { localStorage.setItem(KEY_LS, (k || '').trim()); } catch (_) {} }
  function clearKey() { try { localStorage.removeItem(KEY_LS); } catch (_) {} }

  async function call(key, model, body) {
    const resp = await fetch(BASE + '/' + model + ':generateContent', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
      body: JSON.stringify(body),
    });
    const text = await resp.text();
    let json = null; try { json = JSON.parse(text); } catch (_) {}
    if (!resp.ok) throw new Error((json && json.error && json.error.message) || text || ('HTTP ' + resp.status));
    return json;
  }
  function firstImage(json) {
    const parts = (json && json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts) || [];
    for (const p of parts) if (p.inlineData && p.inlineData.data) return { imageBase64: p.inlineData.data, mimeType: p.inlineData.mimeType || 'image/png' };
    return null;
  }
  function allText(json) {
    const parts = (json && json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts) || [];
    return parts.filter((p) => typeof p.text === 'string').map((p) => p.text).join('');
  }

  const CONS_HINTS = {
    imagem: 'mantenha forte consistência com a imagem de referência',
    estilo: 'preserve o mesmo estilo artístico/render da referência',
    universo: 'mantenha o mesmo universo/ambientação e regras visuais',
    textura: 'preserve as mesmas texturas e materiais',
    personagem: 'mantenha o mesmo personagem de forma consistente',
    composicao: 'mantenha a mesma composição/enquadramento',
  };
  const SCHEMA = {
    type: 'object',
    properties: {
      subject: { type: 'string' }, style: { type: 'string' }, composition: { type: 'string' },
      lighting: { type: 'string' }, palette: { type: 'string' }, mood: { type: 'string' },
      negative: { type: 'string' }, aspectRatio: { type: 'string' }, finalPrompt: { type: 'string' },
    },
    required: ['finalPrompt', 'aspectRatio'],
  };

  async function improvePrompt(key, { prompt, consistency, hasRef }) {
    const cons = (consistency || []).filter((c) => CONS_HINTS[c]);
    const consText = cons.length ? '\nConsistência (respeite): ' + cons.map((c) => CONS_HINTS[c]).join('; ') : '';
    const refText = hasRef ? '\nHá uma imagem de referência na geração; leve-a em conta.' : '';
    const instruction = 'Você é um diretor de arte. Produza um prompt de geração de imagem rico e específico '
      + '(em inglês: assunto, estilo, composição, iluminação, paleta, clima, lente quando fizer sentido) em 16:9 (1920x1080). '
      + 'Não invente texto na imagem a menos que pedido. Responda SOMENTE o JSON.' + consText + refText
      + '\n\nPedido: """' + String(prompt).trim() + '"""';
    const json = await call(key, MODELS.text, {
      contents: [{ role: 'user', parts: [{ text: instruction }] }],
      generationConfig: { responseMimeType: 'application/json', responseSchema: SCHEMA, temperature: 0.7 },
    });
    let data; try { data = JSON.parse(allText(json)); } catch (_) { data = {}; }
    if (!data.finalPrompt) data.finalPrompt = String(prompt).trim();
    if (!data.aspectRatio) data.aspectRatio = '16:9';
    return data;
  }

  async function generateImage(key, { prompt, model, refImageBase64, refMimeType }) {
    const parts = [{ text: String(prompt) }];
    if (refImageBase64) parts.push({ inlineData: { mimeType: refMimeType || 'image/png', data: refImageBase64 } });
    const json = await call(key, imageModel(model), {
      contents: [{ role: 'user', parts }],
      generationConfig: { responseModalities: ['IMAGE'], imageConfig: { aspectRatio: '16:9' } },
    });
    const img = firstImage(json);
    if (!img) throw new Error('O modelo não retornou imagem.');
    return img;
  }

  async function editImage(key, { prompt, model, baseImageBase64, baseMimeType, markupImageBase64, markupMimeType }) {
    const instruction = 'Edite a primeira imagem (base) seguindo a instrução. A segunda imagem (se houver) tem marcações '
      + 'indicando regiões a alterar — use só como guia, sem incluir os traços no resultado. Mantenha o resto fiel.\n\nInstrução: ' + String(prompt).trim();
    const parts = [{ text: instruction }, { inlineData: { mimeType: baseMimeType || 'image/png', data: baseImageBase64 } }];
    if (markupImageBase64) parts.push({ inlineData: { mimeType: markupMimeType || 'image/png', data: markupImageBase64 } });
    const json = await call(key, imageModel(model), {
      contents: [{ role: 'user', parts }],
      generationConfig: { responseModalities: ['IMAGE'], imageConfig: { aspectRatio: '16:9' } },
    });
    const img = firstImage(json);
    if (!img) throw new Error('O modelo não retornou imagem.');
    return img;
  }

  global.DoodleGemini = { MODELS, getKey, setKey, clearKey, improvePrompt, generateImage, editImage };
})(window);
