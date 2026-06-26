/* POST /api/improve-prompt
   { prompt, consistency?: string[], hasRef?: bool }
   -> Gemini (text) improves the prompt and returns structured JSON.
   Sent straight to /api/generate-image by the client (no review step). */
'use strict';
const { MODELS, callGemini, extractText, readJson, fail } = require('./_lib');

const CONSISTENCY_HINTS = {
  imagem: 'mantenha forte consistência com a imagem de referência (mesmo conteúdo/identidade visual)',
  estilo: 'preserve o mesmo estilo artístico/render da referência',
  universo: 'mantenha o mesmo universo/mundo, ambientação e regras visuais',
  textura: 'preserve as mesmas texturas, materiais e acabamento de superfície',
  personagem: 'mantenha o mesmo personagem (traços, roupa, proporções) de forma consistente',
  composicao: 'mantenha a mesma composição/enquadramento e disposição dos elementos',
};

const SCHEMA = {
  type: 'object',
  properties: {
    subject: { type: 'string' },
    style: { type: 'string' },
    composition: { type: 'string' },
    lighting: { type: 'string' },
    palette: { type: 'string' },
    mood: { type: 'string' },
    negative: { type: 'string' },
    aspectRatio: { type: 'string' },
    finalPrompt: { type: 'string' },
  },
  required: ['finalPrompt', 'aspectRatio'],
};

module.exports = async (req, res) => {
  if (req.method !== 'POST') return fail(res, 405, 'Use POST.');
  try {
    const { prompt, consistency, hasRef } = await readJson(req);
    if (!prompt || !String(prompt).trim()) return fail(res, 400, 'Prompt vazio.');

    const cons = Array.isArray(consistency) ? consistency.filter((c) => CONSISTENCY_HINTS[c]) : [];
    const consText = cons.length
      ? '\nRestrições de consistência (respeite TODAS): ' + cons.map((c) => '- ' + CONSISTENCY_HINTS[c]).join('; ')
      : '';
    const refText = hasRef ? '\nHá uma imagem de referência anexada na geração; leve-a em conta.' : '';

    const instruction =
      'Você é um diretor de arte. Receba o pedido do usuário e produza um prompt de geração de imagem '
      + 'rico e específico (em inglês, detalhado: assunto, estilo, composição, iluminação, paleta, clima, '
      + 'lente/câmera quando fizer sentido) para um modelo de imagem, em proporção 16:9 (1920x1080). '
      + 'Não invente texto dentro da imagem a menos que pedido. Responda SOMENTE o JSON do schema.'
      + consText + refText
      + '\n\nPedido do usuário: """' + String(prompt).trim() + '"""';

    const json = await callGemini(MODELS.text, {
      contents: [{ role: 'user', parts: [{ text: instruction }] }],
      generationConfig: { responseMimeType: 'application/json', responseSchema: SCHEMA, temperature: 0.7 },
    });

    const out = extractText(json);
    let data;
    try { data = JSON.parse(out); } catch (_) { data = { finalPrompt: String(prompt).trim(), aspectRatio: '16:9' }; }
    if (!data.finalPrompt) data.finalPrompt = String(prompt).trim();
    if (!data.aspectRatio) data.aspectRatio = '16:9';
    return res.status(200).json(data);
  } catch (e) {
    return fail(res, e.status || 500, e.message || 'Erro ao melhorar o prompt.');
  }
};
