/* POST /api/generate-image
   { prompt, model?: 'nano2'|'pro', refImageBase64?, refMimeType? }
   -> generates a 16:9 image with the chosen Nano Banana model.
   Returns { imageBase64, mimeType }. Client rescales to exactly 1920x1080. */
'use strict';
const { resolveImageModel, callGemini, extractImage, readJson, fail } = require('./_lib');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return fail(res, 405, 'Use POST.');
  try {
    const { prompt, model, refImageBase64, refMimeType } = await readJson(req);
    if (!prompt || !String(prompt).trim()) return fail(res, 400, 'Prompt vazio.');

    const parts = [{ text: String(prompt) }];
    if (refImageBase64) {
      parts.push({ inlineData: { mimeType: refMimeType || 'image/png', data: refImageBase64 } });
    }

    const json = await callGemini(resolveImageModel(model), {
      contents: [{ role: 'user', parts }],
      generationConfig: { responseModalities: ['IMAGE'], imageConfig: { aspectRatio: '16:9' } },
    });

    const img = extractImage(json);
    if (!img) return fail(res, 502, 'O modelo não retornou imagem.');
    return res.status(200).json({ imageBase64: img.data, mimeType: img.mimeType });
  } catch (e) {
    return fail(res, e.status || 500, e.message || 'Erro ao gerar a imagem.');
  }
};
