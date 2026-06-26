/* POST /api/edit-image
   { prompt, model?, baseImageBase64, baseMimeType?, markupImageBase64?, markupMimeType? }
   -> edits the base image guided by the markup overlay + instruction.
   Returns { imageBase64, mimeType }. */
'use strict';
const { resolveImageModel, callGemini, extractImage, readJson, fail } = require('./_lib');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return fail(res, 405, 'Use POST.');
  try {
    const { prompt, model, baseImageBase64, baseMimeType, markupImageBase64, markupMimeType } = await readJson(req);
    if (!baseImageBase64) return fail(res, 400, 'Imagem base ausente.');
    if (!prompt || !String(prompt).trim()) return fail(res, 400, 'Descreva a edição (prompt).');

    const instruction = 'Edite a primeira imagem (base) seguindo a instrução. A segunda imagem (quando houver) '
      + 'são marcações/rabiscos indicando as regiões a alterar — use-as apenas como guia de onde editar, '
      + 'sem incluir os traços no resultado. Mantenha o resto da imagem fiel.\n\nInstrução: ' + String(prompt).trim();

    const parts = [
      { text: instruction },
      { inlineData: { mimeType: baseMimeType || 'image/png', data: baseImageBase64 } },
    ];
    if (markupImageBase64) {
      parts.push({ inlineData: { mimeType: markupMimeType || 'image/png', data: markupImageBase64 } });
    }

    const json = await callGemini(resolveImageModel(model), {
      contents: [{ role: 'user', parts }],
      generationConfig: { responseModalities: ['IMAGE'], imageConfig: { aspectRatio: '16:9' } },
    });

    const img = extractImage(json);
    if (!img) return fail(res, 502, 'O modelo não retornou imagem.');
    return res.status(200).json({ imageBase64: img.data, mimeType: img.mimeType });
  } catch (e) {
    return fail(res, e.status || 500, e.message || 'Erro ao editar a imagem.');
  }
};
