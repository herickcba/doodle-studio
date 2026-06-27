/* POST /api/edit-image
   { prompt, model?, baseImageBase64, baseMimeType?, markupImageBase64?, markupMimeType? }
   -> edits the base image guided by the markup overlay + instruction.
   Returns { imageBase64, mimeType }. */
'use strict';
const { resolveImageModel, callGemini, extractImage, readJson, fail } = require('./_lib');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return fail(res, 405, 'Use POST.');
  try {
    const { prompt, model, baseImageBase64, baseMimeType, maskImageBase64 } = await readJson(req);
    if (!baseImageBase64) return fail(res, 400, 'Imagem base ausente.');
    if (!prompt || !String(prompt).trim()) return fail(res, 400, 'Descreva a edição (prompt).');

    const instruction = maskImageBase64
      ? ('Você é um editor de imagem de precisão (inpainting). A IMAGEM 1 é a base. A IMAGEM 2 é uma MÁSCARA: '
        + 'as áreas em BRANCO indicam EXATAMENTE onde editar; o PRETO deve permanecer pixel-a-pixel idêntico à base. '
        + 'Edite APENAS a região branca da base aplicando a instrução abaixo, combinando perfeitamente iluminação, '
        + 'sombras, perspectiva e estilo com a cena. NÃO gere uma cena nova; NÃO altere fundo, cor de fundo, '
        + 'enquadramento ou composição. NÃO inclua a máscara no resultado.\n\nInstrução: ' + String(prompt).trim())
      : ('Você é um editor de imagem de precisão. EDITE a imagem base mudando SOMENTE: ' + String(prompt).trim()
        + '. Mantenha EXATAMENTE iguais o fundo e sua cor, a iluminação, as sombras, o enquadramento e a composição. '
        + 'NÃO gere uma cena nova.');

    const parts = [
      { text: instruction },
      { inlineData: { mimeType: baseMimeType || 'image/png', data: baseImageBase64 } },
    ];
    if (maskImageBase64) {
      parts.push({ inlineData: { mimeType: 'image/png', data: maskImageBase64 } });
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
