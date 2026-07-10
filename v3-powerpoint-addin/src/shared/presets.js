/* ============================================================
   Doodles B+G — presets embarcados no app, iguais pra todo mundo.
   Aparecem no topo da biblioteca do painel (seção "Doodles B+G"),
   sem botão de excluir; o clique insere como qualquer item salvo.

   COMO ADICIONAR UM PRESET:
   1. Desenhe no painel/tela grande e salve na biblioteca.
   2. Passe o mouse no item salvo e clique no botão ⧉ (copiar código).
   3. Cole o JSON copiado como um item deste array (vírgula entre itens).
   4. Deploy do projeto da extensão (vercel --prod em v3-powerpoint-addin/).

   Formato de cada item (o ⧉ já copia exatamente assim):
     { "name": "Círculo", "kind": "png"|"gif",
       "gif": {duration,fps,holdMs,easing,loop}?,   // só quando kind="gif"
       "payload": { insertSeparate, strokes:[...], config:{...} } }
   O thumbnail é gerado em tempo real a partir dos traços — não é salvo.
   ============================================================ */
window.DoodlePresets = [
  // Cole os presets aqui.
];
