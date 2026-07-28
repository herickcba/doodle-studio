# Extensão (painel Office.js) — fluxo de dados

**Por que este documento existe:** a análise de risco de 21/07/2026 tratou do
"suplemento PowerPoint em VBA", ou seja, apenas a **Faixa**. O CBA Studio tem
três partes que compartilham a mesma versão, e o critério "não enviar imagens
para a internet" — verdadeiro para a Faixa — **não** é verdadeiro para a
Extensão. Este documento registra exatamente o que sai da máquina, quando e para
onde, para que a conclusão da revisão da Faixa não seja generalizada por engano
para o produto inteiro.

**Nenhuma alteração foi feita no código da Extensão na v1.5.0B.** Isto é um
levantamento, não uma correção.

---

## 1. O que sai da máquina

A Extensão faz **duas** coisas na rede. Nada mais.

### 1.1 Geração e edição de imagem (Gemini) — envia conteúdo

Nas abas de imagem, ao usar os botões de gerar/editar, o painel envia para a API
do Gemini (Google):

| Dado enviado | Origem | Quando |
|---|---|---|
| Texto do prompt | Digitado pelo usuário no painel | Gerar, Editar e "Melhorar prompt" |
| Imagem de referência (base64) | Arquivo **escolhido pelo usuário** num seletor de arquivos, ou imagem gerada antes pela própria ferramenta | Só se houver uma referência ativa selecionada |
| Imagem-base (base64) | Imagem da biblioteca do painel que o usuário abriu para editar | Editar |
| Máscara (base64) | Rabisco que o usuário desenhou marcando a região a editar | Editar, se houver rabisco |

**O que não é enviado:** o painel **não varre o deck** nem envia automaticamente
texto de slide, notas, comentários, nome da apresentação, caminhos locais ou
metadados. Todo envio é disparado por um clique explícito em Gerar ou Editar, e
o conteúdo enviado é o que o usuário colocou ali.

Duas ressalvas honestas:

1. Se o usuário escolher no seletor de arquivos uma imagem de cliente, essa
   imagem vai para a API do Google — a ferramenta não tem como distinguir. A
   decisão é do usuário a cada uso, mas não há hoje nenhum aviso na interface
   dizendo que a imagem sai da máquina.
2. Não existe hoje um consentimento explícito nem um interruptor para desligar o
   recurso.

### 1.2 Verificação de versão — não envia conteúdo

`src/taskpane/taskpane.js:292` faz um `GET` a
`https://doodle-studio-sigma.vercel.app/download/version.json` para avisar quando
há versão nova. É uma requisição de leitura, sem corpo, sem parâmetros e sem
identificador do usuário. Nenhum dado da apresentação trafega.

### 1.3 Aba "Otimização"

Roda **inteiramente local** (`src/modules/img-audit.js` não tem nenhuma chamada
de rede). Imagens analisadas ou recomprimidas não saem da máquina.

---

## 2. Como a chave da API é tratada

Há dois caminhos, e a diferença importa:

**Produção (padrão).** O painel chama funções serverless próprias
(`/api/generate-image`, `/api/edit-image`, `/api/improve-prompt`, em
`v3-powerpoint-addin/api/`), que repassam a chamada ao Gemini. A chave vive
**apenas** na variável de ambiente `GEMINI_API_KEY` do projeto na Vercel —
nunca no cliente, nunca no repositório. `api/_lib.js` a lê no servidor e ela não
chega ao navegador. Um scan de segredos roda no diff antes de todo commit
(regra em `CLAUDE.md`).

**Modo de teste (BYO key).** Se o usuário colar a própria chave do Gemini no
painel, ela é guardada em `localStorage` do webview e as chamadas passam a ir
**direto do navegador para o Google**, sem passar pelo nosso servidor
(`src/shared/gemini.js`). Foi feito para testar sem depender de deploy. Duas
consequências a registrar: a chave fica em `localStorage` (não criptografada), e
nesse modo o tráfego não passa por nenhum ponto de controle nosso.

---

## 3. Inventário de destinos externos

| Destino | O que recebe | Disparado por | Escopo |
|---|---|---|---|
| `generativelanguage.googleapis.com` (Google) | Prompt + imagens descritas em §1.1 | Clique em Gerar / Editar / Melhorar prompt | Extensão |
| `doodle-studio-sigma.vercel.app/download/version.json` | Nada (GET) | Abertura do painel | Extensão |
| `appsforoffice.microsoft.com/lib/1/hosted/office.js` | Nada (carrega a biblioteca) | Abertura do painel | Extensão — exigido pela Microsoft |
| — | — | — | **Faixa (VBA): nenhum destino externo** |

---

## 4. Se a CBA quiser aplicar à Extensão o mesmo critério da Faixa

Não recomendo nada aqui — a decisão é de produto, não técnica. As opções, em
ordem de custo:

1. **Aviso + consentimento explícito** antes do primeiro envio, com a frase
   direta ("esta imagem será enviada para a API do Google"). Baixo custo, resolve
   o ponto de transparência.
2. **Interruptor para desligar os recursos de IA**, para decks de clientes com
   cláusula de confidencialidade. Custo médio.
3. **Remover o modo BYO key** em produção, para que todo tráfego passe pelo
   servidor da CBA e possa ser auditado. Custo baixo, mas tira a facilidade de
   teste.
4. **Revisão jurídica dos termos do Gemini** quanto a retenção e uso de conteúdo
   enviado — é a pergunta que o sócio provavelmente vai fazer em seguida, e não é
   uma pergunta de engenharia.

---

Documento relacionado: [`SECURITY-REVIEW-v1.5.0B.md`](SECURITY-REVIEW-v1.5.0B.md)
(revisão da Faixa).
