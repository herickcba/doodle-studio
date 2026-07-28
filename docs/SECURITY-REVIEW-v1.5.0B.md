# Revisão de segurança — CBA Studio, Faixa (add-in VBA) — v1.5.0B

**Objeto da revisão:** `BG-DoodleStudio.ppam` (aba "CBA Studio" no PowerPoint).
**Fonte auditado:** `v3-powerpoint-addin/assets/BG-DoodleStudio.bas` — 2.215 linhas,
módulo único `BG_DoodleStudio`.
**Documento de referência:** análise de risco recebida em 21/07/2026
("Revise e atualize tecnicamente o suplemento PowerPoint em VBA…"), 6 seções.
**Data desta revisão:** 27/07/2026.

> **Resumo em uma linha:** o suplemento já cumpria os controles de
> **comportamento** exigidos (não executa comandos do sistema, não acessa a
> rede, não mexe no SO, não carrega código externo). O que faltava era
> **evidência verificável** — e é isso que esta versão entrega. Um item foi
> recusado com justificativa (assinatura digital) e um foi entregue como
> heurística, não como prova (comparação P-code × fonte).

---

## 1. Situação por seção do documento

| Seção do documento | Situação | Como foi resolvido |
|---|---|---|
| 1. Assinatura digital e integridade | **Parcial — assinatura recusada** | Assinatura não implementada (§6). Integridade entregue por SHA-256 + `build-manifest.json` + vínculo ao commit |
| 2. Eliminação do risco de VBA stomping | **Atendido, com uma ressalva** | Fonte versionado é a origem única; `verify-ppam.sh` compara o fonte embutido no binário com o `.bas` do repo. Comparação com o P-code é heurística (§5) |
| 3. Princípio de menor privilégio | **Já atendido** | Nenhuma alteração de comportamento foi necessária; 5 pontos de fronteira documentados abaixo |
| 4. Validação técnica obrigatória | **Atendido** | `tools/vba-static-scan.sh`, automatizado e bloqueante |
| 5. Entregáveis técnicos | **Atendido, menos os de assinatura** | Lista item a item em §7 |
| 6. Critérios de aceite | **17 de 19** | Os 2 não atendidos são os de assinatura digital (§6) |

---

## 2. Resultado da análise estática

Executada por `tools/vba-static-scan.sh` contra os **27 termos listados na seção 4
do documento** mais **19 superfícies adicionais** que a seção 3 proíbe mas a lista
da seção 4 não nomeia (sistema de arquivos, automação de UI, outras bibliotecas
de rede).

Nenhuma ocorrência dos termos abaixo existe no código, em nenhuma forma:

`Shell` · `WScript` · `PowerShell` · `cmd.exe` · `osascript` · `AppleScript` ·
`CreateObject` · `GetObject` · `CallByName` · `Application.Run` · `Declare` ·
`PtrSafe` · `LoadLibrary` · `RegWrite` · `RegRead` · `RegDelete` · `schtasks` ·
`XMLHTTP` · `ServerXMLHTTP` · `WinHTTP` · `ADODB` · `URLDownloadToFile` ·
`VBProject` · `VBComponents` · `CodeModule` · `Eval` · `FileSystemObject` ·
`ScriptControl` · `MSXML2` · `InternetExplorer` · `WinInet` · `SendKeys` ·
`Kill` · `RmDir` · `FileCopy` · `SetAttr` · `ChDir` · `ChDrive` · `SaveAs`

Consequências diretas, que respondem a boa parte da seção 3 do documento:

- **A Faixa não faz nenhuma chamada de rede.** Não existe código capaz de enviar
  texto de slide, imagem, nota, comentário, nome de cliente, caminho local,
  metadado ou área de transferência para fora da máquina. Não é uma promessa de
  processo: não há API disponível no módulo para isso.
- **Não há execução de comando do sistema operacional**, direta ou indireta —
  sem `Shell`, sem `CreateObject`, sem `Declare`/API nativa, sem interpretador.
- **Não há manipulação do próprio projeto VBA** em tempo de execução, nem
  importação ou geração dinâmica de código.
- **Não há alteração do sistema ou do Office**: nada de registro, tarefas
  agendadas, serviços, variáveis de ambiente, Trust Center, Trusted Locations,
  Protected View ou permissões de arquivo.

O scan também procura **strings concatenadas** que possam montar um comando em
tempo de execução (`.exe`, `.ps1`, `.bat`, `.vbs`, `/bin/`, `powershell`,
`osascript`, `cmd /c`). Nenhuma ocorrência.

### Pontos de fronteira (13 ocorrências justificadas)

Estas são as únicas operações do módulo que tocam recursos fora do documento
aberto. Todas são permitidas pelos próprios critérios da seção 3, e cada uma
está registrada em `tools/vba-scan-allowlist.txt` — o build falha se aparecer
qualquer ocorrência nova que não passe por revisão humana.

| # | Onde | O que faz | Por que é aceitável |
|---|---|---|---|
| 1 | `ConfigPath` (l.133, l.137) | `Environ$("HOME")` / `Environ$("APPDATA")` | Monta o caminho do **próprio** `cba-config.txt`. No Mac, `HOME` dentro do Office aponta para a sandbox do PowerPoint. Não enumera nem lê variáveis de ambiente para outro fim |
| 2 | `EnsureConfigDir` (l.144–147) | `Environ$`, `Dir$`, `MkDir` | Só no Windows, só para criar `%APPDATA%\CBAStudio`. Caminho **fixo**, não derivado de entrada do usuário. `Dir$` testa a existência da própria pasta — não pesquisa disco |
| 3 | `LoadConfig` (l.156) | `Open … For Input` | Lê o próprio arquivo de configuração — explicitamente permitido na seção 3 |
| 4 | `WriteConfig` (l.175) | `Open … For Output` | Grava o próprio arquivo de configuração — idem |
| 5 | `ConfigOpen` (l.242) | `ActivePresentation.FollowHyperlink` | Abre a página de Brand Standards. URL **fixa no código**, disparada **só** por clique no botão "Padrões > Abrir página". Sem query string, sem dados da apresentação, sem identificador do usuário, sem abertura automática na carga |
| 6 | `CropPicture` (l.940, l.943), `PasteTextOnly` (l.1033) | `Application.CommandBars.ExecuteMso` | **Não é execução de comando do sistema**: é o mesmo botão da faixa nativa do PowerPoint, chamado pela API do Office. Late-bound (`Dim cb As Object`) apenas porque o Mac não expõe `View.PasteSpecial` nem o recorte de imagem em tempo de compilação |

A allowlist de URLs é chaveada pela **URL literal**, não pelo procedimento: um
endereço novo bloqueia o build mesmo que apareça dentro de um procedimento que
já tem outra URL liberada.

### Entrada externa: a configuração colada pelo usuário

O único dado externo que o módulo aceita é a string de Brand Standards colada em
`ConfigApply` (botão "Padrões > Aplicar config"). Ela é **dados, nunca código**:

- limite de 20.000 caracteres;
- formato `chave=valor` separado por `;`; chaves desconhecidas são **ignoradas**
  em `ApplyKV` — não há caminho para execução;
- cores validadas dígito a dígito em `HexToRGB` (hex inválido é descartado, a cor
  anterior é mantida);
- números lidos com `Val` (ponto decimal, independente do locale pt-BR);
- **qualquer erro de parse faz rollback** para os padrões de fábrica, com aviso
  ao usuário.

Nenhum caminho de arquivo é lido da configuração — o requisito "rejeitar caminhos
fornecidos dentro da configuração" é atendido por construção: o formato não tem
campo de caminho.

---

## 3. Inventário de dependências

Extraído do stream `PROJECT` de dentro do `vbaProject.bin`:

| Dependência | Fornecedor | Versão | Finalidade | Dados acessados | Risco |
|---|---|---|---|---|---|
| Biblioteca de objetos do PowerPoint | Microsoft | a do Office instalado | Ler e formatar formas, slides e texto | Apresentação aberta | Nenhum adicional — é o host |
| Biblioteca do Office (`Application.CommandBars`) | Microsoft | idem | Recortar imagem, colar só texto | Seleção atual | Nenhum adicional |
| VBA runtime (`Environ`, `Open`, `Dir`, `MkDir`, `Format`, `Val`) | Microsoft | idem | Ler/gravar a própria config | Próprio `cba-config.txt` | Nenhum adicional |

**O projeto não declara nenhuma referência externa** — não há linhas
`Reference=` no stream `PROJECT`. Não há DLL, componente COM registrado,
biblioteca de terceiros, `Declare`/`PtrSafe` ou controle ActiveX. Não há
dependência a remover, e não há nada a substituir por recurso nativo: já é tudo
nativo.

---

## 4. Módulos e procedimentos removidos nesta versão

Regra que passou a valer, escrita no cabeçalho do módulo: **todo procedimento
`Public` do módulo é um callback do ribbon.** Um `.ppam` não expõe macros na
caixa de Macros do PowerPoint no Mac, então não existe ponto de entrada além da
faixa — `Public` sem callback correspondente é, por definição, código inalcançável.

| Procedimento | Por que era inalcançável |
|---|---|
| `ApplyStyleGallery` | Callback de uma galeria de estilos que não existe mais no `customUI14.xml` |
| `GetAnchorText` | Callback de um `editBox` de âncora removido da faixa |
| `SetAnchorText` | idem |
| `RoundEverything` | Callback do botão "arredondar a apresentação inteira", retirado quando as operações passaram a agir só na seleção |
| `BG_AplicarEntrelinha` | Fallback para Option+F8 — não funciona em `.ppam` |
| `BG_AplicarEntrelinhaSlides` | idem |

36 linhas removidas (2.251 → 2.215). Nenhum procedimento auxiliar ficou órfão:
`ApplyStyleById`, `ConfirmBigDeck`, `MakeRounded`, `AnchorCm` e
`DoEntrelinhaSelecao/Slides` continuam sendo chamados por callbacks ativos —
verificado antes da remoção e conferido de novo depois.

Não havia módulos de classe (`.cls`), formulários (`.frm`/`.frx`), variáveis
globais sem uso ou referências desnecessárias a remover: o projeto sempre teve
**um módulo só**. `Option Explicit` já estava presente e o projeto compila sem
erro — não havia variável não declarada a corrigir.

---

## 5. Comparação entre fonte e P-code — o que foi feito e o que **não** dá para provar

`tools/verify-ppam.sh` roda depois do build, sobre o `.ppam` final, e faz quatro
checagens. Ele lê o `vbaProject.bin` com um parser próprio de OLE Compound File +
descompressão MS-OVBA escrito só com a biblioteca padrão do Python — **de
propósito**: um gate de publicação não pode depender de `pip install` para rodar
em outra máquina ou daqui a dois anos.

**1. Fonte embutido × fonte versionado.** Extrai o código-fonte de dentro do
binário e compara com o `.bas` do repositório. É a checagem forte: se o binário
publicado não tiver saído do commit informado, o build para.

> Achado técnico: a comparação é feita **ignorando a caixa**. O VBE normaliza a
> caixa dos identificadores ao compilar (VBA é case-insensitive), então o fonte
> extraído difere do `.bas` em algumas dezenas de bytes — na versão auditada,
> 42 bytes, **todos a mesma letra em caixa diferente** (`L`→`l`, `Bold`→`bold`).
> Diferença de conteúdo, nenhuma. Qualquer divergência além de caixa bloqueia.

**2. Callbacks × procedimentos.** Confere que os 41 callbacks declarados no
`customUI14.xml` têm procedimento no fonte, e avisa sobre procedimentos `Public`
sem callback (a regra da seção 4 acima, automatizada).

**3. Identificadores no P-code sem correspondência no fonte.** É a aproximação
possível do que a seção 2 do documento pede. **Não é uma prova**: é uma
heurística que extrai identificadores do P-code compilado e do stream
`_VBA_PROJECT` e reporta os que não aparecem no fonte. Uma comparação semântica
completa entre as duas representações exigiria desmontar o P-code instrução a
instrução — trabalho de perícia, não de gate automatizado. Isso está declarado
aqui em vez de ser vendido como equivalente.

> **Achado relevante — e a razão de o processo de build mudar.**
> Rodando essa checagem no `.ppam` que estava publicado (v1.5.0), apareceram
> **23 identificadores presentes no cache compilado e ausentes do fonte**:
> `C_AZUL`, `C_ROSA`, `C_BRANCON`, `LooseSpacingPts`, `ScaleFontsIn`,
> `FindRefShapes`, `temBackup`, `temRef`, `refSld`, `refFont`, `refShp`,
> `shpIdx`, `sldIdx`, `ppPasteText`, `ppSaveAsOpenXMLPresentationMacroEnabled`,
> `fonteOK`, `geoOK`, `fundotRou`, `ultSig`, `AlignSel`, `Adjust`, `SaveAs`,
> `msoFade`.
>
> Verificado no histórico do Git: **são procedimentos, constantes e variáveis
> reais que existiram no código e foram removidos** (o Page Size antigo, nos
> commits `2eaa637` e `ccb0ebf`). Não é código malicioso oculto — é resíduo da
> tabela de nomes, porque o arquivo hospedeiro `.pptm` vinha sendo reaproveitado
> entre builds.
>
> Ainda assim, **é exatamente o tipo de achado que sustentaria uma acusação de
> stomping numa auditoria de terceiros**: "há procedimentos compilados sem
> código-fonte correspondente". Por isso o processo de build passou a exigir um
> **arquivo-base limpo** a cada release (`tools/BUILD.md`), que é literalmente o
> que a seção 2 do documento pede ("abra um arquivo-base limpo").
>
> **Resultado depois da correção.** Esta versão foi compilada num `.pptm` criado
> do zero. Os 23 resíduos desapareceram; sobraram **5 tokens** (`JVB7BZ`,
> `iJNaWNo`, `sMZfMz`, `Roundo`, `Resumk`), que são bytes binários do próprio
> P-code caindo por acaso na faixa ASCII — mudam a cada build e não correspondem
> a nenhum identificador. É por isso que essa checagem **avisa em vez de
> bloquear**: o que importa é a forma do token, não a contagem.

**4. Manifesto.** Calcula o SHA-256 do arquivo final, gera
`download/build-manifest.json` e confere o hash contra o arquivo. **O build para
se a árvore de trabalho estiver suja** — um manifesto que declara um commit com
o repositório sujo estaria mentindo.

---

## 6. Assinatura digital — decisão e justificativa

**Não implementada.** Registrado aqui, e no `build-manifest.json`
(`"signatureStatus": "not-signed"`), em vez de omitido.

Motivos:

1. **Não é possível no ambiente principal de desenvolvimento.** A assinatura de
   projeto VBA é aplicada pelo diálogo *Tools > Digital Signature* do VBE, que
   **existe apenas no Office para Windows**. O VBE do macOS — onde o add-in é
   compilado — não oferece esse diálogo, e não há ferramenta de linha de comando
   suportada para assinar um `vbaProject.bin`.
2. **Exige aquisição corporativa.** Um certificado de code signing emitido por CA
   pública requer validação da empresa e, desde as regras do CA/Browser Forum de
   2023, chave privada em token de hardware ou HSM — o que também inviabiliza
   assinatura automatizada num pipeline sem infraestrutura dedicada.
3. **Estado anterior.** Confirmado na auditoria: o `.ppam` que estava publicado
   **também não era assinado** (nenhum stream de assinatura no `vbaProject.bin`).
   Esta versão não introduz uma regressão — ela documenta o que já era o caso.

**Controle compensatório entregue:** SHA-256 do arquivo final + manifesto
vinculado ao commit + comparação do fonte embutido com o fonte versionado. Isso
não substitui a garantia de origem que uma assinatura dá ao usuário final no
momento de abrir o arquivo, mas permite que **qualquer pessoa verifique, depois
do fato, que o binário distribuído corresponde ao código auditado**.

**Para habilitar a assinatura no futuro** basta: obter o certificado corporativo,
assinar no VBE de uma máquina Windows **antes** de calcular o hash (a assinatura
altera o arquivo — a ordem correta já está no checklist do `tools/BUILD.md`), e
preencher `certificateSubject` / `certificateThumbprint` no manifesto.

---

## 7. Entregáveis (seção 5 do documento)

| Pedido | Entregue |
|---|---|
| Código-fonte VBA exportado | `v3-powerpoint-addin/assets/BG-DoodleStudio.bas`. Não há `.cls`/`.frm`/`.frx` — o projeto é de módulo único |
| Estrutura do repositório Git | Já existente; fonte versionado desde o início. `CLAUDE.md` e `HANDOFF.md` descrevem |
| Processo reproduzível de geração do `.ppam` | `tools/BUILD.md` — ver ressalva abaixo |
| Versão recompilada a partir de fonte limpa | Sim, esta versão, a partir de arquivo-base limpo |
| Relatório de comparação fonte × P-code | `tools/verify-ppam.sh` (saída em texto). Heurístico — ver §5 |
| Relatório de análise estática | `tools/vba-static-scan.sh` (saída em texto) |
| Lista de módulos e procedimentos removidos | §4 deste documento |
| Inventário de dependências | §3 deste documento |
| Procedimento de assinatura digital | **Não entregue** — §6 |
| Validação da assinatura | **Não entregue** — §6 |
| Hash SHA-256 do arquivo final | `download/build-manifest.json` |
| `build-manifest.json` | `download/build-manifest.json` |
| Checklist de publicação | `tools/BUILD.md`, seção "Checklist de publicação" |
| Evidência de que o publicado corresponde ao commit | `verify-ppam.sh` passo 1 + `sourceCommit` no manifesto |

**Ressalva sobre "processo reproduzível":** não existe compilador VBA fora do
PowerPoint — no Mac, nenhuma ferramenta compila um `vbaProject.bin` sem abrir o
aplicativo. O passo de compilar no VBE é manual e irredutível. O que o processo
garante não é reprodutibilidade bit a bit, e sim **verificabilidade**: o binário
publicado é comparado com o fonte versionado e vinculado a um commit específico.
Onde o documento pede reprodutibilidade, é isso que está sendo entregue — dito
aqui de forma explícita para não passar por equivalente.

---

## 8. Critérios de aceite (seção 6 do documento)

| Critério | Situação |
|---|---|
| Todos os módulos com `Option Explicit` | ✅ (já era o caso) |
| Projeto gerado a partir do código-fonte versionado | ✅ verificado por `verify-ppam.sh` |
| Processo de build reproduzível | ⚠️ **verificável**, não reproduzível — §7 |
| Código-fonte e P-code compatíveis | ⚠️ sem divergência conhecida; verificação heurística — §5 |
| Sem divergência de compilação não explicada | ✅ as 23 divergências encontradas foram explicadas e a causa, corrigida |
| Sem código ofuscado | ✅ |
| Sem modificação dinâmica do projeto VBA | ✅ |
| Sem execução de comandos do sistema | ✅ |
| Sem alteração de configurações de segurança do Office | ✅ |
| Sem acesso a credenciais | ✅ (nenhum acesso a Keychain, Credential Manager, navegadores ou tokens) |
| Sem leitura arbitrária de arquivos | ✅ só o próprio arquivo de configuração |
| Sem envio de conteúdo da apresentação para a internet | ✅ **na Faixa** — ver a ressalva de escopo abaixo |
| Todas as dependências documentadas | ✅ §3 |
| Arquivo final assinado digitalmente | ❌ §6 |
| Assinatura válida | ❌ §6 |
| Hash SHA-256 registrado | ✅ |
| `build-manifest.json` corresponde ao arquivo final | ✅ conferido pelo próprio script |
| Versão publicada vinculada a um commit | ✅ |

---

## 9. Ressalva de escopo — a Extensão não foi coberta

A análise de risco trata do "suplemento PowerPoint em VBA". O CBA Studio, porém,
tem **três partes** que compartilham a mesma versão: a Faixa (este documento), a
**Extensão** (painel Office.js) e a Landing.

O critério "não enviar imagens para a internet" **é verdadeiro para a Faixa e não
é verdadeiro para a Extensão**: os recursos de gerar e editar imagem enviam a
imagem para a API do Gemini. Isso é função declarada do produto, não um achado
de segurança — mas precisa estar sobre a mesa antes que alguém leia "✅ sem envio
de conteúdo para a internet" e generalize para o produto inteiro.

O fluxo de dados da Extensão está descrito em
[`docs/EXTENSAO-FLUXO-DE-DADOS.md`](EXTENSAO-FLUXO-DE-DADOS.md). Nenhuma
alteração foi feita no código da Extensão nesta versão.
