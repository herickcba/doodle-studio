Attribute VB_Name = "BG_DoodleStudio"
' ============================================================
'  CBA Studio — Tipografia B+G no ribbon (PowerPoint VBA)
'  Aba "CBA Studio" (customUI14.xml) -> chama os callbacks
'  deste modulo. Cada estilo aplica, num clique:
'    FONTE (Avenir Next) + TAMANHO + NEGRITO + COR
'    + PONTO FINAL colorido + ENTRELINHA exata.
'  (A entrelinha so' da' pra setar por VBA; por isso 100% VBA.)
'
'  Callbacks do ribbon:
'    OnLoad(ribbon)         -> guarda o IRibbonUI
'    ApplyStyle(control)    -> aplica o estilo do botao (control.id)
'    Entrelinha(control)    -> entrelinha B+G na SELECAO
'    EntrelinhaSlides(ctrl) -> entrelinha B+G nos SLIDES SELECIONADOS
'  Tambem expostos sem args p/ Option+F8 (fallback):
'    BG_AplicarEntrelinha / BG_AplicarEntrelinhaSlides
' ============================================================
Option Explicit

' Versao do produto (faixa + extensao + landing andam juntas).
' Ao lancar: atualizar tambem download/version.json, manifest.xml e index.html.
Public Const CBA_VERSION As String = "1.5.0"

' ---- Constantes de layout (fonte unica destes numeros) ----
Private Const PT_PER_CM As Single = 28.3465    ' pontos por cm (72 pt/in / 2,54 cm/in)
Private Const ANCHOR_DEFAULT_CM As Single = 1.27 ' ancora padrao (esq. e topo)
Private Const GUIDE_MARGIN_CM As Single = 3.15 ' margem das linhas-guia
Private Const MAX_DEPTH As Long = 12           ' recursao maxima em grupos aninhados
Private Const STYLE_INSERT As String = "dsH5"  ' estilo das caixas inseridas: Texto 24
Private Const SAMPLE_MAX As Long = 24          ' formas amostradas no Page Size

Private gRibbon As IRibbonUI
Private gAnchorCm As Single   ' ancora esq. (cm); 0 = nao setado -> usa default
Private gAnchorTopCm As Single ' ancora topo (cm); 0 = nao setado -> usa default
Private gRadiusOverridePx As Single   ' raio escolhido no dropdown (px @1080); 0 = Padrao (config)
Private gCfgWarned As Boolean         ' aviso de config quebrada ja' mostrado?
Private Const C_TRANSP As Long = 5         ' indice da cor = transparente (nao e' cor)

' ---- Brand Standards (config). SetDefaults preenche os valores de fabrica;
'      LoadConfig sobrescreve a partir de cba-config.txt (Padroes > Aplicar). ----
'   Paleta: 0 rosa, 1 azul, 2 bege, 3 branco, 4 preto. As cores de TIPO sao
'   pal0 (rosa) e pal1 (azul) — fonte unica de verdade da marca.
Private gFonte As String
Private gPal(0 To 4) As Long
Private gRadiusPx As Single            ' raio padrao em px @ canvas de 1080 de altura
Private gStyle As Object               ' Collection: id -> Array(size, bold, role[0=rosa/1=azul])
Private gEnt As Object                 ' Collection: CStr(size) -> entrelinha (multiplo)

Private Type StyleSpec
    found As Boolean
    size As Single
    bold As Boolean
    color As Long
    periodColor As Long   ' -1 = sem ponto colorido
    bgAware As Boolean
End Type

' ============================================================
'  Callbacks do ribbon
' ============================================================
Public Sub OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
    SetDefaults
    LoadConfig
End Sub

' ============================================================
'  BRAND STANDARDS — defaults, carga e gravacao da config
' ============================================================
Private Sub EnsureCfg()
    If gStyle Is Nothing Then SetDefaults: LoadConfig
    ' Se mesmo assim nao carregou, avisa UMA vez (em vez de falhar mudo botao a botao).
    If gStyle Is Nothing And Not gCfgWarned Then
        gCfgWarned = True
        MsgBox "Nao consegui carregar os padroes da marca - os botoes de estilo " & _
               "podem nao funcionar. Reabra o PowerPoint ou reaplique a config " & _
               "(Padroes > Aplicar config).", vbExclamation, "CBA Studio"
    End If
End Sub

Private Sub SetDefaults()
    gFonte = "Avenir Next"
    gPal(0) = RGB(253, 94, 109)    ' rosa  FD5E6D
    gPal(1) = RGB(67, 106, 225)    ' azul  436AE1
    gPal(2) = RGB(238, 236, 230)   ' bege  EEECE6
    gPal(3) = RGB(255, 255, 255)   ' branco
    gPal(4) = RGB(0, 0, 0)         ' preto
    gRadiusPx = 25
    Set gStyle = New Collection
    AddStyle "dsHero", 120, True, 0
    AddStyle "dsMega", 80, True, 1
    AddStyle "dsH1", 60, True, 0
    AddStyle "dsLabelSec", 60, True, 0
    AddStyle "dsCorpo", 44, False, 1
    AddStyle "dsH3", 34, True, 0
    AddStyle "dsH4", 28, True, 1
    AddStyle "dsH5", 24, False, 1
    AddStyle "dsCorpoPilar", 20, False, 1
    AddStyle "dsEyebrow", 18, True, 0
    AddStyle "dsCaption", 16, False, 1
    Set gEnt = New Collection
    AddEnt 120, 0.8
    AddEnt 80, 0.9
    AddEnt 60, 0.9
    AddEnt 44, 1.15
    AddEnt 34, 0.95
    AddEnt 28, 1#
    AddEnt 24, 1#
    AddEnt 20, 1.3
    AddEnt 18, 1#
    AddEnt 16, 1.3
End Sub

Private Sub AddStyle(ByVal id As String, ByVal sz As Single, ByVal b As Boolean, ByVal role As Long)
    On Error Resume Next
    gStyle.Remove id
    On Error GoTo 0
    gStyle.Add Array(sz, b, role), id
End Sub

Private Sub AddEnt(ByVal sz As Long, ByVal mult As Single)
    On Error Resume Next
    gEnt.Remove CStr(sz)
    On Error GoTo 0
    gEnt.Add mult, CStr(sz)
End Sub

Private Function ConfigPath() As String
    ' Mac: Environ("HOME") dentro do Office = a sandbox do PowerPoint
    ' (~/Library/Containers/com.microsoft.Powerpoint/Data), gravavel por VBA
    ' sem pedir permissao. Windows: HOME e' vazio -> usa %APPDATA%\CBAStudio.
    Dim home As String
    home = Environ$("HOME")
    If Len(home) > 0 Then
        ConfigPath = home & "/cba-config.txt"
    Else
        ConfigPath = Environ$("APPDATA") & "\CBAStudio\cba-config.txt"
    End If
End Function

' Garante a pasta da config no Windows (no Mac o path e' a raiz da sandbox).
Private Sub EnsureConfigDir()
    Dim p As String
    If Len(Environ$("HOME")) > 0 Then Exit Sub          ' Mac: nada a criar
    p = Environ$("APPDATA") & "\CBAStudio"
    On Error Resume Next
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    On Error GoTo 0
End Sub

Private Sub LoadConfig()
    Dim p As String, ln As String, s As String, f As Integer
    p = ConfigPath()
    f = FreeFile
    On Error GoTo done
    Open p For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        s = s & ln
    Loop
    Close #f
    If Len(Trim$(s)) > 0 Then ApplyConfigString s
    Exit Sub
done:
    On Error Resume Next
    Close #f
    On Error GoTo 0
End Sub

Private Function WriteConfig(ByVal s As String) As Boolean
    Dim f As Integer
    On Error GoTo fail
    EnsureConfigDir
    f = FreeFile
    Open ConfigPath() For Output As #f
    Print #f, s
    Close #f
    WriteConfig = True
    Exit Function
fail:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    WriteConfig = False
End Function

' formato: pares "chave=valor" separados por ";". Numeros sempre com ponto
' decimal (parse via Val, independente do locale pt-BR).
Private Sub ApplyConfigString(ByVal s As String)
    Dim parts() As String, kv() As String, i As Long
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    parts = Split(s, ";")
    For i = LBound(parts) To UBound(parts)
        If InStr(parts(i), "=") > 0 Then
            kv = Split(parts(i), "=", 2)   ' limit 2: valor com "=" nao e' truncado
            ApplyKV Trim$(kv(0)), Trim$(kv(1))
        End If
    Next i
End Sub

Private Sub ApplyKV(ByVal k As String, ByVal v As String)
    Dim pi As Long, id As String, a() As String, sz As Long
    If k = "fonte" Then
        If Len(v) > 0 Then gFonte = v
    ElseIf k = "radiusPx" Then
        If Val(v) > 0 Then gRadiusPx = CSng(Val(v))
    ElseIf Left$(k, 3) = "pal" Then
        pi = CLng(Val(Mid$(k, 4)))
        If pi >= 0 And pi <= 4 Then
            Dim c As Long
            c = HexToRGB(v)
            If c >= 0 Then gPal(pi) = c    ' hex invalido (-1) -> ignora, mantem a cor
        End If
    ElseIf Left$(k, 2) = "s_" Then
        id = Mid$(k, 3)
        a = Split(v, "|")
        If UBound(a) >= 2 Then AddStyle id, CSng(Val(a(0))), (Trim$(a(1)) = "1"), CLng(Val(a(2)))
    ElseIf Left$(k, 4) = "ent_" Then
        sz = CLng(Val(Mid$(k, 5)))
        AddEnt sz, CSng(Val(v))
    End If
End Sub

' Retorna a cor RGB, ou -1 se o hex for invalido (o caller ignora).
Private Function HexToRGB(ByVal h As String) As Long
    Dim i As Long, r As Long, g As Long, b As Long
    h = Replace(h, "#", "")
    If Len(h) < 6 Then HexToRGB = -1: Exit Function
    For i = 1 To 6                       ' valida: so digitos hex (GGGGGG crashava)
        If InStr("0123456789ABCDEFabcdef", Mid$(h, i, 1)) = 0 Then HexToRGB = -1: Exit Function
    Next i
    r = CLng("&H" & Mid$(h, 1, 2))
    g = CLng("&H" & Mid$(h, 3, 2))
    b = CLng("&H" & Mid$(h, 5, 2))
    HexToRGB = RGB(r, g, b)
End Function

' Padroes > Abrir pagina: abre a tela de Brand Standards no navegador.
Public Sub ConfigOpen(control As IRibbonControl)
    On Error Resume Next
    ActivePresentation.FollowHyperlink "https://doodle-studio-sigma.vercel.app/config.html"
    On Error GoTo 0
End Sub

' Padroes > Aplicar config: cola a string da pagina (1 linha), grava e recarrega.
' Qualquer erro de parse -> rollback (defaults + config salva) e aviso amigavel.
Public Sub ConfigApply(control As IRibbonControl)
    Dim s As String
    s = InputBox("Cole a configuracao copiada da pagina de Brand Standards:", _
                 "CBA Studio - Aplicar padroes")
    If Len(Trim$(s)) = 0 Then Exit Sub
    If Len(s) > 20000 Then
        MsgBox "Essa configuracao e' grande demais - copie de novo na pagina de Brand Standards.", _
               vbExclamation, "CBA Studio"
        Exit Sub
    End If
    On Error GoTo bad
    SetDefaults
    ApplyConfigString s
    On Error GoTo 0
    If WriteConfig(s) Then
        On Error Resume Next
        gRibbon.Invalidate
        On Error GoTo 0
        MsgBox "Padroes aplicados e salvos.", vbInformation, "CBA Studio"
    Else
        MsgBox "Padroes aplicados nesta sessao, mas nao consegui salvar o arquivo" & vbCrLf & _
               "(permita o acesso a' pasta quando o PowerPoint pedir).", vbExclamation, "CBA Studio"
    End If
    Exit Sub
bad:
    ' rollback: volta pros defaults + config salva anterior (arquivo intacto)
    On Error Resume Next
    SetDefaults
    LoadConfig
    On Error GoTo 0
    MsgBox "Configuracao invalida - nada foi alterado." & vbCrLf & _
           "Copie a string de novo na pagina de Brand Standards.", vbExclamation, "CBA Studio"
End Sub

' Aplica o estilo correspondente ao botao clicado.
Public Sub ApplyStyle(control As IRibbonControl)
    ApplyStyleById control.id
End Sub

' Galeria de estilos: o item clicado traz o id (= chave de SpecFor).
Public Sub ApplyStyleGallery(control As IRibbonControl, id As String, index As Integer)
    ApplyStyleById id
End Sub

Private Sub ApplyStyleById(ByVal styleId As String)
    Dim spec As StyleSpec, sel As Object, shp As Object, n As Long
    spec = SpecFor(styleId)
    If Not spec.found Then Exit Sub
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionText Then
        If sel.TextRange.Length > 0 Then
            ApplyStyleToRange sel.TextRange, spec, False
            n = 1
        Else
            ' so' um cursor (sem selecao): aplica na caixa inteira (com bg-aware)
            On Error Resume Next
            For Each shp In sel.ShapeRange
                n = n + ApplyStyleToShape(shp, spec)
            Next shp
            On Error GoTo 0
        End If
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            n = n + ApplyStyleToShape(shp, spec)
        Next shp
    End If
    If n = 0 Then MsgBox "Selecione um texto ou um objeto com texto.", vbInformation, "CBA Studio"
End Sub

Public Sub Entrelinha(control As IRibbonControl)
    DoEntrelinhaSelecao
End Sub

Public Sub EntrelinhaSlides(control As IRibbonControl)
    DoEntrelinhaSlides
End Sub

' ============================================================
'  Alvo das operacoes de PADRONIZAR (tipografia, entrelinha, cores)
'  Regra unica: os SLIDES selecionados no painel; se a selecao for de
'  objetos, o slide deles; sem nada selecionado, o slide atual.
'  NUNCA a apresentacao inteira — padronizar tudo de uma vez era
'  destrutivo demais em decks reais (e sem undo confiavel).
' ============================================================
Private Function TargetSlides() As Collection
    Dim col As Collection, sel As Object, sld As Object
    Set col = New Collection
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.Type = ppSelectionSlides Then
        For Each sld In sel.SlideRange
            col.Add sld
        Next sld
    ElseIf sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        Set sld = sel.ShapeRange(1).Parent
        If Not sld Is Nothing Then col.Add sld
    End If
    If col.Count = 0 Then
        Set sld = CurrentSlide()
        If Not sld Is Nothing Then col.Add sld
    End If
    On Error GoTo 0
    Set TargetSlides = col
End Function

' Sufixo padrao das mensagens: "em N objeto(s), em M slide(s)."
Private Function ScopeMsg(ByVal n As Long, ByVal nSlides As Long) As String
    ScopeMsg = " em " & n & " objeto(s), em " & nSlides & " slide(s)."
End Function

' Operacoes que varrem a apresentacao inteira podem demorar em decks grandes
' (o PowerPoint nao expoe ScreenUpdating no VBA): confirma antes de rodar.
Private Function ConfirmBigDeck(ByVal opName As String) As Boolean
    Dim n As Long
    ConfirmBigDeck = True
    On Error Resume Next
    n = ActivePresentation.Slides.Count
    On Error GoTo 0
    If n > 50 Then
        ConfirmBigDeck = (MsgBox("Isso vai " & opName & " em " & n & " slides e pode levar um tempo." & _
            vbCrLf & "Continuar?", vbQuestion + vbOKCancel, "CBA Studio") = vbOK)
    End If
End Function

' Entrelinha FIXA por step (botoes 0,8x ... 1,3x) aplicada a' selecao.
Public Sub SetEntrelinha(control As IRibbonControl)
    Dim mult As Single, sel As Object, shp As Object, n As Long
    mult = StepMult(control.id)
    If mult <= 0 Then Exit Sub
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionText Then
        If sel.TextRange.Length > 0 Then
            SetSpacingOnRange sel.TextRange, mult
            n = 1
        Else
            On Error Resume Next
            For Each shp In sel.ShapeRange
                n = n + SetSpacingOnShape(shp, mult)
            Next shp
            On Error GoTo 0
        End If
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            n = n + SetSpacingOnShape(shp, mult)
        Next shp
    End If
    If n = 0 Then MsgBox "Selecione um texto ou um objeto com texto.", vbInformation, "CBA Studio"
End Sub

Private Function StepMult(ByVal id As String) As Single
    Select Case id
        Case "ent08":  StepMult = 0.8
        Case "ent09":  StepMult = 0.9
        Case "ent095": StepMult = 0.95
        Case "ent10":  StepMult = 1#
        Case "ent115": StepMult = 1.15
        Case "ent13":  StepMult = 1.3
        Case Else:     StepMult = 0
    End Select
End Function

Private Sub SetSpacingOnRange(ByVal tr As Object, ByVal mult As Single)
    Dim p As Long, para As Object
    For p = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(p, 1)
        para.ParagraphFormat.LineRuleWithin = msoTrue
        para.ParagraphFormat.SpaceWithin = mult
    Next p
End Sub

Private Function SetSpacingOnShape(ByVal shp As Object, ByVal mult As Single, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + SetSpacingOnShape(s, mult, depth + 1)
        Next s
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            SetSpacingOnRange shp.TextFrame.TextRange, mult
            cnt = 1
        End If
    End If
    On Error GoTo 0
    SetSpacingOnShape = cnt
End Function

' Fallback p/ Option+F8 (subs sem argumentos aparecem na caixa de Macros)
Public Sub BG_AplicarEntrelinha()
    DoEntrelinhaSelecao
End Sub
Public Sub BG_AplicarEntrelinhaSlides()
    DoEntrelinhaSlides
End Sub

' ============================================================
'  ALINHAR a's ancoras ESQUERDA e TOPO (medida via dropdown, em cm)
'  Nucleo compartilhado; os callbacks publicos so' delegam.
'  (gAnchorCm / gAnchorTopCm declaradas no topo do modulo)
' ============================================================
Private Function AnchorCm() As Single
    If gAnchorCm <= 0 Then AnchorCm = ANCHOR_DEFAULT_CM Else AnchorCm = gAnchorCm
End Function

Private Function AnchorTopCm() As Single
    If gAnchorTopCm <= 0 Then AnchorTopCm = ANCHOR_DEFAULT_CM Else AnchorTopCm = gAnchorTopCm
End Function

' "Medida: 1,27 cm" — Format "0.##" no locale pt-BR pode deixar um separador
' solto ("3," em vez de "3"); normaliza pra virgula e tira o pendurado.
Private Function AnchorLabelFor(ByVal cm As Single) As String
    Dim t As String
    t = Replace(Format(cm, "0.##"), ".", ",")
    Do While Len(t) > 0 And Right$(t, 1) = ","
        t = Left$(t, Len(t) - 1)
    Loop
    AnchorLabelFor = "Medida: " & t & " cm"
End Function

' Ancora a borda (esquerda ou topo) dos OBJETOS SELECIONADOS a' medida em cm.
Private Sub AnchorAlign(ByVal cm As Single, ByVal topAxis As Boolean)
    Dim sel As Object, shp As Object, n As Long, p As Single
    Set sel = ActiveWindow.Selection
    p = cm * PT_PER_CM
    n = 0
    If sel.Type = ppSelectionShapes Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            If topAxis Then shp.Top = p Else shp.Left = p
            n = n + 1
        Next shp
        On Error GoTo 0
    End If
    If n = 0 Then MsgBox "Selecione um ou mais objetos para ancorar.", vbInformation, "CBA Studio"
End Sub

Public Sub GetAnchorText(control As IRibbonControl, ByRef text As Variant)
    text = Format(AnchorCm, "0.0#")
End Sub

Public Sub SetAnchorText(control As IRibbonControl, text As String)
    Dim v As Single
    v = Val(Replace(text, ",", "."))
    If v > 0 Then gAnchorCm = v
End Sub

' Dropdown de medida (gallery). Label mostra o valor atual; cada item seta
' o global (id "ancD" = padrao; "anc150" -> 1,5 cm; "ancT150" -> topo 1,5 cm).
Public Sub GetAnchorLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = AnchorLabelFor(AnchorCm)
End Sub

Public Sub GetAnchorTopLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = AnchorLabelFor(AnchorTopCm)
End Sub

Public Sub SetAnchorPick(control As IRibbonControl, id As String, index As Integer)
    If id = "ancD" Then
        gAnchorCm = 0                       ' 0 -> AnchorCm() devolve o padrao
    Else
        gAnchorCm = CSng(Val(Mid$(id, 4))) / 100#   ' "anc150" -> 1,5 cm
    End If
    On Error Resume Next
    gRibbon.InvalidateControl "anchorPick"
    On Error GoTo 0
End Sub

Public Sub SetAnchorTopPick(control As IRibbonControl, id As String, index As Integer)
    If id = "ancTD" Then
        gAnchorTopCm = 0                    ' 0 -> AnchorTopCm() devolve o padrao
    Else
        gAnchorTopCm = CSng(Val(Mid$(id, 5))) / 100#   ' "ancT150" -> 1,5 cm
    End If
    On Error Resume Next
    gRibbon.InvalidateControl "anchorTopPick"
    On Error GoTo 0
End Sub

Public Sub AlignAnchorLeft(control As IRibbonControl)
    AnchorAlign AnchorCm, False
End Sub

Public Sub AlignAnchorTop(control As IRibbonControl)
    AnchorAlign AnchorTopCm, True
End Sub

' ============================================================
'  LINHAS-GUIA (toggle CROSS-SLIDE): margem de 3,15 cm em todos os
'  lados + 4 colunas com gutter de 2x o raio padrao, em TODOS os
'  slides. A cor adapta ao fundo de cada slide: fundo claro -> vermelha;
'  escuro/colorido -> branca. Shapes marcadas "CBAGuide*".
' ============================================================
Public Sub ToggleGuides(control As IRibbonControl, pressed As Boolean)
    Dim sld As Object
    On Error Resume Next
    If ActivePresentation.Slides.Count = 0 Then Exit Sub
    If PresHasGuides() Then
        For Each sld In ActivePresentation.Slides
            RemoveGuides sld
        Next sld
    Else
        For Each sld In ActivePresentation.Slides
            DrawGuides sld
        Next sld
    End If
    gRibbon.InvalidateControl "guidesToggle"
    On Error GoTo 0
End Sub

Public Sub GetGuidesPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = PresHasGuides()
End Sub

' Ha' guias no deck? (checa o 1o slide — sao aplicadas em todos de uma vez)
Private Function PresHasGuides() As Boolean
    PresHasGuides = False
    On Error Resume Next
    If ActivePresentation.Slides.Count >= 1 Then
        PresHasGuides = SlideHasGuides(ActivePresentation.Slides(1))
    End If
    On Error GoTo 0
End Function

Private Function SlideHasGuides(ByVal sld As Object) As Boolean
    Dim s As Object
    SlideHasGuides = False
    On Error Resume Next
    For Each s In sld.Shapes
        If Left$(s.Name, 8) = "CBAGuide" Then SlideHasGuides = True: Exit Function
    Next s
    On Error GoTo 0
End Function

Private Sub RemoveGuides(ByVal sld As Object)
    Dim i As Long
    On Error Resume Next
    For i = sld.Shapes.Count To 1 Step -1
        If Left$(sld.Shapes(i).Name, 8) = "CBAGuide" Then sld.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub DrawGuides(ByVal sld As Object)
    Dim mpt As Single, sw As Single, sh As Single, Wi As Single, Hi As Single
    Dim g As Single, c As Single, i As Long, x As Single
    Dim col As Long, wpt As Single
    Dim box As Object, ln As Object
    Dim names() As String, cnt As Long, rng As Object, grp As Object
    Dim xs(0 To 5) As Single

    EnsureCfg
    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    mpt = GUIDE_MARGIN_CM * PT_PER_CM             ' 3,15 cm -> pontos
    Wi = sw - 2 * mpt
    Hi = sh - 2 * mpt
    If Wi <= 0 Or Hi <= 0 Then Exit Sub
    g = 2 * (CDbl(gRadiusPx) / 1080#) * sh        ' gutter = 2x o raio padrao (config)
    c = (Wi - 3 * g) / 4                          ' largura de cada coluna
    col = GuideColor(sld)                         ' vermelha (fundo claro) ou branca (escuro/colorido)
    wpt = 0.75

    ReDim names(0 To 6)
    cnt = 0

    ' moldura (margem de 3,15 cm)
    On Error Resume Next
    Set box = sld.Shapes.AddShape(msoShapeRectangle, mpt, mpt, Wi, Hi)
    box.Fill.Visible = msoFalse
    box.Line.ForeColor.RGB = col
    box.Line.Weight = wpt
    box.Line.DashStyle = msoLineDash
    box.Name = "CBAGuideBox"
    names(cnt) = box.Name: cnt = cnt + 1

    ' 6 linhas internas (bordas dos gutters entre as 4 colunas)
    If c > 0 Then
        xs(0) = mpt + c
        xs(1) = mpt + c + g
        xs(2) = mpt + 2 * c + g
        xs(3) = mpt + 2 * c + 2 * g
        xs(4) = mpt + 3 * c + 2 * g
        xs(5) = mpt + 3 * c + 3 * g
        For i = 0 To 5
            x = xs(i)
            Set ln = sld.Shapes.AddLine(x, mpt, x, sh - mpt)
            ln.Line.ForeColor.RGB = col
            ln.Line.Weight = wpt
            ln.Line.DashStyle = msoLineDash
            ln.Name = "CBAGuideCol" & i
            names(cnt) = ln.Name: cnt = cnt + 1
        Next i
    End If

    ' agrupa (se der) num unico objeto de guia — evita mexer numa linha por vez
    If cnt > 1 Then
        ReDim Preserve names(0 To cnt - 1)
        Set rng = sld.Shapes.Range(names)
        Set grp = rng.Group
        grp.Name = "CBAGuides"
    End If
    On Error GoTo 0
End Sub

' Cor da guia conforme o fundo do slide: claro e neutro -> vermelha;
' escuro OU colorido (saturado) -> branca. Desconhecido -> vermelha.
Private Function GuideColor(ByVal sld As Object) As Long
    Dim rgbBg As Long, r As Long, g As Long, b As Long
    Dim mx As Long, mn As Long, L As Double, chroma As Double
    rgbBg = SlideBgColor(sld)
    If rgbBg < 0 Then GuideColor = RGB(237, 28, 36): Exit Function
    r = rgbBg And &HFF
    g = (rgbBg \ &H100) And &HFF
    b = (rgbBg \ &H10000) And &HFF
    mx = r: If g > mx Then mx = g
    If b > mx Then mx = b
    mn = r: If g < mn Then mn = g
    If b < mn Then mn = b
    L = (0.299 * r + 0.587 * g + 0.114 * b) / 255#
    chroma = (mx - mn) / 255#
    If L >= 0.6 And chroma < 0.18 Then
        GuideColor = RGB(237, 28, 36)     ' fundo claro e neutro -> vermelho
    Else
        GuideColor = RGB(255, 255, 255)   ' escuro ou colorido -> branco
    End If
End Function

' Cor de fundo efetiva do slide (best-effort). -1 se desconhecida.
'  1) maior forma solida cobrindo o slide (deck "desenhado"); 2) fundo do slide.
Private Function SlideBgColor(ByVal sld As Object) As Long
    Dim c As Long
    c = BigCoverColor(sld)
    If c < 0 Then
        On Error Resume Next
        If sld.Background.Fill.Type = msoFillSolid Then c = sld.Background.Fill.ForeColor.RGB
        On Error GoTo 0
    End If
    SlideBgColor = c
End Function

' Cor da forma solida MAIS ATRAS que cobre ~todo o slide (fundo desenhado).
' Varre do backmost (indice 1) e retorna a 1a que cobre. -1 se nao ha'.
Private Function BigCoverColor(ByVal sld As Object) As Long
    Dim s As Object, i As Long, sw As Single, sh As Single, covers As Boolean
    BigCoverColor = -1
    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    On Error Resume Next
    For i = 1 To sld.Shapes.Count
        Set s = sld.Shapes(i)
        If Left$(s.Name, 8) <> "CBAGuide" Then
            covers = (s.Left <= sw * 0.02) And (s.Top <= sh * 0.02) _
                And (s.Left + s.Width >= sw * 0.98) And (s.Top + s.Height >= sh * 0.98)
            If covers Then
                If s.Fill.Visible = msoTrue And s.Fill.Type = msoFillSolid Then
                    BigCoverColor = s.Fill.ForeColor.RGB
                    Exit Function
                End If
            End If
        End If
        Set s = Nothing
    Next i
    On Error GoTo 0
End Function

' Alinhar/distribuir tradicionais. NAO usamos ShapeRange.Align (no Mac o enum
' MsoAlignCmd vem deslocado: 1 caiu em "centers", 4 em "middles"). Calculamos
' o vertice (menor Left / menor Top) manualmente -> alinhamento exato.
Public Sub AlignLeftEdges(control As IRibbonControl)
    Dim sel As Object, shp As Object, mn As Single, got As Boolean
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Selecione dois ou mais objetos para alinhar.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    On Error Resume Next
    got = False
    For Each shp In sel.ShapeRange
        If Not got Then
            mn = shp.Left
            got = True
        ElseIf shp.Left < mn Then
            mn = shp.Left
        End If
    Next shp
    For Each shp In sel.ShapeRange
        shp.Left = mn
    Next shp
    On Error GoTo 0
End Sub

Public Sub AlignTopEdges(control As IRibbonControl)
    Dim sel As Object, shp As Object, mn As Single, got As Boolean
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Selecione dois ou mais objetos para alinhar.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    On Error Resume Next
    got = False
    For Each shp In sel.ShapeRange
        If Not got Then
            mn = shp.Top
            got = True
        ElseIf shp.Top < mn Then
            mn = shp.Top
        End If
    Next shp
    For Each shp In sel.ShapeRange
        shp.Top = mn
    Next shp
    On Error GoTo 0
End Sub

' Distribuir horizontalmente: mantem o mais a' esquerda e o mais a' direita
' fixos e iguala os espacos (gaps) entre os objetos. Tudo manual (enum deslocado).
Public Sub DistributeH(control As IRibbonControl)
    Dim sel As Object, sr As Object, n As Long, i As Long, j As Long, t As Long
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Selecione tres ou mais objetos para distribuir.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    Set sr = sel.ShapeRange
    n = sr.Count
    If n < 3 Then
        MsgBox "Selecione tres ou mais objetos para distribuir.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    Dim idx() As Long, lefts() As Single, widths() As Single
    ReDim idx(1 To n): ReDim lefts(1 To n): ReDim widths(1 To n)
    On Error Resume Next
    For i = 1 To n
        idx(i) = i
        lefts(i) = sr(i).Left
        widths(i) = sr(i).Width
    Next i
    ' ordena idx por Left (bubble; n pequeno)
    For i = 1 To n - 1
        For j = 1 To n - i
            If lefts(idx(j)) > lefts(idx(j + 1)) Then
                t = idx(j): idx(j) = idx(j + 1): idx(j + 1) = t
            End If
        Next j
    Next i
    Dim left0 As Single, rightN As Single, sumW As Single, gap As Single, cur As Single
    left0 = lefts(idx(1))
    rightN = lefts(idx(n)) + widths(idx(n))
    sumW = 0
    For i = 1 To n
        sumW = sumW + widths(i)
    Next i
    gap = (rightN - left0 - sumW) / (n - 1)
    cur = left0
    For i = 1 To n
        sr(idx(i)).Left = cur
        cur = cur + widths(idx(i)) + gap
    Next i
    On Error GoTo 0
End Sub

Private Function CurrentSlide() As Object
    On Error Resume Next
    Set CurrentSlide = ActiveWindow.View.Slide
    On Error GoTo 0
End Function

' Alinha a' ancora todo shape com texto (grupo com texto = move o grupo inteiro)
Private Function AlignTextShape(ByVal shp As Object, ByVal x As Single) As Long
    Dim cnt As Long
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        If GroupHasText(shp) Then
            shp.Left = x
            cnt = 1
        End If
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            shp.Left = x
            cnt = 1
        End If
    End If
    On Error GoTo 0
    AlignTextShape = cnt
End Function

Private Function GroupHasText(ByVal grp As Object, Optional ByVal depth As Long = 0) As Boolean
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim s As Object
    On Error Resume Next
    For Each s In grp.GroupItems
        If s.Type = msoGroup Then
            If GroupHasText(s, depth + 1) Then GroupHasText = True: Exit Function
        ElseIf s.HasTextFrame Then
            If s.TextFrame.HasText Then GroupHasText = True: Exit Function
        End If
    Next s
    On Error GoTo 0
End Function

' ============================================================
'  INSERIR formas (caixa de texto / rounded box no padrao)
' ============================================================
Public Sub InsertTextBox(control As IRibbonControl)
    Dim sld As Object, tb As Object, x As Single, w As Single
    Set sld = CurrentSlide()
    If sld Is Nothing Then Exit Sub
    x = AnchorCm * PT_PER_CM                     ' margem esquerda da ancora
    w = ActivePresentation.PageSetup.SlideWidth / 3
    Set tb = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x, 80, w, 60)
    On Error Resume Next
    tb.TextFrame.WordWrap = msoTrue
    tb.TextFrame.TextRange.Text = "Texto"
    ApplyInsertStyle tb                          ' ja' nasce no Texto 24 (B+G)
    tb.Select
    On Error GoTo 0
End Sub

' Estilo padrao das caixas que a faixa insere: Texto 24 (dsH5) + entrelinha B+G.
' Escreve no Font do TextRange, que vale tanto pro texto ja' existente quanto
' pro que a pessoa for digitar depois (caso do rounded box, que nasce vazio).
Private Sub ApplyInsertStyle(ByVal shp As Object)
    Dim spec As StyleSpec, ls As Single
    spec = SpecFor(STYLE_INSERT)
    If Not spec.found Then Exit Sub
    On Error Resume Next
    With shp.TextFrame.TextRange.Font
        .Name = gFonte
        .size = spec.size
        .Bold = IIf(spec.bold, msoTrue, msoFalse)
        .Color.RGB = spec.color
    End With
    ls = MapSpacing(spec.size)
    If ls > 0 Then
        With shp.TextFrame.TextRange.ParagraphFormat
            .LineRuleWithin = msoTrue
            .SpaceWithin = ls
        End With
    End If
    On Error GoTo 0
End Sub

Public Sub InsertRoundedBox(control As IRibbonControl)
    Dim sld As Object, shp As Object, sw As Single, sh As Single, w As Single, h As Single
    Set sld = CurrentSlide()
    If sld Is Nothing Then Exit Sub
    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    w = sw / 4
    h = sh / 4
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, (sw - w) / 2, (sh - h) / 2, w, h)
    On Error Resume Next
    shp.Fill.Visible = msoTrue
    shp.Fill.Solid
    If SlideBgIsWhite(sld) Then
        shp.Fill.ForeColor.RGB = PaletteColor(2)   ' bege, se o fundo do slide for branco
    Else
        shp.Fill.ForeColor.RGB = PaletteColor(3)   ' branco
    End If
    shp.Line.Visible = msoFalse
    SetRoundedRadius shp                          ' raio padrao (~25px @1080)
    SetRoundedTextInset shp                       ' padding interno = 2x o raio
    ApplyInsertStyle shp                          ' nasce vazio, mas ja' no Texto 24
    shp.Select
    On Error GoTo 0
End Sub

' Crop de imagem: entra no modo de recorte NATIVO do PowerPoint na imagem
' selecionada (ExecuteMso late-bound, mesmo padrao ja' validado no Colar texto).
Public Sub CropPicture(control As IRibbonControl)
    Dim sel As Object, shp As Object, cb As Object, temImagem As Boolean
    temImagem = False
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            If shp.Type = msoPicture Or shp.Type = msoPlaceholder Then temImagem = True
        Next shp
    End If
    On Error GoTo 0
    If Not temImagem Then
        MsgBox "Selecione uma imagem para recortar.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    On Error Resume Next
    Set cb = Application.CommandBars
    Err.Clear
    cb.ExecuteMso "PictureCrop"
    If Err.Number <> 0 Then
        Err.Clear
        cb.ExecuteMso "Crop"                      ' fallback (varia por build do Mac)
    End If
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Nao consegui abrir o recorte por aqui." & vbCrLf & _
               "Use Formato da Imagem > Recortar.", vbExclamation, "CBA Studio"
        Exit Sub
    End If
    On Error GoTo 0
End Sub

' fundo do slide ~ branco? (best-effort; sem leitura -> False)
Private Function SlideBgIsWhite(ByVal sld As Object) As Boolean
    Dim c As Long
    c = -1
    On Error Resume Next
    c = sld.Background.Fill.ForeColor.RGB
    On Error GoTo 0
    SlideBgIsWhite = (c = RGB(255, 255, 255))
End Function

' padding interno = 1x o raio REAL do shape (Adjustments(1) * menor lado);
' em cards menores o raio real e' menor, entao o padding diminui junto.
Private Sub SetRoundedTextInset(ByVal shp As Object)
    Dim shorter As Single, r As Single
    On Error Resume Next
    If shp.Width < shp.Height Then shorter = shp.Width Else shorter = shp.Height
    r = shp.Adjustments(1) * shorter
    If r < 0 Then r = 0
    shp.TextFrame.MarginLeft = r
    shp.TextFrame.MarginRight = r
    shp.TextFrame.MarginTop = r
    shp.TextFrame.MarginBottom = r
    On Error GoTo 0
End Sub

' ============================================================
'  TEXTO: negrito / colar texto / CAPS+loose / expandir conteudo
' ============================================================
Public Sub ToggleBold(control As IRibbonControl)
    Dim sel As Object, shp As Object, makeBold As Boolean
    Set sel = ActiveWindow.Selection
    On Error Resume Next
    If sel.Type = ppSelectionText Then
        makeBold = (sel.TextRange.Font.Bold <> msoTrue)
        sel.TextRange.Font.Bold = IIf(makeBold, msoTrue, msoFalse)
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    makeBold = (shp.TextFrame.TextRange.Font.Bold <> msoTrue)
                    shp.TextFrame.TextRange.Font.Bold = IIf(makeBold, msoTrue, msoFalse)
                End If
            End If
        Next shp
    End If
    On Error GoTo 0
End Sub

' Alinhamento do PARAGRAFO do texto selecionado (ou de todas as caixas
' selecionadas). Mesmo padrao de selecao do ToggleBold.
Public Sub AlignTextLeft(control As IRibbonControl)
    SetParaAlign ppAlignLeft
End Sub
Public Sub AlignTextCenter(control As IRibbonControl)
    SetParaAlign ppAlignCenter
End Sub

Private Sub SetParaAlign(ByVal al As Long)
    Dim sel As Object, shp As Object
    Set sel = ActiveWindow.Selection
    On Error Resume Next
    If sel.Type = ppSelectionText Then
        sel.TextRange.ParagraphFormat.Alignment = al
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then _
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = al
            End If
        Next shp
    End If
    On Error GoTo 0
End Sub

Public Sub PasteTextOnly(control As IRibbonControl)
    ' Mac: View.PasteSpecial nao existe; ExecuteMso (late-bound p/ evitar compile-check)
    Dim cb As Object
    On Error Resume Next
    Set cb = Application.CommandBars
    cb.ExecuteMso "PasteTextOnly"
    On Error GoTo 0
End Sub

' ALLCAPS + character spacing "loose" (tracking proporcional ao tamanho)
Public Sub CapsLoose(control As IRibbonControl)
    Dim sel As Object, shp As Object, shp2 As Object, tr2 As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    On Error Resume Next
    If sel.Type = ppSelectionText Then
        If sel.TextRange.Length > 0 Then
            Set shp2 = sel.ShapeRange(1)
            Set tr2 = shp2.TextFrame2.TextRange.Characters(sel.TextRange.Start, sel.TextRange.Length)
            ApplyCapsLoose tr2
            n = 1
        End If
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    ApplyCapsLoose shp.TextFrame2.TextRange
                    n = n + 1
                End If
            End If
        Next shp
    End If
    On Error GoTo 0
    If n = 0 Then MsgBox "Selecione um texto.", vbInformation, "CBA Studio"
End Sub

Private Sub ApplyCapsLoose(ByVal tr2 As Object)
    On Error Resume Next
    tr2.Font.Name = gFonte      ' Avenir Next
    tr2.Font.Bold = msoTrue    ' -> Avenir Next Bold
    tr2.Font.Caps = 2          ' msoCapsAll (constante nao definida no Mac)
    tr2.Font.Spacing = 3       ' = "Loose" nativo do PowerPoint (3pt fixos)
    On Error GoTo 0
End Sub

' Expandir conteudo: zera as margens internas das formas selecionadas
Public Sub ExpandContent(control As IRibbonControl)
    Dim sel As Object, shp As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            n = n + ZeroInsets(shp)
        Next shp
        On Error GoTo 0
    End If
    If n = 0 Then MsgBox "Selecione uma ou mais formas.", vbInformation, "CBA Studio"
End Sub

Private Function ZeroInsets(ByVal shp As Object, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + ZeroInsets(s, depth + 1)
        Next s
    ElseIf shp.HasTextFrame Then
        shp.TextFrame.MarginLeft = 0
        shp.TextFrame.MarginRight = 0
        shp.TextFrame.MarginTop = 0
        shp.TextFrame.MarginBottom = 0
        cnt = 1
    End If
    On Error GoTo 0
    ZeroInsets = cnt
End Function

' ============================================================
'  ROUNDED CORNERS (raio ~25px num canvas 1920x1080)
' ============================================================
Private Function StdRadiusPts() As Single
    Dim px As Single
    EnsureCfg
    If gRadiusOverridePx > 0 Then px = gRadiusOverridePx Else px = gRadiusPx
    StdRadiusPts = (CDbl(px) / 1080#) * ActivePresentation.PageSetup.SlideHeight
End Function

' ---- Dropdown "Raio" (grupo Formas): Padrao (config) ou um valor fixo ----
Private Function RadiusLabel() As String
    If gRadiusOverridePx > 0 Then
        RadiusLabel = CStr(CLng(gRadiusOverridePx)) & " px"
    Else
        RadiusLabel = "Padrao"
    End If
End Function

' getLabel do gallery — mostra a escolha atual no proprio botao.
Public Sub GetRadiusLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Raio: " & RadiusLabel()
End Sub

' onAction do gallery — id do item = "radDefault" ou "rad<px>".
Public Sub SetRadiusPick(control As IRibbonControl, id As String, index As Integer)
    If id = "radDefault" Then
        gRadiusOverridePx = 0
    Else
        gRadiusOverridePx = CSng(Val(Mid$(id, 4)))
    End If
    On Error Resume Next
    gRibbon.InvalidateControl "radiusPick"
    On Error GoTo 0
End Sub

Private Function IsRoundedType(ByVal t As Long) As Boolean
    IsRoundedType = (t = msoShapeRoundedRectangle)
End Function

' raio padrao = fracao do menor lado (clamp 0..0.5)
Private Sub SetRoundedRadius(ByVal shp As Object)
    Dim shorter As Single, adj As Single
    On Error Resume Next
    If shp.Width < shp.Height Then shorter = shp.Width Else shorter = shp.Height
    If shorter <= 0 Then Exit Sub
    adj = StdRadiusPts / shorter
    If adj > 0.5 Then adj = 0.5
    If adj < 0 Then adj = 0
    shp.Adjustments(1) = adj
    On Error GoTo 0
End Sub

' transforma o shape em rounded (imagem/caixa/retangulo) e seta o raio
Private Function MakeRounded(ByVal shp As Object, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim done As Long, s As Object, doIt As Boolean
    done = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            done = done + MakeRounded(s, depth + 1)
        Next s
    Else
        doIt = False
        Select Case shp.Type
            Case msoPicture, msoTextBox
                doIt = True
            Case msoAutoShape
                If shp.AutoShapeType = msoShapeRectangle Or IsRoundedType(shp.AutoShapeType) Then doIt = True
        End Select
        If doIt Then
            shp.AutoShapeType = msoShapeRoundedRectangle   ' imagem = crop rounded
            SetRoundedRadius shp
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then SetRoundedTextInset shp
            End If
            done = 1
        End If
    End If
    On Error GoTo 0
    MakeRounded = done
End Function

Public Sub RoundCorner(control As IRibbonControl)
    Dim sel As Object, shp As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            n = n + MakeRounded(shp)
        Next shp
        On Error GoTo 0
    End If
    If n = 0 Then MsgBox "Selecione uma forma ou imagem.", vbInformation, "CBA Studio"
End Sub

' Padroniza o raio SO' das formas que ja' sao rounded, pela SELECAO:
' formas selecionadas, ou os slides selecionados no painel; sem selecao,
' o slide atual. Nunca a apresentacao inteira (mais controle).
Public Sub RoundAllShapes(control As IRibbonControl)
    Dim sel As Object, shp As Object, sld As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionShapes Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            n = n + NormalizeRounded(shp)
        Next shp
        On Error GoTo 0
    ElseIf sel.Type = ppSelectionSlides Then
        On Error Resume Next
        For Each sld In sel.SlideRange
            For Each shp In sld.Shapes
                n = n + NormalizeRounded(shp)
            Next shp
        Next sld
        On Error GoTo 0
    Else
        Set sld = CurrentSlide()
        If Not sld Is Nothing Then
            For Each shp In sld.Shapes
                n = n + NormalizeRounded(shp)
            Next shp
        End If
    End If
    MsgBox "Cantos padronizados em " & n & " forma(s) rounded (raio: " & RadiusLabel() & ").", vbInformation, "CBA Studio"
End Sub

Private Function NormalizeRounded(ByVal shp As Object, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object, t As Long
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + NormalizeRounded(s, depth + 1)
        Next s
    ElseIf shp.Type = msoAutoShape Then
        t = shp.AutoShapeType
        If IsRoundedType(t) Then
            SetRoundedRadius shp
            cnt = 1
        End If
    End If
    On Error GoTo 0
    NormalizeRounded = cnt
End Function

' TUDO rounded: arredonda TODOS os objetos elegiveis (retangulos,
' caixas de texto e imagens) de toda a apresentacao. Formas nao
' retangulares (oval, seta, linha) ficam como estao.
Public Sub RoundEverything(control As IRibbonControl)
    If Not ConfirmBigDeck("arredondar as formas") Then Exit Sub
    Dim sld As Object, shp As Object, n As Long
    n = 0
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            n = n + MakeRounded(shp)
        Next shp
    Next sld
    MsgBox "Arredondados " & n & " objetos na apresentacao.", vbInformation, "CBA Studio"
End Sub

' TIRAR rounded pela SELECAO: objetos selecionados, ou os slides selecionados
' no painel; sem selecao de formas usa o slide atual. Nunca a apresentacao toda.
Public Sub UnroundAll(control As IRibbonControl)
    Dim sel As Object, shp As Object, sld As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionShapes Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            n = n + MakeStraight(shp)
        Next shp
        On Error GoTo 0
    ElseIf sel.Type = ppSelectionSlides Then
        On Error Resume Next
        For Each sld In sel.SlideRange
            For Each shp In sld.Shapes
                n = n + MakeStraight(shp)
            Next shp
        Next sld
        On Error GoTo 0
    Else
        Set sld = CurrentSlide()
        If Not sld Is Nothing Then
            For Each shp In sld.Shapes
                n = n + MakeStraight(shp)
            Next shp
        End If
    End If
    MsgBox "Removido o arredondamento de " & n & " forma(s).", vbInformation, "CBA Studio"
End Sub

Private Function MakeStraight(ByVal shp As Object, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + MakeStraight(s, depth + 1)
        Next s
    ElseIf shp.Type = msoAutoShape Or shp.Type = msoPicture Then
        If IsRoundedType(shp.AutoShapeType) Then
            shp.AutoShapeType = msoShapeRectangle
            cnt = 1
        End If
    End If
    On Error GoTo 0
    MakeStraight = cnt
End Function

' ============================================================
'  Barra de cores (paleta da marca) — preenchimento e fonte
'  indices: 0 rosa, 1 azul, 2 bege, 3 branco, 4 preto
'  control.id = "fill0".."fill4" / "font0".."font4"
' ============================================================
Private Function PaletteColor(ByVal i As Long) As Long
    EnsureCfg
    If i >= 0 And i <= 4 Then PaletteColor = gPal(i) Else PaletteColor = gPal(4)
End Function

Private Function ColorIndexFromId(ByVal id As String) As Long
    Dim s As String
    s = Right$(id, 1)
    If s >= "0" And s <= "9" Then ColorIndexFromId = CLng(s) Else ColorIndexFromId = -1
End Function

Public Sub SetShapeFill(control As IRibbonControl)
    Dim sel As Object, shp As Object, i As Long, n As Long
    i = ColorIndexFromId(control.id)
    If i < 0 Then Exit Sub
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            n = n + FillShape(shp, i)
        Next shp
        On Error GoTo 0
    End If
    If n = 0 Then MsgBox "Selecione uma ou mais formas.", vbInformation, "CBA Studio"
End Sub

Private Function FillShape(ByVal shp As Object, ByVal i As Long, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + FillShape(s, i, depth + 1)
        Next s
    Else
        If i = C_TRANSP Then
            shp.Fill.Visible = msoFalse                ' sem preenchimento
        Else
            shp.Fill.Visible = msoTrue
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = PaletteColor(i)
        End If
        cnt = 1
    End If
    On Error GoTo 0
    FillShape = cnt
End Function

Public Sub SetFontColor(control As IRibbonControl)
    Dim sel As Object, shp As Object, shpSel As Object, tr2 As Object, i As Long, n As Long
    i = ColorIndexFromId(control.id)
    If i < 0 Then Exit Sub
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionText Then
        If sel.TextRange.Length > 0 Then
            ' aplica so' aos caracteres selecionados, via Font2 (cor + opacidade)
            On Error Resume Next
            Set shpSel = sel.ShapeRange(1)
            Set tr2 = shpSel.TextFrame2.TextRange.Characters(sel.TextRange.Start, sel.TextRange.Length)
            ApplyFontColor2 tr2, i
            n = 1
            On Error GoTo 0
        Else
            On Error Resume Next
            For Each shp In sel.ShapeRange
                n = n + FontColorShape(shp, i)
            Next shp
            On Error GoTo 0
        End If
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            n = n + FontColorShape(shp, i)
        Next shp
    End If
    If n = 0 Then MsgBox "Selecione um texto ou forma com texto.", vbInformation, "CBA Studio"
End Sub

' Cor/transparencia da fonte SEMPRE via Font2 (TextFrame2): gerencia cor E
' opacidade no mesmo lugar, entao voltar de transparente p/ uma cor solida
' funciona (zera a transparencia). Era o bug: Font.Color.RGB nao mexia no alpha.
Private Sub ApplyFontColor2(ByVal tr2 As Object, ByVal i As Long)
    On Error Resume Next
    tr2.Font.Fill.Visible = msoTrue
    tr2.Font.Fill.Solid
    If i = C_TRANSP Then
        tr2.Font.Fill.Transparency = 1#
    Else
        tr2.Font.Fill.ForeColor.RGB = PaletteColor(i)
        tr2.Font.Fill.Transparency = 0#
    End If
    On Error GoTo 0
End Sub

Private Function FontColorShape(ByVal shp As Object, ByVal i As Long, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + FontColorShape(s, i, depth + 1)
        Next s
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            ApplyFontColor2 shp.TextFrame2.TextRange, i
            cnt = 1
        End If
    End If
    On Error GoTo 0
    FontColorShape = cnt
End Function

' ============================================================
'  Cor de CONTORNO (outline) — espelha o preenchimento
'  control.id = "outline0".."outline5" (5 = transparente/sem linha)
' ============================================================
Public Sub SetShapeOutline(control As IRibbonControl)
    Dim sel As Object, shp As Object, i As Long, n As Long
    i = ColorIndexFromId(control.id)
    If i < 0 Then Exit Sub
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            n = n + OutlineShape(shp, i)
        Next shp
        On Error GoTo 0
    End If
    If n = 0 Then MsgBox "Selecione uma ou mais formas.", vbInformation, "CBA Studio"
End Sub

Private Function OutlineShape(ByVal shp As Object, ByVal i As Long, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + OutlineShape(s, i, depth + 1)
        Next s
    Else
        If i = C_TRANSP Then
            shp.Line.Visible = msoFalse
        Else
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = PaletteColor(i)
            shp.Line.Weight = 3.5                       ' contorno padrao 3,5pt
        End If
        cnt = 1
    End If
    On Error GoTo 0
    OutlineShape = cnt
End Function

' ============================================================
'  AUDITORIA — fontes / cores / raios fora do padrao
'  Saida: janela-resumo + navegacao alerta-a-alerta (Mac-safe).
'  Paleta de referencia = as 5 cores da marca (PaletteColor 0..4).
' ============================================================
Private Function IsPaletteColor(ByVal rgbVal As Long) As Boolean
    Dim i As Long
    For i = 0 To 4
        If rgbVal = PaletteColor(i) Then IsPaletteColor = True: Exit Function
    Next i
End Function

Private Function NearestPaletteColor(ByVal rgbVal As Long) As Long
    Dim i As Long, best As Long, pc As Long
    Dim r As Long, g As Long, b As Long, pr As Long, pg As Long, pb As Long
    Dim d As Double, bd As Double
    r = rgbVal And &HFF
    g = (rgbVal \ &H100) And &HFF
    b = (rgbVal \ &H10000) And &HFF
    bd = 1E+99
    best = PaletteColor(0)
    For i = 0 To 4
        pc = PaletteColor(i)
        pr = pc And &HFF
        pg = (pc \ &H100) And &HFF
        pb = (pc \ &H10000) And &HFF
        d = CDbl(r - pr) * (r - pr) + CDbl(g - pg) * (g - pg) + CDbl(b - pb) * (b - pb)
        If d < bd Then bd = d: best = pc
    Next i
    NearestPaletteColor = best
End Function

Private Function ShorterSide(ByVal shp As Object) As Single
    If shp.Width < shp.Height Then ShorterSide = shp.Width Else ShorterSide = shp.Height
End Function

' Coleta findings recursivamente. col guarda Array(slideIdx, shape, detalhe).
Private Sub ScanShape(ByVal shp As Object, ByVal slideIdx As Long, ByVal col As Collection, _
                      ByRef nf As Long, ByRef nc As Long, ByRef nr As Long, Optional ByVal depth As Long = 0)
    If depth > MAX_DEPTH Then Exit Sub      ' guarda contra grupos aninhados demais
    Dim s As Object, std As Single, adj As Single, nm As String
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            ScanShape s, slideIdx, col, nf, nc, nr, depth + 1
        Next s
        On Error GoTo 0
        Exit Sub
    End If

    ' fonte (nome + cor)
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            nm = shp.TextFrame.TextRange.Font.Name
            If Len(nm) > 0 And StrComp(nm, gFonte, vbTextCompare) <> 0 Then
                col.Add Array(slideIdx, shp, "Fonte fora do padrao: " & nm)
                nf = nf + 1
            End If
            If Not IsPaletteColor(shp.TextFrame.TextRange.Font.Color.RGB) Then
                col.Add Array(slideIdx, shp, "Cor da fonte fora da paleta")
                nc = nc + 1
            End If
        End If
    End If

    ' raio (so' formas ja' arredondadas)
    If shp.Type = msoAutoShape Then
        If IsRoundedType(shp.AutoShapeType) Then
            If ShorterSide(shp) > 0 Then
                std = StdRadiusPts / ShorterSide(shp)
                If std > 0.5 Then std = 0.5
                adj = shp.Adjustments(1)
                If Abs(adj - std) > 0.02 Then
                    col.Add Array(slideIdx, shp, "Raio fora do padrao")
                    nr = nr + 1
                End If
            End If
        End If
    End If

    ' cor de preenchimento
    If shp.Fill.Visible = msoTrue And shp.Fill.Type = msoFillSolid Then
        If Not IsPaletteColor(shp.Fill.ForeColor.RGB) Then
            col.Add Array(slideIdx, shp, "Cor de preenchimento fora da paleta")
            nc = nc + 1
        End If
    End If
    ' cor de contorno
    If shp.Line.Visible = msoTrue Then
        If Not IsPaletteColor(shp.Line.ForeColor.RGB) Then
            col.Add Array(slideIdx, shp, "Cor de contorno fora da paleta")
            nc = nc + 1
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub AuditScan(control As IRibbonControl)
    If Not ConfirmBigDeck("auditar") Then Exit Sub
    Dim col As New Collection, shp As Object
    Dim nf As Long, nc As Long, nr As Long, i As Long
    Dim it As Variant, resp As VbMsgBoxResult
    nf = 0: nc = 0: nr = 0
    For i = 1 To ActivePresentation.Slides.Count
        For Each shp In ActivePresentation.Slides(i).Shapes
            ScanShape shp, i, col, nf, nc, nr
        Next shp
    Next i

    If col.Count = 0 Then
        MsgBox "Auditoria: tudo dentro do padrao.", vbInformation, "CBA Studio"
        Exit Sub
    End If

    resp = MsgBox("Auditoria encontrou:" & vbCrLf & _
        "- " & nf & " fonte(s) fora do padrao" & vbCrLf & _
        "- " & nc & " cor(es) fora da paleta" & vbCrLf & _
        "- " & nr & " raio(s) fora do padrao" & vbCrLf & vbCrLf & _
        "Quer percorrer os alertas um a um?", vbYesNo + vbInformation, "CBA Studio")
    If resp <> vbYes Then Exit Sub

    For i = 1 To col.Count
        it = col(i)
        On Error Resume Next
        ActiveWindow.View.GotoSlide CLng(it(0))
        it(1).Select
        On Error GoTo 0
        resp = MsgBox("Alerta " & i & " de " & col.Count & " - Slide " & it(0) & vbCrLf & _
            it(2) & vbCrLf & vbCrLf & "OK = proximo / Cancelar = parar", _
            vbOKCancel + vbExclamation, "CBA Studio - Auditoria")
        If resp = vbCancel Then Exit For
    Next i
End Sub

' Padronizar tipografia: por run -> Avenir Next (heavy/black/bold -> Bold) e
' tamanho encaixado na escala B+G mais proxima.
Public Sub AuditFixFonts(control As IRibbonControl)
    Dim alvo As Collection, sld As Object, shp As Object, n As Long
    Set alvo = TargetSlides()
    If alvo.Count = 0 Then
        MsgBox "Selecione um ou mais slides no painel de miniaturas.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    n = 0
    For Each sld In alvo
        For Each shp In sld.Shapes
            n = n + FixTypeShape(shp)
        Next shp
    Next sld
    MsgBox "Tipografia padronizada (Avenir Next + peso + tamanho na escala)" & ScopeMsg(n, alvo.Count), vbInformation, "CBA Studio"
End Sub

' ============================================================
'  AUDITORIA: padronizar CORES na paleta (4 alvos), PAGE SIZE
'  e ponte para a analise de IMAGENS (na extensao)
' ============================================================
Private Sub RGBParts(ByVal c As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
    r = c And &HFF
    g = (c \ &H100) And &HFF
    b = (c \ &H10000) And &HFF
End Sub

' Luminancia perceptual aproximada, 0..1.
Private Function LumOf(ByVal c As Long) As Double
    Dim r As Long, g As Long, b As Long
    RGBParts c, r, g, b
    LumOf = (0.299 * r + 0.587 * g + 0.114 * b) / 255#
End Function

Private Sub RGBtoHSL(ByVal c As Long, ByRef h As Double, ByRef s As Double, ByRef l As Double)
    Dim r As Double, g As Double, b As Double, mx As Double, mn As Double, d As Double
    Dim ri As Long, gi As Long, bi As Long
    RGBParts c, ri, gi, bi
    r = ri / 255#: g = gi / 255#: b = bi / 255#
    mx = r: If g > mx Then mx = g
    If b > mx Then mx = b
    mn = r: If g < mn Then mn = g
    If b < mn Then mn = b
    l = (mx + mn) / 2#
    d = mx - mn
    If d < 0.0001 Then
        h = 0#: s = 0#
    Else
        If l > 0.5 Then s = d / (2# - mx - mn) Else s = d / (mx + mn)
        If mx = r Then
            h = 60# * ((g - b) / d)
            If h < 0 Then h = h + 360#
        ElseIf mx = g Then
            h = 60# * ((b - r) / d) + 120#
        Else
            h = 60# * ((r - g) / d) + 240#
        End If
    End If
End Sub

' Fundo do slide como luminancia (sem leitura -> assume claro/branco).
Private Function SlideBgLum(ByVal sld As Object) As Double
    Dim c As Long
    c = -1
    On Error Resume Next
    c = sld.Background.Fill.ForeColor.RGB
    On Error GoTo 0
    If c < 0 Then SlideBgLum = 1# Else SlideBgLum = LumOf(c)
End Function

' Mapeia uma cor qualquer para as 4 da marca (magenta, azul-violeta, bege,
' branco), aproximando por saturacao/matiz e preservando contraste com o fundo.
Private Function MapBrandColor(ByVal c As Long, ByVal bgLum As Double, ByVal isText As Boolean) As Long
    Dim h As Double, s As Double, l As Double, tgt As Long
    EnsureCfg
    If c = gPal(0) Or c = gPal(1) Or c = gPal(2) Or c = gPal(3) Then
        MapBrandColor = c
        Exit Function
    End If
    RGBtoHSL c, h, s, l
    If s < 0.12 Then                          ' neutros (cinzas, preto, off-white)
        If l >= 0.85 Then
            tgt = gPal(3)                     ' branco
        ElseIf l >= 0.55 Then
            tgt = gPal(2)                     ' bege
        Else                                  ' escuros: tudo nas 4 -> por contraste
            If bgLum < 0.5 Then tgt = gPal(3) Else tgt = gPal(1)
        End If
    Else                                      ' saturados: quente -> magenta, frio -> azul-violeta
        If h < 75# Or h >= 300# Then tgt = gPal(0) Else tgt = gPal(1)
    End If
    If isText Then                            ' trava de contraste figura-fundo
        If Abs(LumOf(tgt) - bgLum) < 0.25 Then
            If bgLum < 0.5 Then tgt = gPal(3) Else tgt = gPal(1)
        End If
    End If
    MapBrandColor = tgt
End Function

Private Function FixColorsShape(ByVal shp As Object, ByVal slideLum As Double, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim n As Long, s As Object, c As Long, t As Long, textBg As Double, run As Object
    n = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            n = n + FixColorsShape(s, slideLum, depth + 1)
        Next s
        FixColorsShape = n
        Exit Function
    End If
    If shp.Type = msoPicture Then FixColorsShape = 0: Exit Function
    ' leituras defensivas: se a propriedade falhar no Mac, o If com erro
    ' "cairia" dentro do bloco (Resume Next) e pintaria fill inexistente.
    Dim fVis As Long, fTyp As Long, fTr As Double, lVis As Long
    fVis = 0: Err.Clear
    fVis = shp.Fill.Visible
    If Err.Number <> 0 Then fVis = 0: Err.Clear
    fTyp = -1
    If fVis = msoTrue Then fTyp = shp.Fill.Type
    If Err.Number <> 0 Then fTyp = -1: Err.Clear
    fTr = 1#
    If fTyp = msoFillSolid Then fTr = shp.Fill.Transparency
    If Err.Number <> 0 Then fTr = 1#: Err.Clear
    ' preenchimento solido (e visivel de fato)
    If fVis = msoTrue And fTyp = msoFillSolid And fTr < 0.95 Then
        c = -1: c = shp.Fill.ForeColor.RGB
        If Err.Number = 0 And c >= 0 Then
            t = MapBrandColor(c, slideLum, False)
            If t <> c Then shp.Fill.ForeColor.RGB = t: n = n + 1
        End If
        Err.Clear
    End If
    ' contorno
    lVis = 0: lVis = shp.Line.Visible
    If Err.Number <> 0 Then lVis = 0: Err.Clear
    If lVis = msoTrue Then
        c = -1: c = shp.Line.ForeColor.RGB
        If Err.Number = 0 And c >= 0 Then
            t = MapBrandColor(c, slideLum, False)
            If t <> c Then shp.Line.ForeColor.RGB = t: n = n + 1
        End If
        Err.Clear
    End If
    ' texto, run a run (fundo do texto = fill NOVO do shape, ou o slide)
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            textBg = slideLum
            If fVis = msoTrue And fTyp = msoFillSolid And fTr < 0.95 Then textBg = LumOf(shp.Fill.ForeColor.RGB)
            For Each run In shp.TextFrame.TextRange.Runs
                c = -1: c = run.Font.Color.RGB
                If Err.Number = 0 And c >= 0 Then
                    t = MapBrandColor(c, textBg, True)
                    If t <> c Then run.Font.Color.RGB = t: n = n + 1
                End If
                Err.Clear
            Next run
        End If
    End If
    On Error GoTo 0
    FixColorsShape = n
End Function

Public Sub AuditFixColors(control As IRibbonControl)
    Dim alvo As Collection, sld As Object, shp As Object, n As Long, bg As Double
    Set alvo = TargetSlides()
    If alvo.Count = 0 Then
        MsgBox "Selecione um ou mais slides no painel de miniaturas.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    n = 0
    For Each sld In alvo
        bg = SlideBgLum(sld)
        For Each shp In sld.Shapes
            n = n + FixColorsShape(shp, bg)
        Next shp
    Next sld
    MsgBox "Cores padronizadas na paleta (magenta, azul-violeta, bege, branco)" & ScopeMsg(n, alvo.Count), vbInformation, "CBA Studio"
End Sub

' Forca o tamanho padrao do template B+G: 1583 x 891 pt (55,85 x 31,43 cm, 16:9).
'
' LICAO APRENDIDA (nao refazer): a versao que escalava o conteudo na mao
' quebrava o layout e era lenta. Dois motivos:
'  (a) FONTE — o PowerPoint escala o texto pelo fator de autoajuste da caixa,
'      que NAO aparece em Font.Size. Ler Font.Size igual antes/depois parece
'      "ele nao escalou", mas escalou. Multiplicar Font.Size por cima disso
'      dobrava o tamanho e, em caixas com autoajuste, mudava a ALTURA delas —
'      e ai todo o layout relativo saia do lugar. Fonte: nunca mais tocamos.
'  (b) VELOCIDADE — percorrer cada run de cada caixa de texto de cada slide
'      era o que deixava lento em decks grandes.
' Hoje: confiamos no redimensionamento nativo (que faz o certo quando a
' proporcao e' mantida, o caso normal 16:9 -> 16:9) e so' conferimos, por
' AMOSTRAGEM barata, se ele realmente mexeu. Se nao mexeu, PERGUNTAMOS antes
' de ajustar a geometria — nunca em silencio.
Public Sub FixPageSize(control As IRibbonControl)
    Dim w As Single, hgt As Single
    On Error Resume Next
    w = ActivePresentation.PageSetup.SlideWidth
    hgt = ActivePresentation.PageSetup.SlideHeight
    On Error GoTo 0
    If Abs(w - 1583) < 1 And Abs(hgt - 891) < 1 Then
        MsgBox "A apresentacao ja esta no formato padrao B+G (1583 x 891 pt).", vbInformation, "CBA Studio"
        Exit Sub
    End If

    Dim aviso As String
    aviso = ""
    If w > 1 And hgt > 1 Then
        If Abs((w / hgt) - (1583# / 891#)) > 0.02 Then
            aviso = vbCrLf & vbCrLf & "Atencao: a proporcao atual e' diferente de 16:9, entao alguns" & vbCrLf & _
                    "objetos podem precisar de reposicionamento depois."
        End If
    End If
    If MsgBox("Mudar o tamanho de " & CLng(w) & " x " & CLng(hgt) & " pt para o padrao B+G 1583 x 891 pt (55,85 x 31,43 cm)?" & aviso, _
              vbQuestion + vbOKCancel, "CBA Studio") <> vbOK Then Exit Sub

    ' Amostra barata (ate' 24 formas espalhadas pelo deck) pra saber DEPOIS se
    ' o PowerPoint mexeu no conteudo. Nao percorre o deck inteiro.
    Dim smpSld() As Long, smpShp() As Long, smpW() As Single, nSmp As Long
    nSmp = SampleShapes(smpSld, smpShp, smpW)

    ' O PowerPoint REJEITA (silenciosamente) mudar uma dimensao se o resultado
    ' intermediario ficar com proporcao muito extrema — ex.: pôr 1583 de largura
    ' enquanto a altura ainda e' 540 (1583x540 = 2,9:1) e' bloqueado, e so' a
    ' altura mudava -> slide quase quadrado. Solucao: altura PRIMEIRO (proporcao
    ' vai ficando menos extrema) e repetir em passadas ate' fixar.
    Dim i As Long
    On Error Resume Next
    For i = 1 To 4
        ActivePresentation.PageSetup.SlideHeight = 891
        ActivePresentation.PageSetup.SlideWidth = 1583
    Next i
    Dim novoW As Single, novoH As Single
    novoW = ActivePresentation.PageSetup.SlideWidth
    novoH = ActivePresentation.PageSetup.SlideHeight
    On Error GoTo 0

    Dim chegou As Boolean
    chegou = (Abs(novoW - 1583) < 1 And Abs(novoH - 891) < 1)

    Dim extra As String
    extra = ""
    If nSmp > 0 And chegou And w > 1 And hgt > 1 Then
        If Not SampleMoved(nSmp, smpSld, smpShp, smpW) Then
            ' Caso raro: a pagina mudou e o conteudo ficou parado. Oferece
            ' o ajuste em vez de fazer sozinho (evita surpresa em deck bom).
            If MsgBox("A pagina mudou, mas o PowerPoint nao redimensionou o conteudo." & vbCrLf & vbCrLf & _
                      "Quer que eu redimensione os objetos proporcionalmente?" & vbCrLf & _
                      "(o corpo das fontes nao e' alterado)", _
                      vbQuestion + vbYesNo, "CBA Studio") = vbYes Then
                ' TryScaleContent NUNCA gera erro: se algo falhar, mantem so' o
                ' slide redimensionado e devolve um aviso gentil.
                extra = TryScaleContent(w, hgt, novoW, novoH)
            End If
        End If
    End If

    If chegou Then
        MsgBox "Formato ajustado para 1583 x 891 pt (55,85 x 31,43 cm)." & extra, vbInformation, "CBA Studio"
    Else
        MsgBox "Ajustei para " & CLng(novoW) & " x " & CLng(novoH) & " pt, mas nao cheguei ao alvo exato" & vbCrLf & _
               "(proporcao de origem muito extrema). Ajuste em Design > Tamanho do Slide.", vbExclamation, "CBA Studio"
    End If
End Sub

' Amostra ate' SAMPLE_MAX formas espalhadas pelo deck (no maximo 3 por slide,
' pulando slides em decks grandes) e guarda a largura de cada uma.
Private Function SampleShapes(ByRef sldIx() As Long, ByRef shpIx() As Long, ByRef wid() As Single) As Long
    Dim n As Long, i As Long, j As Long, nSld As Long, passo As Long, noSlide As Long, w As Single
    ReDim sldIx(1 To SAMPLE_MAX)
    ReDim shpIx(1 To SAMPLE_MAX)
    ReDim wid(1 To SAMPLE_MAX)
    n = 0
    On Error Resume Next
    nSld = ActivePresentation.Slides.Count
    If nSld < 1 Then SampleShapes = 0: Exit Function
    passo = 1
    If nSld > 8 Then passo = nSld \ 8
    For i = 1 To nSld Step passo
        noSlide = 0
        For j = 1 To ActivePresentation.Slides(i).Shapes.Count
            If n >= SAMPLE_MAX Or noSlide >= 3 Then Exit For
            w = 0
            w = ActivePresentation.Slides(i).Shapes(j).Width
            If w > 1 Then
                n = n + 1: noSlide = noSlide + 1
                sldIx(n) = i: shpIx(n) = j: wid(n) = w
            End If
        Next j
        If n >= SAMPLE_MAX Then Exit For
    Next i
    On Error GoTo 0
    SampleShapes = n
End Function

' A maioria das formas amostradas mudou de largura? (= o PowerPoint escalou)
Private Function SampleMoved(ByVal n As Long, ByRef sldIx() As Long, ByRef shpIx() As Long, ByRef wid() As Single) As Boolean
    Dim i As Long, mudou As Long, w As Single
    mudou = 0
    On Error Resume Next
    For i = 1 To n
        w = 0
        w = ActivePresentation.Slides(sldIx(i)).Shapes(shpIx(i)).Width
        If w > 0 Then
            If Abs(w - wid(i)) > 0.5 Then mudou = mudou + 1
        End If
    Next i
    On Error GoTo 0
    SampleMoved = (mudou * 2 > n)          ' mais da metade
End Function

' Tenta redimensionar o conteudo proporcionalmente. NUNCA gera erro do VBA:
' o handler pega qualquer imprevisto e o slide continua redimensionado, so'
' com um aviso gentil. Devolve a frase de status pra mensagem final.
Private Function TryScaleContent(ByVal w As Single, ByVal hgt As Single, _
                                 ByVal novoW As Single, ByVal novoH As Single) As String
    Dim k As Single, offX As Single, offY As Single, nEsc As Long, AVISO As String
    AVISO = vbCrLf & "Mantive o tamanho do slide; ajuste os objetos manualmente se precisar."
    TryScaleContent = ""
    On Error GoTo falhou
    If w <= 1 Or hgt <= 1 Then Exit Function
    k = novoW / w
    If (novoH / hgt) < k Then k = novoH / hgt      ' uniforme: nao distorce
    If k <= 0 Or k > 100 Then TryScaleContent = AVISO: Exit Function
    offX = (novoW - w * k) / 2#
    offY = (novoH - hgt * k) / 2#
    nEsc = ScaleAllGeometry(k, offX, offY)
    If nEsc > 0 Then
        TryScaleContent = vbCrLf & "Conteudo redimensionado em " & nEsc & " objeto(s)."
    Else
        TryScaleContent = AVISO
    End If
    Exit Function
falhou:
    ' qualquer erro inesperado: o slide ja' foi redimensionado; nada de dialogo
    TryScaleContent = AVISO
End Function

' Escala SO' a geometria (posicao e tamanho) de todos os slides pelo fator k,
' deslocando por (offX, offY). Nao toca em fonte — ver a licao no topo.
' Cada forma e' tratada em ScaleOneShape com protecao total: uma forma
' problematica e' PULADA, nunca derruba o resto.
Private Function ScaleAllGeometry(ByVal k As Single, ByVal offX As Single, ByVal offY As Single) As Long
    Dim sld As Object, shp As Object, n As Long
    n = 0
    If k <= 0 Or k > 100 Then ScaleAllGeometry = 0: Exit Function
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If ScaleOneShape(shp, k, offX, offY) Then n = n + 1
        Next shp
    Next sld
    On Error GoTo 0
    ScaleAllGeometry = n
End Function

' Escala UMA forma (grupos incluidos: mexer no grupo ja' move os filhos).
' Le cada dimensao pra variavel local, valida faixa sã (sentinelas de
' placeholder retornam valores absurdos que estouram na multiplicacao) e so'
' entao grava. Qualquer erro -> pula a forma. Devolve True se conseguiu.
Private Function ScaleOneShape(ByVal shp As Object, ByVal k As Single, ByVal offX As Single, ByVal offY As Single) As Boolean
    Dim l As Single, t As Single, wd As Single, ht As Single
    ScaleOneShape = False
    On Error Resume Next
    l = 1000000: t = 1000000: wd = -1: ht = -1
    l = shp.Left: t = shp.Top: wd = shp.Width: ht = shp.Height
    If Err.Number <> 0 Then Err.Clear: Exit Function
    ' faixa sã de um slide (em pontos). Fora disso e' sentinela/objeto especial.
    If Abs(l) > 100000 Or Abs(t) > 100000 Then Exit Function
    If wd <= 0 Or wd > 100000 Or ht <= 0 Or ht > 100000 Then Exit Function
    shp.Left = l * k + offX
    shp.Top = t * k + offY
    shp.Width = wd * k
    shp.Height = ht * k
    If Err.Number = 0 Then ScaleOneShape = True
    Err.Clear
End Function

' Sobre: mostra a versao instalada da faixa. Compare com a versao no site
' (doodle-studio-sigma.vercel.app) — atualizar = rodar o instalador de novo.
Public Sub ShowAbout(control As IRibbonControl)
    MsgBox "CBA Studio - faixa v" & CBA_VERSION & vbCrLf & vbCrLf & _
           "Pra atualizar: rode o instalador de novo" & vbCrLf & _
           "(comando do site) e reabra o PowerPoint." & vbCrLf & vbCrLf & _
           "doodle-studio-sigma.vercel.app", vbInformation, "CBA Studio"
End Sub

' A analise de imagens (MB por arquivo, conversao) exige ler o .pptx —
' so' a extensao consegue; este botao aponta o caminho.
Public Sub AuditImages(control As IRibbonControl)
    MsgBox "A analise de imagens fica no painel lateral:" & vbCrLf & vbCrLf & _
           "Inserir > Suplementos > Meus Suplementos > CBA Studio" & vbCrLf & _
           "-> secao 'Imagens do arquivo' -> Analisar.", vbInformation, "CBA Studio"
End Sub

Private Function FixTypeShape(ByVal shp As Object, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object, tr As Object, k As Long
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + FixTypeShape(s, depth + 1)
        Next s
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            Set tr = shp.TextFrame.TextRange
            For k = 1 To tr.Runs.Count
                FixTypeRun tr.Runs(k)
            Next k
            cnt = 1
        End If
    End If
    On Error GoTo 0
    FixTypeShape = cnt
End Function

Private Sub FixTypeRun(ByVal run As Object)
    Dim heavy As Boolean
    On Error Resume Next
    heavy = IsHeavyFont(run.Font.Name, run.Font.Bold)
    run.Font.Name = gFonte                       ' "Avenir Next"
    If heavy Then run.Font.Bold = msoTrue Else run.Font.Bold = msoFalse
    run.Font.size = SnapSize(run.Font.size)
    On Error GoTo 0
End Sub

' peso acima de medio (heavy/black/bold/semibold/demi) -> Avenir Next Bold
Private Function IsHeavyFont(ByVal nm As String, ByVal isBold As Long) As Boolean
    Dim s As String
    s = LCase$(nm)
    If InStr(s, "heavy") > 0 Or InStr(s, "black") > 0 Or InStr(s, "bold") > 0 _
       Or InStr(s, "semibold") > 0 Or InStr(s, "demi") > 0 Then
        IsHeavyFont = True
    ElseIf isBold = msoTrue Then
        IsHeavyFont = True
    End If
End Function

' encaixa o tamanho no valor mais proximo da escala configurada (estilos)
Private Function SnapSize(ByVal sz As Single) As Single
    Dim it As Variant, best As Single, bd As Single, d As Single
    EnsureCfg
    best = sz: bd = 1E+30
    For Each it In gStyle
        d = Abs(sz - it(0))
        If d < bd Then bd = d: best = it(0)
    Next it
    SnapSize = best
End Function

' ============================================================
'  Tabela de estilos B+G (espelha STYLES de typography.js)
' ============================================================
Private Function SpecFor(ByVal id As String) As StyleSpec
    Dim s As StyleSpec, a As Variant
    EnsureCfg
    s.found = False
    s.periodColor = -1
    s.bgAware = False
    On Error Resume Next
    a = gStyle(id)
    On Error GoTo 0
    If IsEmpty(a) Then SpecFor = s: Exit Function
    s.found = True
    s.size = a(0)
    s.bold = a(1)
    If a(2) = 0 Then s.color = gPal(0) Else s.color = gPal(1)
    If id = "dsHero" Then           ' Hero: ponto azul + ciente do fundo
        s.periodColor = gPal(1)
        s.bgAware = True
    End If
    SpecFor = s
End Function

' Mapa tamanho(pt) -> entrelinha (multiplo). 0 = fora do mapa, nao mexe.
Private Function MapSpacing(ByVal sz As Single) As Single
    Dim v As Variant
    EnsureCfg
    On Error Resume Next
    v = gEnt(CStr(CLng(sz)))
    On Error GoTo 0
    If IsEmpty(v) Then MapSpacing = 0 Else MapSpacing = v
End Function

' ============================================================
'  Aplicacao de ESTILO (fonte+tamanho+cor+ponto+entrelinha)
' ============================================================
Private Function ApplyStyleToShape(ByVal shp As Object, ByRef spec As StyleSpec, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object, isBlue As Boolean
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + ApplyStyleToShape(s, spec, depth + 1)
        Next s
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            isBlue = False
            If spec.bgAware Then isBlue = ShapeFillIsBlue(shp)
            ApplyStyleToRange shp.TextFrame.TextRange, spec, isBlue
            cnt = 1
        End If
    End If
    On Error GoTo 0
    ApplyStyleToShape = cnt
End Function

Private Function ShapeFillIsBlue(ByVal shp As Object) As Boolean
    On Error Resume Next
    If shp.Fill.Type = msoFillSolid Then
        ShapeFillIsBlue = (shp.Fill.ForeColor.RGB = gPal(1))
    End If
    On Error GoTo 0
End Function

Private Sub ApplyStyleToRange(ByVal tr As Object, ByRef spec As StyleSpec, ByVal bgBlue As Boolean)
    Dim textColor As Long, periodColor As Long, ls As Single
    Dim p As Long, para As Object, t As String

    textColor = spec.color
    periodColor = spec.periodColor
    If spec.bgAware And bgBlue Then
        textColor = gPal(3)        ' branco
        periodColor = gPal(0)      ' rosa
    End If

    ' fonte / tamanho / negrito / cor
    tr.Font.Name = gFonte
    tr.Font.size = spec.size
    If spec.bold Then tr.Font.Bold = msoTrue Else tr.Font.Bold = msoFalse
    tr.Font.Color.RGB = textColor

    ' entrelinha exata (todos os paragrafos ficam no tamanho do estilo)
    ls = MapSpacing(spec.size)
    If ls > 0 Then
        For p = 1 To tr.Paragraphs.Count
            Set para = tr.Paragraphs(p, 1)
            para.ParagraphFormat.LineRuleWithin = msoTrue
            para.ParagraphFormat.SpaceWithin = ls
        Next p
    End If

    ' ponto final em outra cor
    If periodColor >= 0 Then
        t = tr.Text
        If Len(t) > 0 Then
            If Right$(t, 1) = "." Then
                tr.Characters(Len(t), 1).Font.Color.RGB = periodColor
            End If
        End If
    End If
End Sub

' ============================================================
'  Entrelinha avulsa (le' o tamanho de CADA paragrafo)
'  Para textos ja' formatados onde os tamanhos variam.
' ============================================================
' Aplica a entrelinha B+G ao que estiver selecionado: texto, objetos, ou
' SLIDES selecionados no painel (todas as caixas de texto desses slides).
Private Sub DoEntrelinhaSelecao()
    Dim sel As Object, shp As Object, sld As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionText Then
        ApplySpacingToRange sel.TextRange
        n = 1
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            n = n + ApplyToShape(shp)
        Next shp
    ElseIf sel.Type = ppSelectionSlides Then
        On Error Resume Next
        For Each sld In sel.SlideRange
            For Each shp In sld.Shapes
                n = n + ApplyToShape(shp)
            Next shp
        Next sld
        On Error GoTo 0
    End If
    If n = 0 Then MsgBox "Selecione um texto, objetos ou slides no painel.", vbInformation, "CBA Studio"
End Sub

Private Sub DoEntrelinhaSlides()
    Dim alvo As Collection, sld As Object, shp As Object, n As Long
    Set alvo = TargetSlides()
    If alvo.Count = 0 Then
        MsgBox "Selecione um ou mais slides no painel de miniaturas.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    n = 0
    For Each sld In alvo
        For Each shp In sld.Shapes
            n = n + ApplyToShape(shp)
        Next shp
    Next sld
    MsgBox "Entrelinha B+G aplicada" & ScopeMsg(n, alvo.Count), vbInformation, "CBA Studio"
End Sub

Private Sub ApplySpacingToRange(ByVal tr As Object)
    Dim p As Long, para As Object, ls As Single
    For p = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(p, 1)
        ls = MapSpacing(para.Font.size)
        If ls > 0 Then
            para.ParagraphFormat.LineRuleWithin = msoTrue
            para.ParagraphFormat.SpaceWithin = ls
        End If
    Next p
End Sub

Private Function ApplyToShape(ByVal shp As Object, Optional ByVal depth As Long = 0) As Long
    If depth > MAX_DEPTH Then Exit Function      ' guarda contra grupos aninhados demais
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + ApplyToShape(s, depth + 1)
        Next s
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            ApplySpacingToRange shp.TextFrame.TextRange
            cnt = 1
        End If
    End If
    On Error GoTo 0
    ApplyToShape = cnt
End Function
