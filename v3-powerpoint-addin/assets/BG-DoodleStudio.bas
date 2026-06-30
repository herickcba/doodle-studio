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
'    EntrelinhaTudo(control)-> entrelinha B+G em TODOS os slides
'  Tambem expostos sem args p/ Option+F8 (fallback):
'    BG_AplicarEntrelinha / BG_AplicarEntrelinhaTudo
' ============================================================
Option Explicit

Private gRibbon As IRibbonUI
Private gAnchorCm As Single   ' ancora esq. (cm); 0 = nao setado -> usa default
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
    ' No Mac, Environ("HOME") dentro do Office = a sandbox do PowerPoint
    ' (~/Library/Containers/com.microsoft.Powerpoint/Data), que e' gravavel
    ' por VBA sem pedir permissao. Guardamos a config global ali.
    ConfigPath = Environ$("HOME") & "/cba-config.txt"
End Function

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
            kv = Split(parts(i), "=")
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
        If pi >= 0 And pi <= 4 Then gPal(pi) = HexToRGB(v)
    ElseIf Left$(k, 2) = "s_" Then
        id = Mid$(k, 3)
        a = Split(v, "|")
        If UBound(a) >= 2 Then AddStyle id, CSng(Val(a(0))), (Trim$(a(1)) = "1"), CLng(Val(a(2)))
    ElseIf Left$(k, 4) = "ent_" Then
        sz = CLng(Val(Mid$(k, 5)))
        AddEnt sz, CSng(Val(v))
    End If
End Sub

Private Function HexToRGB(ByVal h As String) As Long
    h = Replace(h, "#", "")
    If Len(h) < 6 Then HexToRGB = 0: Exit Function
    Dim r As Long, g As Long, b As Long
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
Public Sub ConfigApply(control As IRibbonControl)
    Dim s As String
    s = InputBox("Cole a configuracao copiada da pagina de Brand Standards:", _
                 "CBA Studio — Aplicar padroes")
    If Len(Trim$(s)) = 0 Then Exit Sub
    SetDefaults
    ApplyConfigString s
    If WriteConfig(s) Then
        On Error Resume Next
        gRibbon.Invalidate
        On Error GoTo 0
        MsgBox "Padroes aplicados e salvos.", vbInformation, "CBA Studio"
    Else
        MsgBox "Padroes aplicados nesta sessao, mas nao consegui salvar o arquivo" & vbCrLf & _
               "(permita o acesso a' pasta quando o PowerPoint pedir).", vbExclamation, "CBA Studio"
    End If
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

Public Sub EntrelinhaTudo(control As IRibbonControl)
    DoEntrelinhaTudo
End Sub

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

Private Function SetSpacingOnShape(ByVal shp As Object, ByVal mult As Single) As Long
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + SetSpacingOnShape(s, mult)
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
Public Sub BG_AplicarEntrelinhaTudo()
    DoEntrelinhaTudo
End Sub

' ============================================================
'  ALINHAR a' ancora esquerda (medida definivel via editBox, em cm)
'  (gAnchorCm declarada no topo do modulo)
' ============================================================
Private Function AnchorCm() As Single
    If gAnchorCm <= 0 Then AnchorCm = 1.27 Else AnchorCm = gAnchorCm
End Function

Public Sub GetAnchorText(control As IRibbonControl, ByRef text As Variant)
    text = Format(AnchorCm, "0.0#")
End Sub

Public Sub SetAnchorText(control As IRibbonControl, text As String)
    Dim v As Single
    v = Val(Replace(text, ",", "."))
    If v > 0 Then gAnchorCm = v
End Sub

' Alinha a borda esquerda a' ancora.
'  - Se houver OBJETOS selecionados: alinha esses (qualquer tipo).
'  - Se NADA estiver selecionado: alinha automaticamente todas as
'    caixas de TEXTO do slide atual (sem precisar selecionar nada).
' Ancora a borda esquerda dos OBJETOS SELECIONADOS a' medida (cm). So' seleção.
Public Sub AlignAnchorLeft(control As IRibbonControl)
    Dim sel As Object, shp As Object, n As Long, x As Single
    Set sel = ActiveWindow.Selection
    x = AnchorCm * 28.3465                 ' cm -> pontos
    n = 0
    If sel.Type = ppSelectionShapes Then
        On Error Resume Next
        For Each shp In sel.ShapeRange
            shp.Left = x
            n = n + 1
        Next shp
        On Error GoTo 0
    End If
    If n = 0 Then MsgBox "Selecione um ou mais objetos para ancorar.", vbInformation, "CBA Studio"
End Sub

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

Private Function GroupHasText(ByVal grp As Object) As Boolean
    Dim s As Object
    On Error Resume Next
    For Each s In grp.GroupItems
        If s.Type = msoGroup Then
            If GroupHasText(s) Then GroupHasText = True: Exit Function
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
    x = AnchorCm * 28.3465                       ' margem esquerda da ancora
    w = ActivePresentation.PageSetup.SlideWidth / 3
    Set tb = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x, 80, w, 60)
    On Error Resume Next
    tb.TextFrame.WordWrap = msoTrue
    tb.TextFrame.TextRange.Text = "Texto"
    tb.TextFrame.TextRange.Font.Name = gFonte
    tb.Select
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
    shp.Select
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

Private Function ZeroInsets(ByVal shp As Object) As Long
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + ZeroInsets(s)
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
    EnsureCfg
    StdRadiusPts = (CDbl(gRadiusPx) / 1080#) * ActivePresentation.PageSetup.SlideHeight
End Function

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
Private Function MakeRounded(ByVal shp As Object) As Long
    Dim done As Long, s As Object, doIt As Boolean
    done = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            done = done + MakeRounded(s)
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

' normaliza, em toda a apresentacao, SO' as formas que ja' sao rounded
Public Sub RoundAllShapes(control As IRibbonControl)
    Dim sld As Object, shp As Object, n As Long
    n = 0
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            n = n + NormalizeRounded(shp)
        Next shp
    Next sld
    MsgBox "Cantos padronizados em " & n & " formas rounded.", vbInformation, "CBA Studio"
End Sub

Private Function NormalizeRounded(ByVal shp As Object) As Long
    Dim cnt As Long, s As Object, t As Long
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + NormalizeRounded(s)
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

Private Function MakeStraight(ByVal shp As Object) As Long
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + MakeStraight(s)
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

Private Function FillShape(ByVal shp As Object, ByVal i As Long) As Long
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + FillShape(s, i)
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

Private Function FontColorShape(ByVal shp As Object, ByVal i As Long) As Long
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + FontColorShape(s, i)
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

Private Function OutlineShape(ByVal shp As Object, ByVal i As Long) As Long
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + OutlineShape(s, i)
        Next s
    Else
        If i = C_TRANSP Then
            shp.Line.Visible = msoFalse
        Else
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = PaletteColor(i)
            shp.Line.Weight = 2                         ' contorno padrao 2pt
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
                      ByRef nf As Long, ByRef nc As Long, ByRef nr As Long)
    Dim s As Object, std As Single, adj As Single, nm As String
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            ScanShape s, slideIdx, col, nf, nc, nr
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
    Dim sld As Object, shp As Object, n As Long
    n = 0
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            n = n + FixTypeShape(shp)
        Next shp
    Next sld
    MsgBox "Tipografia padronizada (Avenir Next + peso + tamanho na escala) em " & n & " objeto(s).", vbInformation, "CBA Studio"
End Sub

Private Function FixTypeShape(ByVal shp As Object) As Long
    Dim cnt As Long, s As Object, tr As Object, k As Long
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + FixTypeShape(s)
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
Private Function ApplyStyleToShape(ByVal shp As Object, ByRef spec As StyleSpec) As Long
    Dim cnt As Long, s As Object, isBlue As Boolean
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + ApplyStyleToShape(s, spec)
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
Private Sub DoEntrelinhaSelecao()
    Dim sel As Object, shp As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionText Then
        ApplySpacingToRange sel.TextRange
        n = 1
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            n = n + ApplyToShape(shp)
        Next shp
    End If
    If n = 0 Then MsgBox "Selecione um texto ou um objeto com texto.", vbInformation, "CBA Studio"
End Sub

Private Sub DoEntrelinhaTudo()
    Dim sld As Object, shp As Object, n As Long
    n = 0
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            n = n + ApplyToShape(shp)
        Next shp
    Next sld
    MsgBox "Entrelinha B+G aplicada em " & n & " caixas de texto, em todos os slides.", vbInformation, "CBA Studio"
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

Private Function ApplyToShape(ByVal shp As Object) As Long
    Dim cnt As Long, s As Object
    cnt = 0
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            cnt = cnt + ApplyToShape(s)
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
