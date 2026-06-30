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

' Cores da marca (RGB usa ordem R,G,B)
Private Const C_ROSA As Long = 7036668     ' FC5E6D -> RGB(252, 94,109)
Private Const C_AZUL As Long = 14773315    ' 436AE1 -> RGB( 67,106,225)
Private Const C_BRANCO As Long = 16777215  ' FFFFFF
Private Const FONTE As String = "Avenir Next"
Private Const C_TRANSP As Long = 5         ' indice da cor = transparente

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

' Alinhar/distribuir tradicionais (idMso ObjectsAlignLeft/Top nao renderizam no Mac;
' usamos ShapeRange.Align/Distribute relativo aos objetos selecionados).
Public Sub AlignLeftEdges(control As IRibbonControl)
    AlignSel 1          ' msoAlignLefts
End Sub

Public Sub AlignTopEdges(control As IRibbonControl)
    AlignSel 4          ' msoAlignTops
End Sub

Private Sub AlignSel(ByVal mode As Long)
    Dim sel As Object
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Selecione dois ou mais objetos para alinhar.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    On Error Resume Next
    sel.ShapeRange.Align mode, msoFalse        ' relativo aos objetos selecionados
    On Error GoTo 0
End Sub

Public Sub DistributeH(control As IRibbonControl)
    Dim sel As Object
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Selecione tres ou mais objetos para distribuir.", vbInformation, "CBA Studio"
        Exit Sub
    End If
    On Error Resume Next
    sel.ShapeRange.Distribute 1, msoFalse      ' msoDistributeHorizontally, relativo
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
    tb.TextFrame.TextRange.Font.Name = FONTE
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
    tr2.Font.Caps = 2          ' msoCapsAll (constante nao definida no Mac)
    tr2.Font.Spacing = LooseSpacingPts(tr2.Font.size)
    On Error GoTo 0
End Sub

' tracking "loose" ~ 6% do tamanho da fonte (em pontos)
Private Function LooseSpacingPts(ByVal sz As Single) As Single
    If sz <= 0 Then sz = 18
    LooseSpacingPts = sz * 0.06
End Function

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
    StdRadiusPts = (25# / 1080#) * ActivePresentation.PageSetup.SlideHeight
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
    Select Case i
        Case 0: PaletteColor = RGB(253, 94, 109)   ' FD5E6D rosa
        Case 1: PaletteColor = RGB(67, 106, 225)    ' 436AE1 azul
        Case 2: PaletteColor = RGB(238, 236, 230)   ' EEECE6 bege
        Case 3: PaletteColor = RGB(255, 255, 255)   ' branco
        Case 4: PaletteColor = RGB(0, 0, 0)         ' preto
        Case Else: PaletteColor = RGB(0, 0, 0)
    End Select
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
            If Len(nm) > 0 And StrComp(nm, FONTE, vbTextCompare) <> 0 Then
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
    run.Font.Name = FONTE                       ' "Avenir Next"
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

' encaixa o tamanho no valor mais proximo da escala B+G
Private Function SnapSize(ByVal sz As Single) As Single
    Dim sc As Variant, i As Long, best As Single, bd As Single, d As Single
    sc = Array(16, 18, 20, 24, 28, 34, 44, 60, 80, 120)
    best = sz: bd = 1E+30
    For i = LBound(sc) To UBound(sc)
        d = Abs(sz - sc(i))
        If d < bd Then bd = d: best = sc(i)
    Next i
    SnapSize = best
End Function

' ============================================================
'  Tabela de estilos B+G (espelha STYLES de typography.js)
' ============================================================
Private Function SpecFor(ByVal id As String) As StyleSpec
    Dim s As StyleSpec
    s.found = True
    s.periodColor = -1
    s.bgAware = False
    Select Case id
        Case "dsHero":       s.size = 120: s.bold = True:  s.color = C_ROSA: s.periodColor = C_AZUL: s.bgAware = True
        Case "dsMega":       s.size = 80:  s.bold = True:  s.color = C_AZUL
        Case "dsH1":         s.size = 60:  s.bold = True:  s.color = C_ROSA
        Case "dsLabelSec":   s.size = 60:  s.bold = True:  s.color = C_ROSA
        Case "dsCorpo":      s.size = 44:  s.bold = False: s.color = C_AZUL
        Case "dsH3":         s.size = 34:  s.bold = True:  s.color = C_ROSA
        Case "dsH4":         s.size = 28:  s.bold = True:  s.color = C_AZUL
        Case "dsH5":         s.size = 24:  s.bold = False: s.color = C_AZUL
        Case "dsCorpoPilar": s.size = 20:  s.bold = False: s.color = C_AZUL
        Case "dsEyebrow":    s.size = 18:  s.bold = True:  s.color = C_ROSA
        Case "dsCaption":    s.size = 16:  s.bold = False: s.color = C_AZUL
        Case Else:           s.found = False
    End Select
    SpecFor = s
End Function

' Mapa tamanho(pt) -> entrelinha (multiplo) dos estilos B+G
Private Function MapSpacing(ByVal sz As Single) As Single
    Select Case CLng(sz)
        Case 120: MapSpacing = 0.8
        Case 80:  MapSpacing = 0.9
        Case 60:  MapSpacing = 0.9
        Case 44:  MapSpacing = 1.15
        Case 34:  MapSpacing = 0.95
        Case 28:  MapSpacing = 1#
        Case 24:  MapSpacing = 1#
        Case 20:  MapSpacing = 1.3
        Case 18:  MapSpacing = 1#
        Case 16:  MapSpacing = 1.3
        Case Else: MapSpacing = 0      ' fora do mapa: nao mexe
    End Select
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
        ShapeFillIsBlue = (shp.Fill.ForeColor.RGB = C_AZUL)
    End If
    On Error GoTo 0
End Function

Private Sub ApplyStyleToRange(ByVal tr As Object, ByRef spec As StyleSpec, ByVal bgBlue As Boolean)
    Dim textColor As Long, periodColor As Long, ls As Single
    Dim p As Long, para As Object, t As String

    textColor = spec.color
    periodColor = spec.periodColor
    If spec.bgAware And bgBlue Then
        textColor = C_BRANCO
        periodColor = C_ROSA
    End If

    ' fonte / tamanho / negrito / cor
    tr.Font.Name = FONTE
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
