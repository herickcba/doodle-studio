Attribute VB_Name = "BG_Tipografia"
' ============================================================
'  B+G Tipografia — entrelinha + inserir bloco (PowerPoint VBA)
'  A API de add-in nao seta entrelinha; o VBA seta. Cole este
'  modulo no VBE (Option+F11 > Inserir > Modulo) ou importe o .bas.
'  Rode por Option+F8, ou adicione ao Quick Access Toolbar (1 clique).
'
'  BG_AplicarEntrelinha  -> finaliza o estilo aplicado no painel:
'      le o tamanho de cada paragrafo da selecao e aplica a
'      entrelinha B+G correspondente (multiplo).
'  BG_InserirBloco       -> cria uma caixa de texto ja 100% no estilo
'      (fonte, tamanho, cor, ponto final E entrelinha) no slide ativo.
' ============================================================
Option Explicit

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

Private Sub ApplySpacingToRange(ByVal tr As Object)
    Dim p As Long, para As Object, ls As Single
    For p = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(p, 1)
        ls = MapSpacing(para.Font.Size)
        If ls > 0 Then
            para.ParagraphFormat.LineRuleWithin = msoTrue
            para.ParagraphFormat.SpaceWithin = ls
        End If
    Next p
End Sub

' === Aplica a entrelinha B+G a selecao (texto ou objeto[s]) ===
Public Sub BG_AplicarEntrelinha()
    Dim sel As Object, shp As Object, n As Long
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionText Then
        ApplySpacingToRange sel.TextRange
        n = 1
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    ApplySpacingToRange shp.TextFrame.TextRange
                    n = n + 1
                End If
            End If
        Next shp
    End If
    If n = 0 Then MsgBox "Selecione um texto ou um objeto com texto.", vbInformation, "B+G Tipografia"
End Sub

' === Insere uma caixa de texto ja no estilo escolhido ===
Public Sub BG_InserirBloco()
    Dim choice As String, k As Integer
    choice = InputBox( _
        "Escolha o estilo:" & vbCrLf & _
        "1  Hero (120 / 0.8x)" & vbCrLf & _
        "2  Mega (80 / 0.9x)" & vbCrLf & _
        "3  H1 Titulo (60 / 0.9x)" & vbCrLf & _
        "4  Label de Secao (60 / 0.9x)" & vbCrLf & _
        "5  Corpo de Texto (44 / 1.15x)" & vbCrLf & _
        "6  H3 (34 / 0.95x)" & vbCrLf & _
        "7  H4 (28 / 1.0x)" & vbCrLf & _
        "8  H5 (24 / 1.0x)" & vbCrLf & _
        "9  Corpo de Pilar (20 / 1.3x)" & vbCrLf & _
        "10 Eyebrow (18 / 1.0x)" & vbCrLf & _
        "11 Caption (16 / 1.3x)", "B+G — Inserir bloco", "1")
    If choice = "" Then Exit Sub
    k = Val(choice)
    If k < 1 Or k > 11 Then Exit Sub

    Dim RED As Long, BLUE As Long
    RED = RGB(252, 94, 109)    ' FC5E6D
    BLUE = RGB(67, 106, 225)   ' 436AE1

    Dim sample As String, sz As Single, ls As Single, isBold As Boolean
    Dim col As Long, perCol As Long, hasPeriod As Boolean
    hasPeriod = False: perCol = 0
    Select Case k
        Case 1:  sample = "Ideia grande":      sz = 120: ls = 0.8:  isBold = True:  col = RED:  perCol = BLUE: hasPeriod = True
        Case 2:  sample = "Mega statement":    sz = 80:  ls = 0.9:  isBold = True:  col = BLUE
        Case 3:  sample = "Titulo de pagina":  sz = 60:  ls = 0.9:  isBold = True:  col = RED
        Case 4:  sample = "Label de secao":    sz = 60:  ls = 0.9:  isBold = True:  col = RED
        Case 5:  sample = "Corpo de texto":    sz = 44:  ls = 1.15: isBold = False: col = BLUE
        Case 6:  sample = "Corpo descritivo":  sz = 34:  ls = 0.95: isBold = True:  col = RED
        Case 7:  sample = "Subtitulo de pilar": sz = 28: ls = 1#:   isBold = True:  col = BLUE
        Case 8:  sample = "Texto descritivo":  sz = 24:  ls = 1#:   isBold = False: col = BLUE
        Case 9:  sample = "Corpo de pilar":    sz = 20:  ls = 1.3:  isBold = False: col = BLUE
        Case 10: sample = "Label / eyebrow":   sz = 18:  ls = 1#:   isBold = True:  col = RED
        Case 11: sample = "Legenda / caption": sz = 16:  ls = 1.3:  isBold = False: col = BLUE
    End Select

    Dim sld As Object, shp As Object, tr As Object
    Set sld = ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 60, 840, 220)
    Set tr = shp.TextFrame.TextRange
    If hasPeriod Then tr.Text = sample & "." Else tr.Text = sample
    tr.Font.Name = "Avenir Next"
    If isBold Then tr.Font.Bold = msoTrue Else tr.Font.Bold = msoFalse
    tr.Font.Size = sz
    tr.Font.Color.RGB = col
    tr.ParagraphFormat.LineRuleWithin = msoTrue
    tr.ParagraphFormat.SpaceWithin = ls
    If hasPeriod Then tr.Characters(Len(tr.Text), 1).Font.Color.RGB = perCol
    shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
End Sub
