Attribute VB_Name = "BG_Tipografia"
' ============================================================
'  B+G Tipografia — entrelinha (PowerPoint VBA add-in)
'  A API de add-in nao seta entrelinha; o VBA seta. A escolha de
'  tipografia (fonte/tamanho/cor/ponto) e feita na barra lateral
'  do Doodle Studio; este modulo cuida SO da entrelinha.
'
'  BG_AplicarEntrelinha      -> aplica a entrelinha B+G a SELECAO
'                               (texto, objeto[s] e grupos).
'  BG_AplicarEntrelinhaTudo  -> aplica em TODOS os slides, em todas
'                               as caixas de texto (inclui grupos).
'  Rode por Option+F8 (ou botao no Quick Access Toolbar).
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

' Aplica a uma shape; entra em grupos recursivamente. Retorna nº de caixas tocadas.
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

' === Entrelinha B+G na SELECAO (texto, objeto[s], grupos) ===
Public Sub BG_AplicarEntrelinha()
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
    If n = 0 Then MsgBox "Selecione um texto ou um objeto com texto.", vbInformation, "B+G Tipografia"
End Sub

' === Entrelinha B+G em TODOS os slides (todas as caixas, incl. grupos) ===
Public Sub BG_AplicarEntrelinhaTudo()
    Dim sld As Object, shp As Object, n As Long
    n = 0
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            n = n + ApplyToShape(shp)
        Next shp
    Next sld
    MsgBox "Entrelinha B+G aplicada em " & n & " caixas de texto, em todos os slides.", vbInformation, "B+G Tipografia"
End Sub
