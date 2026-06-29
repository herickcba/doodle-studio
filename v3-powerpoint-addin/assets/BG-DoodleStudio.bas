Attribute VB_Name = "BG_DoodleStudio"
' ============================================================
'  Doodle Studio — Tipografia B+G no ribbon (PowerPoint VBA)
'  Aba "Doodle Studio" (customUI14.xml) -> chama os callbacks
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

' Cores da marca (RGB usa ordem R,G,B)
Private Const C_ROSA As Long = 7036668     ' FC5E6D -> RGB(252, 94,109)
Private Const C_AZUL As Long = 14773315    ' 436AE1 -> RGB( 67,106,225)
Private Const C_BRANCO As Long = 16777215  ' FFFFFF
Private Const FONTE As String = "Avenir Next"

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
    Dim spec As StyleSpec, sel As Object, shp As Object, n As Long
    spec = SpecFor(control.id)
    If Not spec.found Then Exit Sub
    Set sel = ActiveWindow.Selection
    n = 0
    If sel.Type = ppSelectionText Then
        ApplyStyleToRange sel.TextRange, spec, False
        n = 1
    ElseIf sel.Type = ppSelectionShapes Then
        For Each shp In sel.ShapeRange
            n = n + ApplyStyleToShape(shp, spec)
        Next shp
    End If
    If n = 0 Then MsgBox "Selecione um texto ou um objeto com texto.", vbInformation, "Doodle Studio"
End Sub

Public Sub Entrelinha(control As IRibbonControl)
    DoEntrelinhaSelecao
End Sub

Public Sub EntrelinhaTudo(control As IRibbonControl)
    DoEntrelinhaTudo
End Sub

' Fallback p/ Option+F8 (subs sem argumentos aparecem na caixa de Macros)
Public Sub BG_AplicarEntrelinha()
    DoEntrelinhaSelecao
End Sub
Public Sub BG_AplicarEntrelinhaTudo()
    DoEntrelinhaTudo
End Sub

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
    If n = 0 Then MsgBox "Selecione um texto ou um objeto com texto.", vbInformation, "Doodle Studio"
End Sub

Private Sub DoEntrelinhaTudo()
    Dim sld As Object, shp As Object, n As Long
    n = 0
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            n = n + ApplyToShape(shp)
        Next shp
    Next sld
    MsgBox "Entrelinha B+G aplicada em " & n & " caixas de texto, em todos os slides.", vbInformation, "Doodle Studio"
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
