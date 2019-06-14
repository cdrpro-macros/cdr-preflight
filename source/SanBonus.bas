Attribute VB_Name = "SanBonus"
Option Explicit


Sub ShowTIL()
  uShowTIL.Show
End Sub


'========================================================================================
Sub applyMultiply()
  Dim s As Shape, r As New ShapeRange
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then MsgBox "Invalid selection!", vbCritical, "info": Exit Sub
  boostStart "Multiply Apply"
  applyMultiply2 ActiveSelectionRange
  ActiveDocument.ClearSelection
  boostFinish endUndoGroup:=True
End Sub
Private Sub applyMultiply2(r As ShapeRange)
  Dim s As Shape
  On Error Resume Next
  
  For Each s In r
      If s.Type = cdrGroupShape Then
          applyMultiply2 s.Shapes.All
      Else
          s.Transparency.ApplyUniformTransparency 0
          With s.Transparency
              .AppliedTo = cdrApplyToFillAndOutline
              .MergeMode = cdrMergeMultiply
          End With
      End If
      If Not s.PowerClip Is Nothing Then applyMultiply2 s.PowerClip.Shapes.All
  Next s
End Sub



'========================================================================================
'========================================================================================
'========================================================================================
Sub SelectTxt2Curve()
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then MsgBox "Invalid selection!", vbCritical, "info": Exit Sub
  boostStart "Convert Text to Curves"
  SelectTxt2Curve2 ActiveSelectionRange
  ActiveDocument.ClearSelection
  boostFinish endUndoGroup:=True
End Sub

Private Sub SelectTxt2Curve2(r As ShapeRange)
  Dim s As Shape, f As New Fill
  On Error Resume Next
  Set f = ActiveDocument.CreateFill()
  For Each s In r
      If s.Type = cdrGroupShape Then SelectTxt2Curve2 s.Shapes.All Else _
          If s.Type = cdrTextShape Then s.ConvertToCurves
      If Not s.PowerClip Is Nothing Then SelectTxt2Curve2 s.PowerClip.Shapes.All
  Next s
End Sub





'========================================================================================
'========================================================================================
'========================================================================================
Sub fitFrame2Content()
  Dim s As Shape, d As Document
  If ActiveDocument Is Nothing Then Exit Sub
  Set d = ActiveDocument
  
  If ActiveSelectionRange.Count <> 1 Then MsgBox "Invalid selection!", vbCritical, "info": Exit Sub
  Set s = ActiveShape
  If s.Type <> cdrTextShape Then MsgBox "Only for Paragraph Text!", vbCritical, "info": Exit Sub
  If s.text.Type <> cdrParagraphText Then MsgBox "Only for Paragraph Text!", vbCritical, "info": Exit Sub
  
  d.Unit = cdrMillimeter
  d.ReferencePoint = cdrTopLeft
  
  boostStart "Fit Frame To Content"
  If s.text.Overflow = True Then
  fitFrameB s
  Else
  fitFrameT s
  End If
  boostFinish endUndoGroup:=True
  End Sub
Private Sub fitFrameT(s As Shape)
  Dim sW#, sh#
  s.GetSize sW, sh: s.SetSize sW, sh - 1
  If s.text.Overflow = False Then fitFrameT s Else s.GetSize sW, sh: s.SetSize sW, sh + 1: Exit Sub
  End Sub
Private Sub fitFrameB(s As Shape)
  Dim sW#, sh#
  s.GetSize sW, sh: s.SetSize sW, sh + 1
  If s.text.Overflow = True Then fitFrameB s Else Exit Sub
  End Sub


    
'========================================================================================
'========================================================================================
'========================================================================================
Sub CreateGuidelineFromNodes()
    Dim s As Shape, sr As New ShapeRange, n As Node, c&
    Dim nr As NodeRange, nr2 As NodeRange
    If ActiveDocument Is Nothing Then Exit Sub
    Set sr = ActiveSelectionRange: c = 0
    If sr.Count > 2 Or sr.Count < 1 Then Beep: Exit Sub
    On Error Resume Next
    
    If sr.Count = 1 Then
        If sr(1).Type = cdrCurveShape Then Set nr = sr(1).Curve.Selection
        If nr.Count <> 2 Then Beep: GoTo myExit
        ActiveLayer.CreateGuide nr.FirstNode.PositionX, nr.FirstNode.PositionY, _
        nr.LastNode.PositionX, nr.LastNode.PositionY
    Else
        If sr(1).Type = cdrCurveShape Then Set nr = sr(1).Curve.Selection
            If nr.Count < 1 Or nr.Count > 2 Then Beep: GoTo myExit
            If nr.Count = 2 Then _
            ActiveLayer.CreateGuide nr.FirstNode.PositionX, nr.FirstNode.PositionY, _
            nr.LastNode.PositionX, nr.LastNode.PositionY: GoTo myExit
            
        If sr(2).Type = cdrCurveShape Then Set nr2 = sr(2).Curve.Selection
            If nr2.Count <> 1 Then Beep: GoTo myExit
            ActiveLayer.CreateGuide nr(1).PositionX, nr(1).PositionY, _
            nr2(1).PositionX, nr2(1).PositionY
    End If
    
myExit:
    Set nr = Nothing
    Set nr2 = Nothing
    Set sr = Nothing
    ActiveDocument.ClearSelection
End Sub
'========================================================================================
'========================================================================================
'========================================================================================
Sub CreateGuidelineNodAndRotCenter()
    Dim s As Shape, sr As New ShapeRange, n As Node
    Dim nr As NodeRange
    If ActiveDocument Is Nothing Then Exit Sub
    Set sr = ActiveSelectionRange
    If sr.Count <> 1 Then Beep: Exit Sub
    On Error Resume Next

    If sr(1).Type = cdrCurveShape Then Set nr = sr(1).Curve.Selection
    If nr.Count <> 1 Then Beep: Exit Sub
    Set s = ActiveLayer.CreateGuide(nr.FirstNode.PositionX, nr.FirstNode.PositionY, _
    sr(1).RotationCenterX, sr(1).RotationCenterY)
    
    If (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then _
        s.Guide.SetPointAndAngle nr.FirstNode.PositionX, nr.FirstNode.PositionY, s.Guide.Angle + 90

    Set nr = Nothing
    Set sr = Nothing
    ActiveDocument.ClearSelection
End Sub
'========================================================================================
'========================================================================================
'========================================================================================
Sub CreateGuidelineNod_VG()
    Dim sr As New ShapeRange, n As Node
    Dim nr As NodeRange
    If ActiveDocument Is Nothing Then Exit Sub
    Set sr = ActiveSelectionRange
    If sr.Count <> 1 Then Beep: Exit Sub
    On Error Resume Next
    
    If sr(1).Type = cdrCurveShape Then Set nr = sr(1).Curve.Selection
    If nr.Count <> 1 Then Beep: Exit Sub
    
    If (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then
        ActiveLayer.CreateGuide nr.FirstNode.PositionX, nr.FirstNode.PositionY, _
        nr.FirstNode.PositionX + 1, nr.FirstNode.PositionY
    Else
        ActiveLayer.CreateGuide nr.FirstNode.PositionX, nr.FirstNode.PositionY, _
        nr.FirstNode.PositionX, nr.FirstNode.PositionY + 1
    End If

    Set nr = Nothing
    Set sr = Nothing
    ActiveDocument.ClearSelection
End Sub



'========================================================================================
'========================================================================================
'========================================================================================
Sub CreateGuidelineNod_Angle()
    Dim s As Shape, sr As New ShapeRange, n As Node, c#
    Dim nr As NodeRange
    If ActiveDocument Is Nothing Then Exit Sub
    Set sr = ActiveSelectionRange: c = 0
    If sr.Count <> 1 Then Beep: Exit Sub
    On Error Resume Next
    
    If sr(1).Type = cdrCurveShape Then Set nr = sr(1).Curve.Selection
    If nr.Count <> 1 Then Beep: Exit Sub
    
    c = InputBox("Angle", "Angle", 0)
    If c <= 0 Or c > 360 Then Exit Sub
    
    Set s = ActiveLayer.CreateGuide(nr.FirstNode.PositionX, nr.FirstNode.PositionY, _
    nr.FirstNode.PositionX + 1, nr.FirstNode.PositionY)
    s.Guide.SetPointAndAngle nr.FirstNode.PositionX, nr.FirstNode.PositionY, c
    

    Set nr = Nothing
    Set sr = Nothing
    ActiveDocument.ClearSelection
End Sub















