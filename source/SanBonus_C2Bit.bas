Attribute VB_Name = "SanBonus_C2Bit"
Option Explicit

Sub conv2bit_setup(): bon_ConvToBit.Show: End Sub
Sub conv2bit_pr1(): conv2bit_all 1: End Sub
Sub conv2bit_pr2(): conv2bit_all 2: End Sub
Sub conv2bit_pr3(): conv2bit_all 3: End Sub

Private Sub conv2bit_all(presetLong&)
  Dim sr As ShapeRange, sh As Shape, grWork As Boolean
  Dim s$, a$(), cm As cdrImageType, d As Boolean, p As Boolean, ov As Boolean
  Dim al As cdrAntiAliasingType, t As Boolean, l&
  
  If ActiveDocument Is Nothing Then Exit Sub
  On Error Resume Next
  
  Select Case presetLong
  Case 1: s = GetSetting(macroName, "conv2bit", "Preset1")
  Case 2: s = GetSetting(macroName, "conv2bit", "Preset2")
  Case 3: s = GetSetting(macroName, "conv2bit", "Preset3")
  End Select
  If s = "" Then MsgBox "Preset " & presetLong & " not found", vbInformation: Exit Sub
  
  Set sr = ActiveSelectionRange: l = 0
  If sr.Count < 1 Then _
  MsgBox "Invalid selection!", vbCritical, "info": Exit Sub
  
  a = Split(s, "|")
  
  Select Case a(1)
  Case "Grayscale": cm = cdrGrayscaleImage
  Case "CMYKColor": cm = cdrCMYKColorImage
  Case "RGBColor": cm = cdrRGBColorImage
  End Select
  
  If a(2) = "True" Then d = True Else d = False
  If a(3) = "True" Then p = True Else p = False
  If a(4) = "True" Then ov = True Else ov = False
  If a(6) = "True" Then al = cdrNormalAntiAliasing Else al = cdrNoAntiAliasing
  If a(7) = "True" Then t = True Else t = False
  
  boostStart "Convert To Bitmap (Preset " & presetLong & ")"
  
  grWork = (GetSetting(macroName, "conv2bit", "grWork", "1") = "1")
  If grWork Then
      Application.Status.BeginProgress "Convert Progress", False
      For Each sh In sr
          l = l + 1: Application.Status.Progress = l / sr.Count * 100
          sh.ConvertToBitmapEx cm, d, t, CLng(a(0)), al, p, ov, CLng(a(5))
      Next sh
      Application.Status.EndProgress
  Else
      sr.ConvertToBitmapEx cm, d, t, CLng(a(0)), al, p, ov, CLng(a(5))
  End If

  boostFinish endUndoGroup:=True
  ActiveDocument.ClearSelection
End Sub

