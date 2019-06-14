Attribute VB_Name = "SanBonus_color"
Option Explicit
Private sCol As cdrColorType, sImg As cdrImageType, sCou&, c&


Sub convToRGB()
  sCol = cdrColorRGB: sImg = cdrRGBColorImage: convTo "Convert To RGB"
  End Sub
Sub convToCMYK()
  sCol = cdrColorCMYK: sImg = cdrCMYKColorImage: convTo "Convert To CMYK"
  End Sub
Sub convToGRAY()
  Dim c&
  sCol = cdrColorGray: sImg = cdrGrayscaleImage: convTo "Convert To GRAY"
  End Sub
Sub convToGrayEx()
  If VersionMajor < 15 Then MsgBox "Only for CorelDRAW X5 and above!"
  Dim c&
  c = ColorManager.ColorEngine
  ColorManager.ColorEngine = 4
  sCol = cdrColorGray: sImg = cdrGrayscaleImage: convTo "Convert To GRAY"
  ColorManager.ColorEngine = c
  End Sub



Private Sub convTo(ts$)
        Dim sr As ShapeRange: Set sr = New ShapeRange
        On Error Resume Next
        If ActiveDocument Is Nothing Then Exit Sub
        Set sr = ActiveSelectionRange
        
        If sr.Count = 0 Then
            MsgBox "Not selected !    ", vbInformation, " info"
        Else
            boostStart ts
            sCou = 0: c = 0: shCount sr
            Application.Status.BeginProgress "Convert...", False
            convTo2 sr
            Application.Status.EndProgress
            
            ActiveDocument.ClearSelection
            sr.AddToSelection
            boostFinish endUndoGroup:=True
        End If
        End Sub
Private Sub convTo2(sr As ShapeRange)
        Dim s As Shape, fc As FountainColor
        On Error Resume Next
        For Each s In sr
            If s.Type = cdrGroupShape Then
                convTo2 s.Shapes.All
            Else
                Select Case s.Type
                Case cdrBitmapShape: If s.Bitmap.Mode <> sImg Then s.Bitmap.ConvertTo sImg
                Case Else
                    If s.CanHaveFill Then
                    Select Case s.Fill.Type
                    Case cdrUniformFill: fillAndOutlineColor s.Fill.UniformColor
                    Case cdrFountainFill
                        For Each fc In s.Fill.Fountain.Colors
                        fillAndOutlineColor fc.Color
                        Next fc
                    End Select
                    End If
                    If s.CanHaveOutline Then _
                        If s.Outline.Type <> cdrNoOutline Then fillAndOutlineColor s.Outline.Color
                End Select
                c = c + 1: Application.Status.Progress = c / sCou * 100
            End If
            If Not s.PowerClip Is Nothing Then convTo2 s.PowerClip.Shapes.All
        Next s
        End Sub


Private Sub fillAndOutlineColor(myColor2 As Color)
        If myColor2.Type <> sCol Then
        Select Case sCol
            Case cdrColorRGB: myColor2.ConvertToRGB
            Case cdrColorCMYK
                If myColor2.Type = cdrColorPantone _
                Or myColor2.Type = cdrColorSpot Then _
                myColor2.ConvertToRGB: myColor2.ConvertToCMYK Else myColor2.ConvertToCMYK
            Case cdrColorGray: myColor2.ConvertToGray
        End Select
        End If
        End Sub
        
        
Private Sub shCount(sr As ShapeRange)
        Dim s As Shape
        On Error Resume Next
        For Each s In sr
            If Not s.PowerClip Is Nothing Then shCount s.PowerClip.Shapes.All
            If s.Type = cdrGroupShape Then shCount s.Shapes.All Else sCou = sCou + 1
        Next s
        End Sub

