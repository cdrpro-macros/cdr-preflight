Attribute VB_Name = "info_functions"
Option Explicit


'====================================================================================
'=================         Инфо для списка в главном окне       =====================
'====================================================================================

Public Function san_ColorModeView(sr As ShapeRange, srNew As ShapeRange, cType As cdrColorType)
        Dim s As Shape, fc As FountainColor, mySrAdd As Boolean
        On Error Resume Next
        For Each s In sr
        mySrAdd = False
        If s.Type = cdrGroupShape Then
            san_ColorModeView s.Shapes.All, srNew, cType
        Else
            If s.CanHaveFill Then
                Select Case s.Fill.Type
                Case cdrUniformFill: If s.Fill.UniformColor.Type = cType Then srNew.Add s: mySrAdd = True
                Case cdrFountainFill
                    For Each fc In s.Fill.Fountain.Colors
                        If fc.Color.Type = cType Then srNew.Add s: mySrAdd = True
                    Next
                End Select
            End If

            If s.CanHaveOutline Then
                If s.Outline.Color.Type = cType Then
                If mySrAdd = False Then srNew.Add s
                End If
            End If
        End If
        If Not s.PowerClip Is Nothing Then san_ColorModeView s.PowerClip.Shapes.All, srNew, cType
        Next s
        End Function


'Overprints =======================================================
Public Function OverprintFillView2(sr As ShapeRange, sr2 As ShapeRange)
        Dim s As Shape
        On Error Resume Next
        For Each s In sr
            If Not s.PowerClip Is Nothing Then OverprintFillView2 s.PowerClip.Shapes.All, sr2
    
            If s.Type = cdrGroupShape Then
                OverprintFillView2 s.Shapes.All, sr2
            Else
                If s.CanHaveFill And s.OverprintFill Then sr2.Add s
                If s.CanHaveOutline And s.OverprintOutline Then sr2.Add s
            End If
        Next s
        End Function



'Найти все SPOT заливки и обводки =======================================================
Public Function SpotView(sr As ShapeRange, sr2 As ShapeRange)
        Dim s As Shape, fc As FountainColor, mySrAdd As Boolean
        On Error Resume Next
        For Each s In sr
            mySrAdd = False
            If s.Type = cdrGroupShape Then
                SpotView s.Shapes.All, sr2
            Else
                If s.CanHaveFill Then
                    Select Case s.Fill.Type
                    Case cdrUniformFill
                        If s.Fill.UniformColor.Type = cdrColorPantone Or _
                        s.Fill.UniformColor.Type = cdrColorSpot Then sr2.Add s: mySrAdd = True
                    Case cdrFountainFill
                        For Each fc In s.Fill.Fountain.Colors
                            If fc.Color.Type = cdrColorPantone Or _
                            fc.Color.Type = cdrColorSpot Then sr2.Add s: mySrAdd = True
                        Next
                    End Select
                End If
    
                If s.CanHaveOutline Then
                    If s.Outline.Color.Type = cdrColorPantone Or _
                    s.Outline.Color.Type = cdrColorSpot Then
                    If mySrAdd = False Then sr2.Add s
                    End If
                End If
            End If
            If Not s.PowerClip Is Nothing Then SpotView s.PowerClip.Shapes.All, sr2
        Next s
        End Function


'Найти все ТИЛ =======================================================
Public Function san_TILView(sr As ShapeRange, sr2 As ShapeRange)
        Dim s As Shape, fc As FountainColor, mySrAdd As Boolean
        Dim til&
        On Error Resume Next
        
        til = GetSetting(macroName, sREGAPPOPT, "TILFill", "280")
        For Each s In sr
            mySrAdd = False
            If s.Type = cdrGroupShape Then
                san_TILView s.Shapes.All, sr2
            Else
                If s.CanHaveFill Then
                    Select Case s.Fill.Type
                    Case cdrUniformFill: If s.Fill.UniformColor.Type = cdrColorCMYK Then _
                    If s.Fill.UniformColor.CMYKCyan + s.Fill.UniformColor.CMYKMagenta + _
                    s.Fill.UniformColor.CMYKYellow + s.Fill.UniformColor.CMYKBlack >= til Then _
                    sr2.Add s: mySrAdd = True
                    Case cdrFountainFill
                        For Each fc In s.Fill.Fountain.Colors
                            If fc.Color.Type = cdrColorCMYK Then _
                            If fc.Color.CMYKCyan + fc.Color.CMYKMagenta + _
                            fc.Color.CMYKYellow + fc.Color.CMYKBlack >= til Then _
                            sr2.Add s: mySrAdd = True
                        Next
                    End Select
                End If
    
                If s.CanHaveOutline Then
                    If s.Outline.Color.Type = cdrColorCMYK Then
                        If s.Outline.Color.CMYKCyan + s.Outline.Color.CMYKMagenta + _
                        s.Outline.Color.CMYKYellow + s.Outline.Color.CMYKBlack >= til Then _
                        If mySrAdd = False Then sr2.Add s
                    End If
                End If
            End If
            If Not s.PowerClip Is Nothing Then san_TILView s.PowerClip.Shapes.All, sr2
        Next s
        End Function


'Найти все CMYK100 =======================================================
Public Function san_CMYK100View(sr As ShapeRange, sr2 As ShapeRange)
        Dim s As Shape, fc As FountainColor, mySrAdd As Boolean
        On Error Resume Next
        For Each s In sr
            mySrAdd = False
            If s.Type = cdrGroupShape Then
                san_CMYK100View s.Shapes.All, sr2
            Else
                If s.CanHaveFill Then
                    Select Case s.Fill.Type
                    Case cdrUniformFill
                        With s.Fill.UniformColor
                            If .Type = cdrColorCMYK Then _
                            If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then _
                            sr2.Add s: mySrAdd = True
                        End With
                    Case cdrFountainFill
                        For Each fc In s.Fill.Fountain.Colors
                            With fc.Color
                                If .Type = cdrColorCMYK Then _
                                If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then _
                                sr2.Add s: mySrAdd = True
                            End With
                        Next
                    End Select
                End If
    
                If s.CanHaveOutline Then
                    With s.Outline.Color
                        If .Type = cdrColorCMYK Then _
                        If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then _
                        If mySrAdd = False Then sr2.Add s
                    End With
                End If
            End If
            If Not s.PowerClip Is Nothing Then san_CMYK100View s.PowerClip.Shapes.All, sr2
        Next s
        End Function
        
        
        
'====================================================================================
'===============================     Small Font Size     ============================
'====================================================================================
Public Function mySmallFontSize$(s As Shape)
        Dim tr As TextRange, tSizeMin As Single, c&
        tSizeMin = GetSetting(macroName, sREGAPPOPT, "SmalFontPt", "6")
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each tr In s.text.Story.Words
            If tr.Size < tSizeMin Then c = c + 1
        Next tr
        mySmallFontSize = " (" & c & " words)"
        End Function
        
        
'====================================================================================
'==============================     Small Font Color     ============================
'====================================================================================
Public Function mySmallFontColor$(s As Shape)
        Dim tr As TextRange, tSize As Single, c&, FF As FountainColor
        c = 0: tSize = GetSetting(macroName, sREGAPPOPT, "SmalFontColor", "12")
        If myResumeErr Then On Error Resume Next
        For Each tr In s.text.Story.Words
            If tr.Size < tSize Then
            Select Case tr.Fill.Type
                Case cdrUniformFill
                    If myTextScanPtColor(tr.Fill.UniformColor) Then c = c + 1: GoTo myNext
                Case cdrFountainFill
                    For Each FF In tr.Fill.Fountain.Colors
                        If myTextScanPtColor(FF.Color) Then c = c + 1: GoTo myNext
                    Next
            End Select
            End If
myNext:
        Next tr
        mySmallFontColor = " (" & c & " words)"
        End Function
Private Function myTextScanPtColor(c As Color) As Boolean
        Dim l&
        l = 0
        If myResumeErr Then On Error Resume Next
        If c.Type = cdrColorCMYK Then
        With c
            If .CMYKCyan > 0 Then l = l + 1
            If .CMYKMagenta > 0 Then l = l + 1
            If .CMYKYellow > 0 Then l = l + 1
            If .CMYKBlack > 0 Then l = l + 1
        End With
        End If
        If l > 1 Then myTextScanPtColor = True
        End Function

