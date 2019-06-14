Attribute VB_Name = "ctc_Fill_Outline"
Option Explicit

'====================================================================================
'================================     Fill INFO     =================================
'====================================================================================
Public Function myFillInfo(myShape As Shape)
        Dim myFountainColor As FountainColor
        If myResumeErr Then On Error Resume Next
        Select Case myShape.Fill.Type
            Case cdrFountainFill: list_FountainFill.Add myShape
                'If myShape.Fill.Fountain.MidPoint <> 50 Then sFonFillMP = sFonFillMP + 1
                For Each myFountainColor In myShape.Fill.Fountain.Colors
                    myFillColorInfo myFountainColor.Color, myFountainColor, myShape
                Next
            Case cdrHatchFill: shhf = shhf + 1
            Case cdrPatternFill: list_fillPattern.Add myShape
            Case cdrPostscriptFill: list_fillPS.Add myShape
            Case cdrTextureFill: list_fillTexture.Add myShape
            Case cdrUniformFill: shuf = shuf + 1
                myFillColorInfo myShape.Fill.UniformColor
            Case cdrNoFill: shnf = shnf + 1
        End Select
        If myShape.OverprintFill = True Then oPrinf = oPrinf + 1
        End Function
        
        
Public Function myFillColorInfo(myColor As Color, _
                                Optional ByVal ffColor As FountainColor = Nothing, _
                                Optional ByVal myShape As Shape = Nothing)
                                
        Dim til&, c&
        til = GetSetting(macroName, sREGAPPOPT, "TILFill", "280")
        
        With myColor
            Select Case .Type
                Case cdrColorBlackAndWhite: scolBW = scolBW + 1
                Case cdrColorCMY: scolCMY = scolCMY + 1
                Case cdrColorCMYK
                    scolCMYK = scolCMYK + 1
                    uColorTIL = .CMYKCyan + .CMYKMagenta + .CMYKYellow + .CMYKBlack
                    If uColorTIL >= til Then uColorTIL300 = uColorTIL300 + 1
                    'If uColorTIL = 400 Then scolCMYK100 = scolCMYK100 + 1
                    If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then
                        scolCMYK100 = scolCMYK100 + 1
                    End If
                Case cdrColorGray: scolGray = scolGray + 1
                Case cdrColorHLS: scolHLS = scolHLS + 1
                Case cdrColorHSB: scolHSB = scolHSB + 1
                Case cdrColorLab: scolLab = scolLab + 1
                Case cdrColorMixed: scolMix = scolMix + 1
                Case cdrColorMultiChannel: scolMulti = scolMulti + 1
                Case cdrColorPantone: scolPan = scolPan + 1
                Case cdrColorPantoneHex: scolPanH = scolPanH + 1
                Case cdrColorRegistration: scolReg = scolReg + 1
                Case cdrColorRGB: scolRGB = scolRGB + 1
                Case cdrColorSpot: scolSpot = scolSpot + 1
                Case cdrColorUserInk: scolUserInk = scolUserInk + 1
                Case cdrColorYIQ: scolYIQ = scolYIQ + 1
            End Select
        End With
        
        If Not ffColor Is Nothing Then
          If ffColor.Opacity < 255 Then list_EffTransparency.Add myShape
        End If
        End Function
'====================================================================================
'==============================     Outline INFO     ================================
'====================================================================================
Public Function myOutlineInfo(myShape As Shape)
        Dim til&, myOutlW$
        If myResumeErr Then On Error Resume Next
        til = GetSetting(macroName, sREGAPPOPT, "TILFill", "280")
        myOutlW = GetSetting(macroName, sREGAPPOPT, "OutlineWidthMin", "0.0762")
        
        Select Case myShape.Outline.Type
            Case cdrOutline: sOuLineN = sOuLineN + 1
            Case cdrEnhancedOutline: sOuLineEnh = sOuLineEnh + 1
            Case cdrNoOutline
                If myShape.Outline.ScaleWithShape = True Then list_OutlineProbl.Add myShape
        End Select
        
        With myShape.Outline.Color
            Select Case .Type
                Case cdrColorBlackAndWhite: scolBW = scolBW + 1
                Case cdrColorCMY: scolCMY = scolCMY + 1
                Case cdrColorCMYK: scolCMYK = scolCMYK + 1
                    uColorTIL = .CMYKCyan + .CMYKMagenta + .CMYKYellow + .CMYKBlack
                    If uColorTIL >= til Then uColorTIL300 = uColorTIL300 + 1
                    'If uColorTIL = 400 Then sOuLineCMYK100 = sOuLineCMYK100 + 1
                    If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then
                        scolCMYK100 = scolCMYK100 + 1
                    End If
                Case cdrColorGray: scolGray = scolGray + 1
                Case cdrColorHLS: scolHLS = scolHLS + 1
                Case cdrColorHSB: scolHSB = scolHSB + 1
                Case cdrColorLab: scolLab = scolLab + 1
                Case cdrColorMixed: scolMix = scolMix + 1
                Case cdrColorMultiChannel: scolMulti = scolMulti + 1
                Case cdrColorPantone: scolPan = scolPan + 1
                Case cdrColorPantoneHex: scolPanH = scolPanH + 1
                Case cdrColorRegistration: scolReg = scolReg + 1
                Case cdrColorRGB: scolRGB = scolRGB + 1
                Case cdrColorSpot: scolSpot = scolSpot + 1
                Case cdrColorUserInk: scolUserInk = scolUserInk + 1
                Case cdrColorYIQ: scolYIQ = scolYIQ + 1
            End Select
        End With
        
        If myShape.OverprintOutline = True Then oPrino = oPrino + 1
        If myShape.Outline.Type <> cdrNoOutline Then
            If myShape.Outline.Width <= val(Replace(myOutlW, ",", ".")) And myShape.Outline.Width > 0 Then list_OutlineMin.Add myShape
            If myShape.Outline.ScaleWithShape = True Then list_OutLineScal.Add myShape
        End If
        End Function
'====================================================================================
'===========================     Find Min ColorLim     ==============================
'====================================================================================
Public Sub scanColorSmLim(s As Shape)
        Dim FF As FountainColor
        If myResumeErr Then On Error Resume Next
        If s.CanHaveFill Then
            Select Case s.Fill.Type
            Case cdrFountainFill
                For Each FF In s.Fill.Fountain.Colors
                    If FF.Color.Type = cdrColorCMYK Then _
                    If scanColorSmLim2(FF.Color) = True Then _
                        list_ColorSmalLim.Add s: sColorSmalLim = sColorSmalLim + 1
                    If FF.Color.Type = cdrColorSpot _
                    Or FF.Color.Type = cdrColorPantone _
                    Or FF.Color.Type = cdrColorPantoneHex Then _
                    If scanTintSmLim(FF.Color) = True Then _
                        list_ColorSmalLim.Add s: sColorSmalLim = sColorSmalLim + 1
                Next
            Case cdrUniformFill
                If s.Fill.UniformColor.Type = cdrColorCMYK Then _
                If scanColorSmLim2(s.Fill.UniformColor) = True Then _
                    list_ColorSmalLim.Add s: sColorSmalLim = sColorSmalLim + 1
                If s.Fill.UniformColor.Type = cdrColorSpot _
                Or s.Fill.UniformColor.Type = cdrColorPantone _
                Or s.Fill.UniformColor.Type = cdrColorPantoneHex Then _
                If scanTintSmLim(s.Fill.UniformColor) = True Then _
                    list_ColorSmalLim.Add s: sColorSmalLim = sColorSmalLim + 1
            End Select
        End If
        
        If s.CanHaveOutline Then
            If s.Outline.Type <> cdrNoOutline Then _
            If s.Outline.Color.Type = cdrColorCMYK Then _
            If scanColorSmLim2(s.Outline.Color) = True Then _
                list_ColorSmalLim.Add s: sColorSmalLim = sColorSmalLim + 1
            If s.Outline.Color.Type = cdrColorSpot _
            Or s.Outline.Color.Type = cdrColorPantone _
            Or s.Outline.Color.Type = cdrColorPantoneHex Then _
            If scanTintSmLim(s.Outline.Color) = True Then _
                list_ColorSmalLim.Add s: sColorSmalLim = sColorSmalLim + 1
        End If
        End Sub
Private Function scanColorSmLim2(c As Color) As Boolean
        Dim cl&
        cl = GetSetting(macroName, sREGAPPOPT, "myMinColor", "10")
        If c.CMYKCyan > 0 And c.CMYKCyan < cl Then scanColorSmLim2 = True: Exit Function
        If c.CMYKMagenta > 0 And c.CMYKMagenta < cl Then scanColorSmLim2 = True: Exit Function
        If c.CMYKYellow > 0 And c.CMYKYellow < cl Then scanColorSmLim2 = True: Exit Function
        If c.CMYKBlack > 0 And c.CMYKBlack < cl Then scanColorSmLim2 = True
        End Function
Private Function scanTintSmLim(c As Color) As Boolean
        Dim cl&
        cl = GetSetting(macroName, sREGAPPOPT, "myMinColor", "10")
        If c.Tint < cl Then scanTintSmLim = True
        End Function
