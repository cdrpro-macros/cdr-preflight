Attribute VB_Name = "ctc_infoOsn"
Option Explicit

'====================================================================================
'====================                                            ====================
'====================     Основная функция поиска информации     ====================
'====================                                            ====================
'====================================================================================
Public Function infoDocStart2(myShapeRange As ShapeRange)
        Dim myShape As Shape, myEffect As Effect, ef As Effect, _
        myDpiMin&, myDpiMax&, nodCount&

        On Error GoTo ErrHandler
        'myShapeRange.UngroupAll
        nodCount = GetSetting(macroName, sREGAPPOPT, "NodesCount", "8000")

    'Цикл для ShapeRange ===================
    For Each myShape In myShapeRange
            If myShape.Type = cdrGroupShape Then infoDocStart2 myShape.Shapes.All
            
            'StatusProgress
            mySc = mySc + 1: myStatusProgress mySc, myCountS

            If myShape.CanHaveFill Or myShape.CanHaveOutline Then list_CanFillOutline.Add myShape


        'Select Case Type ===========================
        Select Case myShape.Type
            'Table ============================
            Case cdrCustomShape
              If myShape.Custom.typeID = "Table" Then list_Table.Add myShape

            'Curve ============================
            Case cdrCurveShape
              myFillInfo myShape
              myOutlineInfo myShape
              scanColorSmLim myShape
              If myShape.Curve.Closed = False Then list_noCloseCur.Add myShape
              sCurNod = sCurNod + myShape.Curve.Nodes.Count
              If myShape.Curve.Nodes.Count > nodCount Then list_NodesMax.Add myShape
              If myShape.Transparency.Type <> cdrNoTransparency Then list_EffTransparency.Add myShape

            'Rectangle, Ellipse, Polygon ======
            Case cdrRectangleShape, cdrEllipseShape, cdrPolygonShape
              myFillInfo myShape
              myOutlineInfo myShape
              scanColorSmLim myShape
              If myShape.Transparency.Type <> cdrNoTransparency Then list_EffTransparency.Add myShape

            'Perfect ==========================
            Case cdrPerfectShape: list_PerfSh.Add myShape
              myFillInfo myShape
              myOutlineInfo myShape
              scanColorSmLim myShape
              If myShape.Transparency.Type <> cdrNoTransparency Then list_EffTransparency.Add myShape

            'Symbol ===========================
            Case cdrSymbolShape: list_symbol.Add myShape

            'Texts ============================
            Case cdrTextShape: list_Text.Add myShape
            myTextScanPt myShape
            Select Case myShape.text.Type
                'Case cdrArtisticText: ta = ta + 1
                Case cdrParagraphText
                If myShape.text.Overflow = True Then list_TextOver.Add myShape
            End Select
            If myShape.Effects.Count > 0 Then
                For Each myEffect In myShape.Effects
                If myEffect.Type = cdrTextOnPath Then tcOnP = tcOnP + 1
                Next myEffect
            End If
            myFillInfo myShape
            myOutlineInfo myShape
            scanColorSmLim myShape
            If myShape.Transparency.Type <> cdrNoTransparency Then list_EffTransparency.Add myShape

            'OLEObject =========================
            Case cdrOLEObjectShape: list_OLE.Add myShape
            Select Case myShape.OLE.ProgID
              Case "CorelBarCode.16": sBarCode = sBarCode + 1
              Case "CorelBarCode.15": sBarCode = sBarCode + 1
              Case "CorelBarCode.14": sBarCode = sBarCode + 1
              Case "CorelBarCode.13": sBarCode = sBarCode + 1
              Case "CorelBarCode.12": sBarCode = sBarCode + 1
              Case "CorelBarCode.11": sBarCode = sBarCode + 1
              Case "CorelBarCode.10": sBarCode = sBarCode + 1
              Case "CorelBarCode.9": sBarCode = sBarCode + 1
            End Select

            'Bitmaps ===========================
            Case cdrBitmapShape: list_Allbit.Add myShape
            myBitType myShape
            If myShape.Fill.Type <> cdrNoFill Then bitFill = bitFill + 1
            If myShape.Outline.Type <> cdrNoOutline Then bitOutl = bitOutl + 1
            If myShape.Bitmap.ResolutionX <> myShape.Bitmap.ResolutionY Then list_BitXY.Add myShape

            myDpiMin = GetSetting(macroName, sREGAPPOPT, "myDpiMin", "250")
            myDpiMax = GetSetting(macroName, sREGAPPOPT, "myDpiMax", "320")

            If myShape.Bitmap.ResolutionX < myDpiMin _
            Or myShape.Bitmap.ResolutionY < myDpiMin Then _
            list_MinDPI.Add myShape: GoTo myNextSh

            If myShape.Bitmap.ResolutionX > myDpiMax _
            Or myShape.Bitmap.ResolutionY > myDpiMax Then _
            list_MaxDPI.Add myShape
myNextSh:
            If myShape.Bitmap.CropEnvelopeModified = True Then list_BitCrop.Add myShape
            If myShape.RotationAngle <> 0 Then list_BitRot.Add myShape

            If myShape.OverprintBitmap = True Then list_BitOverpr.Add myShape
            If myShape.Transparency.Type <> cdrNoTransparency Then list_EffTransparency.Add myShape
            If myShape.Bitmap.Transparent = True Then list_BitTr.Add myShape
            If myShape.Bitmap.ExternallyLinked = True Then list_BitLink.Add myShape
'
            'EPS, MeshFlii ======================
            Case cdrEPSShape: list_EPS.Add myShape
            Case cdrMeshFillShape: list_fillMesh.Add myShape
                myOutlineInfo myShape: scanColorSmLim myShape
                If myShape.Transparency.Type <> cdrNoTransparency Then list_EffTransparency.Add myShape
            'ArtisticMedia, Bevel, Blend, Shadow, Contour, Extrud
            Case cdrArtisticMediaGroupShape: list_EffArtisticMedia.Add myShape
            Case cdrCustomEffectGroupShape: list_EffBevel.Add myShape
            Case cdrBlendGroupShape: list_EffBlend.Add myShape
            Case cdrDropShadowGroupShape: list_EffShadow.Add myShape
            Case cdrContourGroupShape: list_EffContour.Add myShape
            Case cdrExtrudeGroupShape: list_EffExtrude.Add myShape

        End Select 'Select Case Type End ========================


            'Effects ==============================================
            If myShape.Effects.Count > 0 Then
            For Each ef In myShape.Effects
                Select Case ef.Type
                Case cdrControlPath: list_ControlPath.Add myShape
                Case cdrDistortion: list_EffDistortion.Add myShape
                Case cdrEnvelope: list_EffEnvelope.Add myShape
                Case cdrLens: list_EffLens.Add myShape
                Case cdrPerspective: list_EffPerspective.Add myShape
                End Select
            Next ef
            End If

            'PowerClips ===========================================
            If Not myShape.PowerClip Is Nothing Then _
            list_PoweClip.Add myShape: _
            infoDocStart2 myShape.PowerClip.Shapes.All

    Next myShape 'Конец цикла для ShapeRange ===================

Exit Function
ErrHandler:
errCount = errCount + 1
If Err.Source <> "AppStatus" Then _
    errStr = errStr & Err.Source & " -- " & Err.Number & _
    " -- " & Err.Description & " -- " & Err.LastDllError & vbCr
Err.Clear
Resume Next
End Function






'====================================================================================
'===================     Определение цветовой модели Bitmaps     ====================
'====================================================================================
Public Function myBitType(s As Shape)
  If myResumeErr Then On Error Resume Next
  Select Case s.Bitmap.Mode
    Case cdrRGBColorImage: list_BitRGB.Add s
    Case cdrPalettedImage: list_BitPal.Add s
    Case cdrLABImage: list_BitLAB.Add s
    Case cdrGrayscaleImage: list_BitGr.Add s
    Case cdrDuotoneImage: list_BitDuo.Add s
    Case cdrCMYKMultiChannelImage: bCMYKm = bCMYKm + 1
    Case cdrCMYKColorImage: list_BitCMYK.Add s
    Case cdrBlackAndWhiteImage: list_BitBW.Add s
    Case cdr16ColorsImage: b16 = b16 + 1
    Case cdrSpotMultiChannelImage: list_BitDevN.Add s
  End Select
End Function
        
        
        
        
        
        
        
'====================================================================================
'===============================     Bed text objects     ===========================
'====================================================================================
Private Sub myTextScanPt(s As Shape)
        Dim tr As TextRange, FF As FountainColor
        Dim tSize As Single, tSizeMin As Single
        If myResumeErr Then On Error Resume Next
        tSizeMin = GetSetting(macroName, sREGAPPOPT, "SmalFontPt", "6")
        tSize = GetSetting(macroName, sREGAPPOPT, "SmalFontColor", "12")
        
        For Each tr In s.text.Story.Words
            If tr.Size < tSizeMin Then list_txtSmalPt.Add s
            If tr.Size < tSize Then
            Select Case tr.Fill.Type
                Case cdrUniformFill
                    If myTextScanPtColor(tr.Fill.UniformColor) Then list_txtSmalCol.Add s
                Case cdrFountainFill
                    For Each FF In tr.Fill.Fountain.Colors
                        If myTextScanPtColor(FF.Color) Then list_txtSmalCol.Add s
                    Next
            End Select
            'tr.Outline
            'tr.Font
            End If
        Next tr
        End Sub
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
        
        
        
        
        
        
'====================================================================================
'========================     Status Progress for INFO     ==========================
'====================================================================================
Private Sub myStatusProgress(c&, cc&)
        Dim l&
        If myResumeErr Then On Error Resume Next
        l = c / cc * 100
        If l <= 100 Then Application.Status.Progress = l Else Application.Status.Progress = 100
        End Sub
