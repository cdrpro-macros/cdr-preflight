Attribute VB_Name = "cp_Convert"
Option Explicit

Private cfg As New clsConfig

Public _
    convBitMode As cdrImageType, _
    convColorType As cdrColorType
    'conv_multPaDoc As Boolean

Public _
    conv_tToCur As Boolean, _
    conv_sBevelBit As Boolean, _
    conv_sContArtMedia As Boolean, _
    conv_sDropShaToBit As Boolean, _
    conv_sLens2bit As Boolean, _
    conv_sBlendB As Boolean, _
    conv_sDimSep As Boolean, _
    conv_sDropShaMulty As Boolean, _
    conv_PowerToBit As Boolean, _
    conv_sSymbToShape As Boolean, _
    conv_sContBr As Boolean, _
    conv_sDistirtToCur As Boolean, _
    conv_sEnvelToCur As Boolean
    
Public _
    conv_meshToBit As Boolean, _
    conv_FtFillToBit As Boolean

Public _
    conv_bCMYK As Boolean, _
    conv_bCMYKm As Boolean, _
    conv_bSpot As Boolean, _
    conv_bRGB As Boolean, _
    conv_b16col As Boolean, _
    conv_bBW As Boolean, _
    conv_bD As Boolean, _
    conv_bG As Boolean, _
    conv_bL As Boolean, _
    conv_bP As Boolean, _
    conv_M_dpi&, _
    conv_CG_dpi&, _
    conv_bCrop As Boolean, _
    conv_bAngle0 As Boolean, _
    conv_bOverprint As Boolean, _
    conv_bLinkBr As Boolean

Public _
    conv_bProfile As Boolean, _
    conv_bTransp As Boolean, _
    conv_baa&, _
    conv_bOverPrBlack As Boolean, _
    conv_OverPrBlackLim&

Public _
    conv_sOutScaleNo As Boolean, _
    conv_sMiterLimit As Boolean, _
    conv_sMiterLimitValue#, _
    conv_cb_OutlineWEdit$, _
    conv_c_OutlineWEdit As Boolean, _
    conv_sOverprint As Boolean, _
    conv_sOverprintBlack As Boolean, _
    conv_PatternToBit As Boolean, _
    conv_TexturToBit As Boolean, _
    conv_PSFillToBit As Boolean

Public _
    conv_cb_ColorType$, _
    conv_cb_ChaResol As Boolean, _
    conv_DPIbox$, _
    conv_myDpiMin&, _
    conv_myDpiMax&

Public _
    conv_sBW As Boolean, _
    conv_sCMYK As Boolean, _
    conv_sCMY As Boolean, _
    conv_sGr As Boolean, _
    conv_sHLS As Boolean, _
    conv_sHSB As Boolean, _
    conv_sLAB As Boolean, _
    conv_sPanH As Boolean, _
    conv_sReg As Boolean, _
    conv_sRGB As Boolean, _
    conv_sChangeUserColor As Boolean, _
    conv_sSpot As Boolean, _
    conv_sYIQ As Boolean, _
    conv_sUserInk As Boolean
    
Public _
    conv_cmVisible As Boolean, _
    conv_cmPrint As Boolean, _
    conv_cmEnable As Boolean, _
    conv_cmNVisible As Boolean
    'conv_cmNPrint As Boolean, _
    'conv_cmNEnable As Boolean
    
    





Function cmLoadConvPresets() As String
  Dim s$: s = "<select id=""ConvPresetsListSel"" onchange=""cmChangeConvPreset()""><option value=""_none"">Default</option>"
  Dim c&, i&, presName$, selected$
  c = GetSetting(macroName, "Convert", "PresetsCount", 0)
  selected = GetSetting(macroName, "Convert", "PresetsLast", "0")
  For i = 1 To c
    presName = GetSetting(macroName, "Convert", "Presets" & i & "Name")
    If presName <> "" Then
      If CLng(selected) = i Then
        s = s & "<option value=""" & i & """ selected=""selected"">" & presName & "</option>"
      Else
        s = s & "<option value=""" & i & """>" & presName & "</option>"
      End If
    End If
  Next i
  cmLoadConvPresets = s & "</select>"
End Function


Function cmChangeConvPreset(id$)
  SaveSetting macroName, "Convert", "PresetsLast", IIf(id = "_none", "0", id)
End Function




'====================================================================================
'===========================                             ============================
'===========================        Start Convert        ============================
'===========================                             ============================
'====================================================================================
Function cmConverter(presetID$) As String
  If myResumeErr Then On Error Resume Next
  
  If IsDo(True) = False Then Exit Function
  
  Dim myShapeRange As ShapeRange
  '====================================================
  Set myDoc = ActiveDocument
  Set myOldPage = ActivePage
  Set myMasterPage = myDoc.MasterPage
  Set myShapeRange = New ShapeRange
  
  Select Case GetSetting(macroName, sREGAPPOPT, "cb_Unit", "millimeters")
    Case "millimeters": myUnitWork = cdrMillimeter
    Case "points": myUnitWork = cdrPoint
    Case Else: MsgBox "No Unit   ", vbCritical, "Warning": Exit Function
  End Select
  myUnit = myDoc.Unit
  myDoc.Unit = myUnitWork
  
  'Загружаем настройки
  If presetID = "_none" Then
    Call conv_LoadSetting
  Else
    Call conv_LoadSetting2(presetID)
    Call conv_LoadSetting
  End If
  
  myBeforeWork                       'обнуляем переменные
  
  boostStart "ConvertToPrint"
  Status.BeginProgress "Convert...", False

  'myUnit = myDoc.Unit
  'myDoc.Unit = myUnitWork

    Set myShapeRange = ActiveSelectionRange
    If myShapeRange.Count = 0 Then myConvDoc: myDoc.ClearSelection Else myConvSel myShapeRange

  'myDoc.Unit = myUnit
  myOldPage.Activate
  
  Application.Status.EndProgress
  boostFinish endUndoGroup:=True
  
  myDoc.Unit = myUnit
  
  Set myDoc = Nothing
  Set myOldPage = Nothing
  Set myMasterPage = Nothing
  Set myShapeRange = Nothing
  
  '====================================================
  'If (GetSetting(macroName, sREGAPPOPT, "ErrLogSave", "0")) = "1" Then If errCount > 0 Then myConvErrLogWr
  
  cmConverter = "<p class=""alert"">Operation is completed!</p>"
End Function






'====================================================================================
'==============================    Convert NotSel     ===============================
'====================================================================================
Private Sub myConvDoc()
    myLayerEnable
    If myResumeErr Then On Error Resume Next
    ctc_Convert_finde.myFindeShapes "OLE_EPS_Symb"
        myConv1et
    If conv_PowerToBit Then
        ctc_Convert_finde.myFindeShapes "PowerClip"
        myConv2et
        End If
    ctc_Convert_finde.myFindeShapes "Effect"
        myConv3et
    ctc_Convert_finde.myFindeShapes "Txt_Bit"
        myConv4et
    ctc_Convert_finde.myFindeShapes "Fill_Outline"
        myConv5et
    'Восстанавливаем слои после работы в них
    myLayerDesable
End Sub
        
        
'====================================================================================
'============================    Convert in Select     ==============================
'====================================================================================
Private Sub myConvSel(sr As ShapeRange)
  Dim sr2 As ShapeRange
  Set sr2 = New ShapeRange
  If myResumeErr Then On Error Resume Next
  
  Set sr2 = myOldPage.Shapes.All
  sr2.RemoveRange sr
  ctc_Convert_finde.myFindeShapesEd sr, "OLE_EPS_Symb"
  myConv1et
  
  If conv_PowerToBit Then
  Set sr = myOldPage.Shapes.All
  sr.RemoveRange sr2
  ctc_Convert_finde.myFindeShapesEd sr, "PowerClip"
  myConv2et
  End If
  
  Set sr = myOldPage.Shapes.All
  sr.RemoveRange sr2
  ctc_Convert_finde.myFindeShapesEd sr, "Effect"
  myConv3et
  
  Set sr = myOldPage.Shapes.All
  sr.RemoveRange sr2
  ctc_Convert_finde.myFindeShapesEd sr, "Txt_Bit"
  myConv4et
  
  Set sr = myOldPage.Shapes.All
  sr.RemoveRange sr2
  ctc_Convert_finde.myFindeShapesEd sr, "Fill_Outline"
  myConv5et
      
  Set sr = myOldPage.Shapes.All
  sr.RemoveRange sr2
  myDoc.ClearSelection
  sr.CreateSelection
End Sub
            
            
Private Sub myConv1et()
  If conv_sSymbToShape Then myConvSymb list_symbol
  'MsgBox "OLE " & list_OLE.Count
  'MsgBox "EPS " & list_EPS.Count
  If conv_sDimSep Then myConvDimension list_Dimension
  Status.Progress = 20
End Sub

Private Sub myConv2et()
  myConvPowClip list_PoweClip
  Status.Progress = 30
End Sub

Private Sub myConv3et()
  myConvDS list_EffShadow
  If conv_sBlendB Then myConvBlend list_EffBlend
  If conv_sContBr Then myConvContour list_EffContour
  Status.Progress = 40
  If conv_sDistirtToCur Then myConvDist list_EffDistortion
  If conv_sEnvelToCur Then myConvEnvelop list_EffEnvelope
  If conv_sBevelBit Then myConvBevel list_EffBevel
  Status.Progress = 50
  If conv_sContArtMedia Then myConvArtMedia list_EffArtisticMedia
  If conv_sLens2bit Then myConvLens list_EffLens
  Status.Progress = 60
End Sub

Private Sub myConv4et()
  If conv_tToCur Then myConvText list_Text 'Else myConvText2 list_Text
  Status.Progress = 70
  myBitConv list_Allbit
  Status.Progress = 80
End Sub

Private Sub myConv5et()
  myFillConv list_CanFill
  myOutlineConv list_CanOutline
  myOutlineConvPr list_OutlineProbl
  Status.Progress = 90
  If conv_meshToBit Then myMeshFillConv list_fillMesh
  Status.Progress = 100
End Sub







'====================================================================================
'========================                                   =========================
'========================         Конверторы типов          =========================
'========================                                   =========================
'====================================================================================


' Symbol to object
Private Sub myConvSymb(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            s.Symbol.RevertToShapes
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
' Dimension
Private Sub myConvDimension(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            s.Separate
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
' PowerClip To Bitmap
Private Sub myConvPowClip(sr As ShapeRange)
        Dim s As Shape, c&, eff As Effect, srEf As ShapeRange
        Set srEf = New ShapeRange: c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            If s.Effects.Count > 0 Then
                For Each eff In s.Effects
                    If eff.Type = cdrDropShadow Then
                        srEf.Add s
                        myConvDS srEf
                        'Если тень останется, то её нужно отделить, чтобы она не пропала.
                        Set srEf = eff.Separate
                        Exit For
                    End If
                Next
            End If
            myConvToBit s
            'c = c + 1: myProgress c, sr.Count
        Next s
        End Sub
        
        
        
        
' DropShadow
Private Sub myConvDS(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            's.Layer.Page.Activate
            myConvDS2 s
'            If s.PowerClipParent Is Nothing Then
'                myConvDS2 s
'            Else
'                's.PowerClipParent.PowerClip.EnterEditMode
'                myConvDS2 s
'                's.PowerClipParent.PowerClip.LeaveEditMode
'            End If
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
Private Sub myConvDS2(s As Shape)
        Dim sr2 As ShapeRange, eff As Effect
        Set sr2 = New ShapeRange
        If myResumeErr Then On Error Resume Next
        For Each eff In s.Effects
            If conv_sDropShaToBit Then
                If conv_sDropShaMulty Then _
                If eff.DropShadow.Color.Name = "Black" Then _
                If eff.DropShadow.MergeMode <> cdrMergeMultiply Then _
                eff.DropShadow.MergeMode = cdrMergeMultiply

                Set sr2 = eff.Separate
                myConvToBit sr2(1)
                
            Else
                If conv_sDropShaMulty Then _
                If eff.DropShadow.Color.Name = "Black" Then _
                If eff.DropShadow.MergeMode <> cdrMergeMultiply Then _
                eff.DropShadow.MergeMode = cdrMergeMultiply
            End If
        Next eff
        End Sub
        
        
' Blend Br
Private Sub myConvBlend(sr As ShapeRange)
        Dim s As Shape, c&, eff As Effect
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            For Each eff In s.Effects
            If eff.Type = cdrBlend Then eff.Separate
            Next eff
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
' Contour
Private Sub myConvContour(sr As ShapeRange)
        Dim s As Shape, c&, eff As Effect
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            For Each eff In s.Effects
            If eff.Type = cdrContour Then eff.Separate
            Next eff
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
' Distortion
Private Sub myConvDist(sr As ShapeRange)
        Dim s As Shape, c&, eff As Effect
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            For Each eff In s.Effects
            If eff.Type = cdrDistortion Then s.ConvertToCurves
            Next eff
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
' Envelop
Private Sub myConvEnvelop(sr As ShapeRange)
        Dim s As Shape, c&, eff As Effect
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            For Each eff In s.Effects
            If eff.Type = cdrEnvelope Then s.ConvertToCurves
            Next eff
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
' Bevel
Private Sub myConvBevel(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            's.Layer.Page.Activate
            myConvBevel2 s
'            If s.PowerClipParent Is Nothing Then
'                myConvBevel2 s
'            Else
'                s.PowerClipParent.PowerClip.EnterEditMode
'                myConvBevel2 s
'                s.PowerClipParent.PowerClip.LeaveEditMode
'            End If
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
Private Sub myConvBevel2(s As Shape)
        If myResumeErr Then On Error Resume Next
        myConvToBit s
        End Sub
        
        
        
'ArtisticMedia
Private Sub myConvArtMedia(sr As ShapeRange)
        Dim s As Shape, c&
        Dim s1 As Shape
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            Set s1 = s.GetLinkedShapes(cdrLinkAllConnections).Item(1)
            s.Separate
            s1.Delete
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
'Lens Convert to Bit
Private Sub myConvLens(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            's.Layer.Page.Activate
            myConvToBit s
            'c = c + 1
            'myProgress c, sr.Count
        Next s
        End Sub
        
        
        
        
' Text Conv
Private Sub myConvText(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            s.ConvertToCurves
            'c = c + 1: 'myProgress c, sr.Count
        Next s
        End Sub

        
' Bitmap
Private Sub myBitConv(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            myBitConv2 s
            'c = c + 1: 'myProgress c, sr.Count
        Next s
        End Sub
        
Private Sub myBitConv2(s As Shape)
        If myResumeErr Then On Error Resume Next
        If s.Type = cdrBitmapShape Then
            If s.Bitmap.ExternallyLinked Then
                If conv_bLinkBr Then s.Bitmap.ResolveLink
                If s.Bitmap.ExternallyLinked = False Then GoTo AfterBr
            Else
AfterBr:
                If conv_bOverprint Then s.OverprintBitmap = False
                If conv_bCrop Then
                    If s.Bitmap.CropEnvelopeModified Then s.Bitmap.Crop
                End If
                If conv_bAngle0 Then If s.RotationAngle <> 0 Then myConvToBit s: Exit Sub
                myBitModeConv s
                
                'Меняем разрешение ===============================
                If conv_cb_ChaResol Then
                    If s.Bitmap.Mode = cdrBlackAndWhiteImage Then
                        If s.Bitmap.ResolutionX <> conv_M_dpi Or s.Bitmap.ResolutionY <> conv_M_dpi Then _
                        s.Bitmap.Resample , , False, conv_M_dpi, conv_M_dpi
                    Else
                        If conv_DPIbox = "Auto (min/max)" Then
                            If s.Bitmap.ResolutionX < conv_myDpiMin Or s.Bitmap.ResolutionY < conv_myDpiMin Then
                            s.Bitmap.Resample , , conv_baa, conv_myDpiMin, conv_myDpiMin
                            ElseIf s.Bitmap.ResolutionX > conv_myDpiMax Or s.Bitmap.ResolutionY > conv_myDpiMax Then
                            s.Bitmap.Resample , , conv_baa, conv_myDpiMax, conv_myDpiMax
                            End If
                        ElseIf conv_DPIbox = "User" Then
                            If s.Bitmap.ResolutionX <> conv_CG_dpi Or s.Bitmap.ResolutionY <> conv_CG_dpi Then _
                            s.Bitmap.Resample , , conv_baa, conv_CG_dpi, conv_CG_dpi
                        End If
                    End If
                End If 'Меняем разрешение ========================
                
            End If
        End If
        End Sub
Private Sub myBitModeConv(s As Shape)
        On Error GoTo myBitErr
        Select Case s.Bitmap.Mode
          Case cdrRGBColorImage: If conv_bRGB Then s.Bitmap.ConvertTo convBitMode
          Case cdrPalettedImage: If conv_bP Then s.Bitmap.ConvertTo convBitMode
          Case cdrLABImage: If conv_bL Then s.Bitmap.ConvertTo convBitMode
          Case cdrGrayscaleImage: If conv_bG Then s.Bitmap.ConvertTo convBitMode
          Case cdrDuotoneImage: If conv_bD Then s.Bitmap.ConvertTo convBitMode
          Case cdrCMYKMultiChannelImage: If conv_bCMYKm Then s.Bitmap.ConvertTo convBitMode
          Case cdrCMYKColorImage: If conv_bCMYK Then s.Bitmap.ConvertTo convBitMode
          Case cdrBlackAndWhiteImage: If conv_bBW Then s.Bitmap.ConvertTo convBitMode
          Case cdr16ColorsImage: If conv_b16col Then s.Bitmap.ConvertTo convBitMode
          Case cdrSpotMultiChannelImage: If conv_bSpot Then s.Bitmap.ConvertTo convBitMode
        End Select
        Exit Sub
myBitErr:
        MsgBox Err.Number & vbCr & Err.Description
        Err.Clear
        Resume Next
        End Sub
        
        
        
        
' MeshFill
Private Sub myMeshFillConv(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        On Error Resume Next
        For Each s In sr
            's.Layer.Page.Activate
            myConvToBit s
            'c = c + 1: 'myProgress c, sr.Count
        Next s
        End Sub
        

        
' Fill
Private Sub myFillConv(sr As ShapeRange)
        Dim s As Shape, c&, fc As FountainColor
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            Select Case s.Fill.Type
                Case cdrUniformFill
                fillAndOutlineColor s.Fill.UniformColor
                
                If conv_sOverprint Then
                    If conv_sOverprintBlack Then
                        If s.Fill.UniformColor.Name(True) <> "C:0 M:0 Y:0 K:100" Then _
                        s.OverprintFill = False
                    Else
                        s.OverprintFill = False
                    End If
                End If
                
                Case cdrFountainFill
                If conv_sOverprint Then s.OverprintFill = False
                If conv_FtFillToBit Then
                    myConvToBit s
                Else
                    For Each fc In s.Fill.Fountain.Colors
                        fillAndOutlineColor fc.Color
                    Next fc
                End If
                
                Case cdrPatternFill: If conv_PatternToBit Then myConvToBit s: GoTo nextShape
                Case cdrTextureFill: If conv_TexturToBit Then myConvToBit s: GoTo nextShape
                Case cdrPostscriptFill: If conv_PSFillToBit Then myConvToBit s: GoTo nextShape
            End Select
nextShape:
        'c = c + 1: 'myProgress c, sr.Count
        Next s
        End Sub
        
        
' Outline
Private Sub myOutlineConv(sr As ShapeRange)
        Dim s As Shape, c&
        Dim myOutlW#
        c = 0
        On Error GoTo myErr
        For Each s In sr
            If conv_c_OutlineWEdit Then
                myOutlW = val(Replace(GetSetting(macroName, sREGAPPOPT, "OutlineWidthMin", "0.0762"), ",", "."))
                Select Case conv_cb_OutlineWEdit
                Case "Enlarge to..."
                    If s.Outline.Width < myOutlW Then s.Outline.Width = myOutlW
                Case "Remove"
                    If s.Outline.Width < myOutlW Then
                        s.Outline.ScaleWithShape = False
                        s.Outline.Type = cdrNoOutline
                        GoTo myNext
                    End If
                End Select
            End If
            
            fillAndOutlineColor s.Outline.Color
            If conv_sOutScaleNo Then s.Outline.ScaleWithShape = False
    
            If conv_sOverprint Then
                If conv_sOverprintBlack Then
                    If s.Outline.Color.Name(True) <> "C:0 M:0 Y:0 K:100" Then _
                    s.OverprintOutline = False
                Else
                    s.OverprintOutline = False
                End If
            End If
            
            'миттер лимит
            If conv_sMiterLimit Then _
            If s.Outline.MiterLimit <> conv_sMiterLimitValue Then s.Outline.MiterLimit = conv_sMiterLimitValue
myNext:
        'c = c + 1: 'myProgress c, sr.Count
        Next s
        Exit Sub
myErr:
        If Err.Number = -2147467259 Then
            Err.Clear
            Resume Next
        End If
        End Sub
        
        
' Outline Pr
Private Sub myOutlineConvPr(sr As ShapeRange)
        Dim s As Shape, c&
        c = 0
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            If s.Outline.Type = cdrNoOutline And s.Outline.ScaleWithShape = True Then _
                s.Outline.ScaleWithShape = False: s.Outline.Type = cdrNoOutline
            'c = c + 1: 'myProgress c, sr.Count
        Next s
        End Sub
        
        
        
Private Sub fillAndOutlineColor(myColor As Color)
        Select Case myColor.Type
            Case cdrColorRGB: If conv_sRGB Then fillAndOutlineColorConv myColor
            Case cdrColorBlackAndWhite: If conv_sBW Then fillAndOutlineColorConv myColor
            Case cdrColorGray: If conv_sGr Then fillAndOutlineColorConv myColor
            Case cdrColorLab: If conv_sLAB Then fillAndOutlineColorConv myColor
            
            Case cdrColorPantone: If conv_sSpot Then fillAndOutlineColorConv2 myColor
            Case cdrColorSpot: If conv_sSpot Then fillAndOutlineColorConv2 myColor
            Case cdrColorPantoneHex: If conv_sPanH Then fillAndOutlineColorConv2 myColor
            
            Case cdrColorCMYK: If conv_sCMYK Then fillAndOutlineColorConv myColor
            Case cdrColorCMY: If conv_sCMY Then fillAndOutlineColorConv myColor
            Case cdrColorRegistration: If conv_sReg Then fillAndOutlineColorConv myColor
            Case cdrColorHSB: If conv_sHSB Then fillAndOutlineColorConv myColor
            Case cdrColorHLS: If conv_sHLS Then fillAndOutlineColorConv myColor
            Case cdrColorYIQ: If conv_sYIQ Then fillAndOutlineColorConv myColor
            Case cdrColorUserInk: If conv_sUserInk Then fillAndOutlineColorConv myColor
        End Select
        End Sub
Private Sub fillAndOutlineColorConv(myColor As Color)
        If conv_sChangeUserColor Then If ReplaceUserColor(myColor) Then Exit Sub
        Select Case convColorType
            Case cdrColorCMYK: myColor.ConvertToCMYK
            Case cdrColorGray: myColor.ConvertToGray
            Case cdrColorRGB: myColor.ConvertToRGB
        End Select
        End Sub
Private Sub fillAndOutlineColorConv2(myColor As Color)
        If conv_sChangeUserColor Then If ReplaceUserColor(myColor) Then Exit Sub
        Select Case convColorType
            Case cdrColorCMYK: myColor.ConvertToRGB: myColor.ConvertToCMYK
            Case cdrColorGray: myColor.ConvertToGray
            Case cdrColorRGB: myColor.ConvertToRGB
        End Select
        End Sub
Private Function ReplaceUserColor(c As Color) As Boolean
        Dim i&, a$()
        ReplaceUserColor = False
        For i = 1 To GetSetting(macroName, "ColorReplacer", "ColorListCount", 0)
            a = Split(GetSetting(macroName, "ColorReplacer", "ColorList" & i), "|")
            cfg.SetFindStr a(0)
            cfg.SetReplaceStr a(1)
            If c.Name(True) = cfg.clrFind.Name(True) Then
                c.CopyAssign cfg.clrReplace
                ReplaceUserColor = True
                Exit Function
            End If
        Next i
        End Function





            
' ConvertToBitmapEx =====================================================
Private Sub myConvToBit(s As Shape)
        'If conv_multPaDoc Then
        Dim sp As Shape
        If Not s.PowerClipParent Is Nothing Then
            Set sp = s.PowerClipParent
            sp.PowerClip.EnterEditMode
            s.ConvertToBitmapEx convBitMode, False, conv_bTransp, conv_CG_dpi, _
            conv_baa, conv_bProfile, conv_bOverPrBlack, conv_OverPrBlackLim
            sp.PowerClip.LeaveEditMode
        Else
            s.Layer.Activate
            s.ConvertToBitmapEx convBitMode, False, conv_bTransp, conv_CG_dpi, _
            conv_baa, conv_bProfile, conv_bOverPrBlack, conv_OverPrBlackLim
        End If
        Set sp = Nothing
        End Sub







'====================================================================================
'==============================      Load Settings      =============================
'====================================================================================
Private Sub conv_LoadSetting()
  If myResumeErr Then On Error Resume Next
  
  conv_cb_ColorType = GetSetting(macroName, "Convert", "cb_ColorType", "CMYK Color")
  Select Case conv_cb_ColorType
  Case "CMYK Color"
      convBitMode = cdrCMYKColorImage
      convColorType = cdrColorCMYK
  Case "Grey Color"
      convBitMode = cdrGrayscaleImage
      convColorType = cdrColorGray
  Case "RGB Color"
      convBitMode = cdrRGBColorImage
      convColorType = cdrColorRGB
  End Select
  
  conv_tToCur = GetSetting(macroName, "Convert", "tToCur", "1")

  conv_sLens2bit = GetSetting(macroName, "Convert", "sLens2bit", "0")
  conv_sBevelBit = GetSetting(macroName, "Convert", "sBevelBit", "1")
  conv_sContArtMedia = GetSetting(macroName, "Convert", "sContArtMedia", "1")
  conv_sDropShaToBit = GetSetting(macroName, "Convert", "sDropShaToBit", "1")
  conv_sDropShaMulty = GetSetting(macroName, "Convert", "sDropShaMulty", "0")
  conv_sBlendB = GetSetting(macroName, "Convert", "sBlendB", "1")
  conv_sContBr = GetSetting(macroName, "Convert", "sContBr", "1")
  conv_sDistirtToCur = GetSetting(macroName, "Convert", "sDistirtToCur", "1")
  conv_sEnvelToCur = GetSetting(macroName, "Convert", "sEnvelToCur", "1")
  
  conv_PowerToBit = GetSetting(macroName, "Convert", "PowerToBit", "0")
  conv_sSymbToShape = GetSetting(macroName, "Convert", "sSymbToShape", "1")
  conv_sDimSep = GetSetting(macroName, "Convert", "sDimSep", "1")
  
  conv_FtFillToBit = GetSetting(macroName, "Convert", "FtFillToBit", "0")
  conv_meshToBit = GetSetting(macroName, "Convert", "meshToBit", "1")
  
  conv_DPIbox = GetSetting(macroName, "Convert", "myDPIbox", "Auto (min/max)")
  
  conv_myDpiMin = GetSetting(macroName, sREGAPPOPT, "myDpiMin", "250")
  conv_myDpiMax = GetSetting(macroName, sREGAPPOPT, "myDpiMax", "320")
  
  conv_baa = GetSetting(macroName, "Convert", "bAAliasing", "1")
  conv_bProfile = GetSetting(macroName, "Convert", "bProfile", "1")
  conv_bTransp = GetSetting(macroName, "Convert", "bTransp", "1")
  conv_bOverPrBlack = GetSetting(macroName, "Convert", "bOverPrBlack", "0")
  
  conv_bCrop = GetSetting(macroName, "Convert", "bCrop", "1")
  conv_bAngle0 = GetSetting(macroName, "Convert", "bAngle0", "1")
  conv_bOverprint = GetSetting(macroName, "Convert", "bOverprint", "1")
  conv_bLinkBr = GetSetting(macroName, "Convert", "bLinkBr", "0")
  
  conv_bCMYK = GetSetting(macroName, "Convert", "bCMYK", "0")
  conv_bCMYKm = GetSetting(macroName, "Convert", "bCMYKm", "0")
  conv_bSpot = GetSetting(macroName, "Convert", "bSpot", "0")
  conv_b16col = GetSetting(macroName, "Convert", "b16col", "1")
  conv_bBW = GetSetting(macroName, "Convert", "bBW", "0")
  conv_bD = GetSetting(macroName, "Convert", "bD", "1")
  conv_bG = GetSetting(macroName, "Convert", "bG", "0")
  conv_bL = GetSetting(macroName, "Convert", "bL", "1")
  conv_bP = GetSetting(macroName, "Convert", "bP", "1")
  conv_bRGB = GetSetting(macroName, "Convert", "bRGB", "1")
  
  conv_cb_ChaResol = GetSetting(macroName, "Convert", "cb_ChaResol", "1")
  conv_M_dpi = GetSetting(macroName, "Convert", "M_dpi", "600")
  conv_CG_dpi = GetSetting(macroName, "Convert", "CG_dpi", "300")
  
  conv_sOutScaleNo = GetSetting(macroName, "Convert", "sOutScaleNo", "1")
  conv_sMiterLimit = GetSetting(macroName, "Convert", "sMiterLimit", "0")
  conv_sMiterLimitValue = GetSetting(macroName, "Convert", "sMiterLimitValue", "45")
  conv_sOverprint = GetSetting(macroName, "Convert", "sOverprint", "1")
  conv_sOverprintBlack = GetSetting(macroName, "Convert", "sOverprintBlack", "0")
  conv_OverPrBlackLim = GetSetting(macroName, "Convert", "tb_OverPrBlackLim", "95")
  
  conv_PatternToBit = GetSetting(macroName, "Convert", "PatternToBit", "1")
  conv_TexturToBit = GetSetting(macroName, "Convert", "TexturToBit", "1")
  conv_PSFillToBit = GetSetting(macroName, "Convert", "PSFillToBit", "1")
              
  conv_sCMYK = GetSetting(macroName, "Convert", "sCMYK", "0")
  conv_sBW = GetSetting(macroName, "Convert", "sBW", "0")
  conv_sCMY = GetSetting(macroName, "Convert", "sCMY", "0")
  conv_sGr = GetSetting(macroName, "Convert", "sGr", "0")
  conv_sHLS = GetSetting(macroName, "Convert", "sHLS", "1")
  conv_sHSB = GetSetting(macroName, "Convert", "sHSB", "1")
  conv_sLAB = GetSetting(macroName, "Convert", "sLAB", "1")
  conv_sPanH = GetSetting(macroName, "Convert", "sPanH", "1")
  conv_sReg = GetSetting(macroName, "Convert", "sReg", "1")
  conv_sRGB = GetSetting(macroName, "Convert", "sRGB", "1")
  conv_sSpot = GetSetting(macroName, "Convert", "sSpot", "1")
  conv_sYIQ = GetSetting(macroName, "Convert", "sYIQ", "1")
  conv_sUserInk = GetSetting(macroName, "Convert", "sUserInk", "1")
  conv_sChangeUserColor = GetSetting(macroName, "Convert", "sChangeUserColor", "0")
  
  conv_cmVisible = GetSetting(macroName, "Convert", "cmVisible", "1")
  conv_cmPrint = GetSetting(macroName, "Convert", "cmPrint", "1")
  conv_cmEnable = GetSetting(macroName, "Convert", "cmEnable", "1")
  conv_cmNVisible = GetSetting(macroName, "Convert", "cmNVisible", "0")
  conv_cb_OutlineWEdit = GetSetting(macroName, "Convert", "cb_OutlineWEdit", "Enlarge to...")
  conv_c_OutlineWEdit = GetSetting(macroName, "Convert", "c_OutlineWEdit", "0")
End Sub

'====================================================================================
'==============================      Save Settings      =============================
'====================================================================================
Private Sub conv_LoadSetting2(myPresN$)
  Dim a$(): a = Split(GetSetting(macroName, "Convert", "Presets" & myPresN), "|")
  
  SaveSetting macroName, "Convert", "cb_ColorType", a(1)
  SaveSetting macroName, "Convert", "tToCur", a(2)
  SaveSetting macroName, "Convert", "PowerToBit", a(3)
  SaveSetting macroName, "Convert", "sSymbToShape", a(4)
  SaveSetting macroName, "Convert", "sLens2bit", a(5)
  SaveSetting macroName, "Convert", "sBlendB", a(6)
  SaveSetting macroName, "Convert", "sContBr", a(7)
  SaveSetting macroName, "Convert", "sBevelBit", a(8)
  SaveSetting macroName, "Convert", "sContArtMedia", a(9)
  SaveSetting macroName, "Convert", "sEnvelToCur", a(10)
  'sExtrudBr = a(11)
  SaveSetting macroName, "Convert", "sDistirtToCur", a(12)
  'sPerspToCur = a(13)
  SaveSetting macroName, "Convert", "sDropShaToBit", a(14)
  SaveSetting macroName, "Convert", "sDropShaMulty", a(15)
  SaveSetting macroName, "Convert", "sCMYK", a(16)
  SaveSetting macroName, "Convert", "sCMY", a(17)
  SaveSetting macroName, "Convert", "sRGB", a(18)
  SaveSetting macroName, "Convert", "sBW", a(19)
  SaveSetting macroName, "Convert", "sGr", a(20)
  SaveSetting macroName, "Convert", "sLAB", a(21)
  SaveSetting macroName, "Convert", "sSpot", a(22)
  SaveSetting macroName, "Convert", "sPanH", a(23)
  SaveSetting macroName, "Convert", "sReg", a(24)
  SaveSetting macroName, "Convert", "sHSB", a(25)
  SaveSetting macroName, "Convert", "sHLS", a(26)
  SaveSetting macroName, "Convert", "sYIQ", a(27)
  SaveSetting macroName, "Convert", "sOutScaleNo", a(28)
  SaveSetting macroName, "Convert", "sOverprint", a(29)
  SaveSetting macroName, "Convert", "sOverprintBlack", a(30)
  SaveSetting macroName, "Convert", "meshToBit", a(31)
  SaveSetting macroName, "Convert", "PatternToBit", a(32)
  SaveSetting macroName, "Convert", "TexturToBit", a(33)
  SaveSetting macroName, "Convert", "PSFillToBit", a(34)
  SaveSetting macroName, "Convert", "cb_ChaResol", a(35)
  SaveSetting macroName, "Convert", "myDPIbox", a(36)
  SaveSetting macroName, "Convert", "CG_dpi", a(37)
  SaveSetting macroName, "Convert", "M_dpi", a(38)
  SaveSetting macroName, "Convert", "bTransp", a(39)
  SaveSetting macroName, "Convert", "bAAliasing", a(40)
  SaveSetting macroName, "Convert", "bProfile", a(41)
  SaveSetting macroName, "Convert", "bOverPrBlack", a(42)
  SaveSetting macroName, "Convert", "tb_OverPrBlackLim", a(43)
  SaveSetting macroName, "Convert", "bCMYK", a(44)
  SaveSetting macroName, "Convert", "bCMYKm", a(45)
  SaveSetting macroName, "Convert", "bBW", a(46)
  SaveSetting macroName, "Convert", "bG", a(47)
  SaveSetting macroName, "Convert", "bRGB", a(48)
  SaveSetting macroName, "Convert", "bL", a(49)
  SaveSetting macroName, "Convert", "bP", a(50)
  SaveSetting macroName, "Convert", "b16col", a(51)
  SaveSetting macroName, "Convert", "bD", a(52)
  SaveSetting macroName, "Convert", "bSpot", a(53)
  SaveSetting macroName, "Convert", "bCrop", a(54)
  SaveSetting macroName, "Convert", "bAngle0", a(55)
  SaveSetting macroName, "Convert", "bOverprint", a(56)
  SaveSetting macroName, "Convert", "bLinkBr", a(57)
  
  If CLng(a(0)) >= 2 Then
      SaveSetting macroName, "Convert", "FtFillToBit", a(58)
  Else
      SaveSetting macroName, "Convert", "FtFillToBit", "0"
  End If
  If CLng(a(0)) >= 3 Then
      SaveSetting macroName, "Convert", "sMiterLimit", a(59)
      SaveSetting macroName, "Convert", "sMiterLimitValue", a(60)
      SaveSetting macroName, "Convert", "sUserInk", a(61)
      SaveSetting macroName, "Convert", "sChangeUserColor", a(62)
      SaveSetting macroName, "Convert", "cmVisible", a(63)
      SaveSetting macroName, "Convert", "cmPrint", a(64)
      SaveSetting macroName, "Convert", "cmEnable", a(65)
      SaveSetting macroName, "Convert", "cmNVisible", a(66)
      SaveSetting macroName, "Convert", "cb_OutlineWEdit", a(67)
      SaveSetting macroName, "Convert", "c_OutlineWEdit", a(68)
  Else
      SaveSetting macroName, "Convert", "sMiterLimit", "0"
      SaveSetting macroName, "Convert", "sMiterLimitValue", "45"
      SaveSetting macroName, "Convert", "sUserInk", "0"
      SaveSetting macroName, "Convert", "sChangeUserColor", "0"
      SaveSetting macroName, "Convert", "cmVisible", "1"
      SaveSetting macroName, "Convert", "cmPrint", "1"
      SaveSetting macroName, "Convert", "cmEnable", "1"
      SaveSetting macroName, "Convert", "cmNVisible", "0"
      SaveSetting macroName, "Convert", "cb_OutlineWEdit", "Enlarge to..."
      SaveSetting macroName, "Convert", "c_OutlineWEdit", "0"
  End If
End Sub
