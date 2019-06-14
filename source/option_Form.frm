VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} option_Form 
   Caption         =   "Converter Options"
   ClientHeight    =   9885.001
   ClientLeft      =   42
   ClientTop       =   406
   ClientWidth     =   11732
   OleObjectBlob   =   "option_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "option_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit








Private Sub cm_showOptColorRep_Click()
        ColorReplacer.Show
        End Sub
Private Sub cm_showOptColorRep_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_showOptColorRep.SpecialEffect = fmSpecialEffectSunken
        End Sub
Private Sub cm_showOptColorRep_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_showOptColorRep.SpecialEffect = fmSpecialEffectRaised
        End Sub
Private Sub Frame7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_showOptColorRep.SpecialEffect = fmSpecialEffectEtched
        End Sub
        

Private Sub cm_showLayersOptions_Click()
        OptionsInfo_Layers.Show
        End Sub
Private Sub cm_showLayersOptions_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_showLayersOptions.SpecialEffect = fmSpecialEffectSunken
        End Sub
Private Sub cm_showLayersOptions_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_showLayersOptions.SpecialEffect = fmSpecialEffectRaised
        End Sub
Private Sub Frame9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_showLayersOptions.SpecialEffect = fmSpecialEffectEtched
        End Sub






Private Sub Command_OK_Click()
        Unload Me
        End Sub
Private Sub Command_OK_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectSunken
        End Sub
Private Sub Command_OK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectRaised
        End Sub
        
        
        
        
        
        
        
        
        

'====================================================================================
'=================================     Save Preset     ==============================
'====================================================================================
Private Sub cm_presSave_Click()
  If List2.text = "" Then Exit Sub
  
  Dim msg&
  msg = MsgBox("Are you sure you want to save a preset?   ", vbQuestion + vbOKCancel, macroName & " " & macroVersion)
  If msg <> 1 Then Exit Sub
  
  Dim strPres$
  strPres = strPresetDo
  
  Dim i&: i = List2.ListIndex + 1
  SaveSetting macroName, "Convert", "Presets" & i, strPres
End Sub
Private Sub cm_presSave_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  cm_presSave.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub cm_presSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  cm_presSave.SpecialEffect = fmSpecialEffectRaised
End Sub

        
'====================================================================================
'==================================     Add Preset     ==============================
'====================================================================================
Private Sub cm_addPres_Click()
  Dim strPres$, strPresN$, strPresCount$, c&, s$
  c = List2.ListCount + 1
  
  strPresN = InputBox("Enter name for new preset", "Create a New Preset")
  If strPresN = "" Then Exit Sub

  strPres = strPresetDo
  
  SaveSetting macroName, "Convert", "Presets" & c, strPres
  SaveSetting macroName, "Convert", "Presets" & c & "Name", strPresN
  
  List2.AddItem strPresN
  SaveSetting macroName, "Convert", "PresetsCount", c
End Sub
Private Sub cm_addPres_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  cm_addPres.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub cm_addPres_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  cm_addPres.SpecialEffect = fmSpecialEffectRaised
End Sub
        
        
Private Function strPresetDo() As String
  Dim cmVisible As Boolean, cmPrint As Boolean, cmEnable As Boolean, cmNVisible As Boolean
  
  cmVisible = (GetSetting(macroName, "Convert", "cmVisible", "1") = "1")
  cmPrint = (GetSetting(macroName, "Convert", "cmPrint", "1") = "1")
  cmEnable = (GetSetting(macroName, "Convert", "cmEnable", "1") = "1")
  cmNVisible = (GetSetting(macroName, "Convert", "cmNVisible", "0") = "1")
  
  strPresetDo = convPresVer & "|" & cb_ColorType & "|" & IIf(tToCur, "1", "0") & "|" & IIf(PowerToBit, "1", "0") & "|" & IIf(sSymbToShape, "1", "0") & "|" & IIf(sLens2bit, "1", "0") & "|" & IIf(sBlendB, "1", "0") & "|" & IIf(sContBr, "1", "0") & "|" & IIf(sBevelBit, "1", "0") & "|" & _
  IIf(sContArtMedia, "1", "0") & "|" & IIf(sEnvelToCur, "1", "0") & "|" & IIf(sExtrudBr, "1", "0") & "|" & IIf(sDistirtToCur, "1", "0") & "|" & IIf(sPerspToCur, "1", "0") & "|" & IIf(sDropShaToBit, "1", "0") & "|" & IIf(sDropShaMulty, "1", "0") & "|" & _
  IIf(sCMYK, "1", "0") & "|" & IIf(sCMY, "1", "0") & "|" & IIf(sRGB, "1", "0") & "|" & IIf(sBW, "1", "0") & "|" & IIf(sGr, "1", "0") & "|" & IIf(sLAB, "1", "0") & "|" & IIf(sSpot, "1", "0") & "|" & IIf(sPanH, "1", "0") & "|" & IIf(sReg, "1", "0") & "|" & _
  IIf(sHSB, "1", "0") & "|" & IIf(sHLS, "1", "0") & "|" & IIf(sYIQ, "1", "0") & "|" & IIf(sOutScaleNo, "1", "0") & "|" & IIf(sOverprint, "1", "0") & "|" & IIf(sOverprintBlack, "1", "0") & "|" & IIf(meshToBit, "1", "0") & "|" & IIf(PatternToBit, "1", "0") & "|" & _
  IIf(TexturToBit, "1", "0") & "|" & IIf(PSFillToBit, "1", "0") & "|" & IIf(cb_ChaResol, "1", "0") & "|" & myDPIbox & "|" & CG_dpi & "|" & M_dpi & "|" & IIf(bTransp, "1", "0") & "|" & IIf(bAAliasing, "1", "0") & "|" & IIf(bProfile, "1", "0") & "|" & _
  IIf(bOverPrBlack, "1", "0") & "|" & tb_OverPrBlackLim & "|" & IIf(bCMYK, "1", "0") & "|" & IIf(bCMYKm, "1", "0") & "|" & IIf(bBW, "1", "0") & "|" & IIf(bG, "1", "0") & "|" & IIf(bRGB, "1", "0") & "|" & IIf(bL, "1", "0") & "|" & IIf(bP, "1", "0") & "|" & _
  IIf(b16col, "1", "0") & "|" & IIf(bD, "1", "0") & "|" & IIf(bSpot, "1", "0") & "|" & IIf(bCrop, "1", "0") & "|" & IIf(bAngle0, "1", "0") & "|" & IIf(bOverprint, "1", "0") & "|" & IIf(bLinkBr, "1", "0") & "|" & IIf(FtFillToBit, "1", "0") & "|" & IIf(sMiterLimit, "1", "0") & "|" & _
  sMiterLimitValue & "|" & IIf(scUserInk, "1", "0") & "|" & IIf(sChangeUserColor, "1", "0") & "|" & IIf(cmVisible, "1", "0") & "|" & IIf(cmPrint, "1", "0") & "|" & IIf(cmEnable, "1", "0") & "|" & IIf(cmNVisible, "1", "0") & "|" & cb_OutlineWEdit & "|" & IIf(c_OutlineWEdit, "1", "0")
  '& "|" & IIf(cmNPrint, "1", "0") & "|" & IIf(cmNEnable, "1", "0")
End Function


'====================================================================================
'==================================     Del Preset     ==============================
'====================================================================================
Private Sub cm_DelPres_Click()
  If List2.text = "" Then Exit Sub
  Dim msg&
  msg = MsgBox("Are you sure you want to delete a preset?   ", vbQuestion + vbOKCancel, macroName & " " & macroVersion)
  If msg <> 1 Then Exit Sub
  
  Dim i&: i = List2.ListIndex + 1
  
  Dim c&, i2&
  c = List2.ListCount
  
  If i < c Then
    For i2 = i + 1 To c Step 1
      SaveSetting macroName, "Convert", "Presets" & i, GetSetting(macroName, "Convert", "Presets" & i2)
      SaveSetting macroName, "Convert", "Presets" & i & "Name", GetSetting(macroName, "Convert", "Presets" & i2 & "Name")
      i = i + 1
    Next i2
    DeleteSetting macroName, "Convert", "Presets" & c
    DeleteSetting macroName, "Convert", "Presets" & c & "Name"
  Else
    DeleteSetting macroName, "Convert", "Presets" & i
    DeleteSetting macroName, "Convert", "Presets" & i & "Name"
  End If
  
  SaveSetting macroName, "Convert", "PresetsCount", c - 1
  
  List2.Clear
  myLoadPresets
End Sub
Private Sub cm_DelPres_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  cm_DelPres.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub cm_DelPres_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  cm_DelPres.SpecialEffect = fmSpecialEffectRaised
End Sub
        










'====================================================================================
'===============================     DblClick Preset     ============================
'====================================================================================
Private Sub List2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  If List2.text = "" Then Exit Sub
  
  Dim i&: i = List2.ListIndex + 1
  Dim a$(): a = Split(GetSetting(macroName, "Convert", "Presets" & i), "|")

  cb_ColorType = a(1)
  tToCur = a(2)
  PowerToBit = a(3)
  sSymbToShape = a(4)
  sLens2bit = a(5)
  sBlendB = a(6)
  sContBr = a(7)
  sBevelBit = a(8)
  sContArtMedia = a(9)
  sEnvelToCur = a(10)
  'sExtrudBr = a(11)
  sDistirtToCur = a(12)
  'sPerspToCur = a(13)
  sDropShaToBit = a(14)
  sDropShaMulty = a(15)
  sCMYK = a(16)
  sCMY = a(17)
  sRGB = a(18)
  sBW = a(19)
  sGr = a(20)
  sLAB = a(21)
  sSpot = a(22)
  sPanH = a(23)
  sReg = a(24)
  sHSB = a(25)
  sHLS = a(26)
  sYIQ = a(27)
  sOutScaleNo = a(28)
  sOverprint = a(29)
  sOverprintBlack = a(30)
  meshToBit = a(31)
  PatternToBit = a(32)
  TexturToBit = a(33)
  PSFillToBit = a(34)
  cb_ChaResol = a(35)
  myDPIbox = a(36)
  CG_dpi = a(37)
  M_dpi = a(38)
  bTransp = a(39)
  bAAliasing = a(40)
  bProfile = a(41)
  bOverPrBlack = a(42)
  tb_OverPrBlackLim = a(43)
  bCMYK = a(44)
  bCMYKm = a(45)
  bBW = a(46)
  bG = a(47)
  bRGB = a(48)
  bL = a(49)
  bP = a(50)
  b16col = a(51)
  bD = a(52)
  bSpot = a(53)
  bCrop = a(54)
  bAngle0 = a(55)
  bOverprint = a(56)
  bLinkBr = a(57)

  If CLng(a(0)) >= 2 Then
      FtFillToBit = a(58)
  Else
      FtFillToBit = 0
  End If
  If CLng(a(0)) >= 3 Then
      sMiterLimit = a(59)
      sMiterLimitValue = a(60)
      scUserInk = a(61)
      sChangeUserColor = a(62)
      SaveSetting macroName, "Convert", "cmVisible", a(63)
      SaveSetting macroName, "Convert", "cmPrint", a(64)
      SaveSetting macroName, "Convert", "cmEnable", a(65)
      SaveSetting macroName, "Convert", "cmNVisible", a(66)
      SaveSetting macroName, "Convert", "cb_OutlineWEdit", a(67)
      SaveSetting macroName, "Convert", "c_OutlineWEdit", a(68)
  Else
      sMiterLimit = 0
      sMiterLimitValue = 45
      scUserInk = 0
      sChangeUserColor = 0
      SaveSetting macroName, "Convert", "cmVisible", "1"
      SaveSetting macroName, "Convert", "cmPrint", "1"
      SaveSetting macroName, "Convert", "cmEnable", "1"
      SaveSetting macroName, "Convert", "cmNVisible", "0"
      SaveSetting macroName, "Convert", "cb_OutlineWEdit", "Enlarge to..."
      SaveSetting macroName, "Convert", "c_OutlineWEdit", "0"
  End If
End Sub
        
        
        
        
'====================================================================================
'=================================     Load Preset     ==============================
'====================================================================================
Private Sub myLoadPresets()
  Dim c&, i&, presName$
  c = GetSetting(macroName, "Convert", "PresetsCount", 0)
  For i = 1 To c
    presName = GetSetting(macroName, "Convert", "Presets" & i & "Name")
    If presName <> "" Then List2.AddItem presName
  Next i
End Sub
        

        
        
        


'=======================================================================
'=======================================================================


Private Sub cb_ColorType_Change()
        SaveSetting macroName, "Convert", "cb_ColorType", cb_ColorType.text
        End Sub







Private Sub sLens2bit_Click()
        SaveSetting macroName, "Convert", "sLens2bit", IIf(sLens2bit, "1", "0")
        End Sub
Private Sub sBevelBit_Click()
        SaveSetting macroName, "Convert", "sBevelBit", IIf(sBevelBit, "1", "0")
        End Sub
Private Sub sContArtMedia_Click()
        SaveSetting macroName, "Convert", "sContArtMedia", IIf(sContArtMedia, "1", "0")
        End Sub
Private Sub sDropShaToBit_Click()
        SaveSetting macroName, "Convert", "sDropShaToBit", IIf(sDropShaToBit, "1", "0")
        End Sub
Private Sub sDropShaMulty_Click()
        SaveSetting macroName, "Convert", "sDropShaMulty", IIf(sDropShaMulty, "1", "0")
        End Sub
Private Sub sBlendB_Click()
        SaveSetting macroName, "Convert", "sBlendB", IIf(sBlendB, "1", "0")
        End Sub
Private Sub sContBr_Click()
        SaveSetting macroName, "Convert", "sContBr", IIf(sContBr, "1", "0")
        End Sub
Private Sub sDistirtToCur_Click()
        SaveSetting macroName, "Convert", "sDistirtToCur", IIf(sDistirtToCur, "1", "0")
        End Sub
Private Sub sEnvelToCur_Click()
        SaveSetting macroName, "Convert", "sEnvelToCur", IIf(sEnvelToCur, "1", "0")
        End Sub


Private Sub PowerToBit_Click()
        SaveSetting macroName, "Convert", "PowerToBit", IIf(PowerToBit, "1", "0")
        End Sub
Private Sub sSymbToShape_Click()
        SaveSetting macroName, "Convert", "sSymbToShape", IIf(sSymbToShape, "1", "0")
        End Sub
Private Sub sDimSep_Click()
        SaveSetting macroName, "Convert", "sDimSep", IIf(sDimSep, "1", "0")
        End Sub

        
Private Sub FtFillToBit_Click()
        SaveSetting macroName, "Convert", "FtFillToBit", IIf(FtFillToBit, "1", "0")
        End Sub
Private Sub meshToBit_Click()
        SaveSetting macroName, "Convert", "meshToBit", IIf(meshToBit, "1", "0")
        End Sub



'=======================================================================
'=======================================================================

Private Sub sOutScaleNo_Click()
        SaveSetting macroName, "Convert", "sOutScaleNo", IIf(sOutScaleNo, "1", "0")
        End Sub
        
Private Sub sMiterLimit_Click()
        SaveSetting macroName, "Convert", "sMiterLimit", IIf(sMiterLimit, "1", "0")
        End Sub
Private Sub sMiterLimitValue_Change()
        Dim ml$
        ml = sMiterLimitValue: ml = Replace(ml, ",", ".")
        sMiterLimitValue.text = ml
        SaveSetting macroName, "Convert", "sMiterLimitValue", val(ml)
        End Sub
Private Sub c_OutlineWEdit_Click()
        SaveSetting macroName, "Convert", "c_OutlineWEdit", IIf(c_OutlineWEdit, "1", "0")
        End Sub
Private Sub cb_OutlineWEdit_Change()
        SaveSetting macroName, "Convert", "cb_OutlineWEdit", cb_OutlineWEdit.text
        End Sub

Private Sub sOverprint_Click()
        SaveSetting macroName, "Convert", "sOverprint", IIf(sOverprint, "1", "0")
        If sOverprint = 0 Then
        sOverprintBlack.Enabled = False
        Else
        sOverprintBlack.Enabled = True
        End If
        End Sub

Private Sub sOverprintBlack_Click()
        SaveSetting macroName, "Convert", "sOverprintBlack", IIf(sOverprintBlack, "1", "0")
        End Sub



'=======================================================================
'=======================================================================



Private Sub sBW_Click()
        SaveSetting macroName, "Convert", "sBW", IIf(sBW, "1", "0")
        End Sub
Private Sub sCMYK_Click()
        SaveSetting macroName, "Convert", "sCMYK", IIf(sCMYK, "1", "0")
        End Sub
Private Sub sCMY_Click()
        SaveSetting macroName, "Convert", "sCMY", IIf(sCMY, "1", "0")
        End Sub
Private Sub sGr_Click()
        SaveSetting macroName, "Convert", "sGr", IIf(sGr, "1", "0")
        End Sub
Private Sub sHLS_Click()
        SaveSetting macroName, "Convert", "sHLS", IIf(sHLS, "1", "0")
        End Sub
Private Sub sHSB_Click()
        SaveSetting macroName, "Convert", "sHSB", IIf(sHSB, "1", "0")
        End Sub
Private Sub sLAB_Click()
        SaveSetting macroName, "Convert", "sLAB", IIf(sLAB, "1", "0")
        End Sub
Private Sub sPanH_Click()
        SaveSetting macroName, "Convert", "sPanH", IIf(sPanH, "1", "0")
        End Sub
Private Sub sReg_Click()
        SaveSetting macroName, "Convert", "sReg", IIf(sReg, "1", "0")
        End Sub
Private Sub sRGB_Click()
        SaveSetting macroName, "Convert", "sRGB", IIf(sRGB, "1", "0")
        End Sub
Private Sub sChangeUserColor_Click()
        SaveSetting macroName, "Convert", "sChangeUserColor", IIf(sChangeUserColor, "1", "0")
        End Sub
Private Sub sSpot_Click()
        SaveSetting macroName, "Convert", "sSpot", IIf(sSpot, "1", "0")
        End Sub
Private Sub sYIQ_Click()
        SaveSetting macroName, "Convert", "sYIQ", IIf(sYIQ, "1", "0")
        End Sub
Private Sub scUserInk_Click()
        SaveSetting macroName, "Convert", "sUserInk", IIf(scUserInk, "1", "0")
        End Sub


'=======================================================================
'=======================================================================

Private Sub PatternToBit_Click()
        SaveSetting macroName, "Convert", "PatternToBit", IIf(PatternToBit, "1", "0")
        End Sub

Private Sub TexturToBit_Click()
        SaveSetting macroName, "Convert", "TexturToBit", IIf(TexturToBit, "1", "0")
        End Sub
Private Sub PSFillToBit_Click()
        SaveSetting macroName, "Convert", "PSFillToBit", IIf(PSFillToBit, "1", "0")
        End Sub




'=======================================================================
'=======================================================================

Private Sub tToCur_Click()
        SaveSetting macroName, "Convert", "tToCur", IIf(tToCur, "1", "0")
        End Sub


'=======================================================================
'=======================================================================

Private Sub bAngle0_Click()
        SaveSetting macroName, "Convert", "bAngle0", IIf(bAngle0, "1", "0")
        End Sub
Private Sub bCrop_Click()
        SaveSetting macroName, "Convert", "bCrop", IIf(bCrop, "1", "0")
        End Sub
Private Sub bOverprint_Click()
        SaveSetting macroName, "Convert", "bOverprint", IIf(bOverprint, "1", "0")
        End Sub
Private Sub bLinkBr_Click()
        SaveSetting macroName, "Convert", "bLinkBr", IIf(bLinkBr, "1", "0")
        End Sub






Private Sub cb_ChaResol_Click()
        SaveSetting macroName, "Convert", "cb_ChaResol", IIf(cb_ChaResol, "1", "0")
        End Sub
Private Sub myDPIbox_Change()
        SaveSetting macroName, "Convert", "myDPIbox", myDPIbox
        End Sub
Private Sub CG_dpi_Change()
        SaveSetting macroName, "Convert", "CG_dpi", CG_dpi
        End Sub
Private Sub M_dpi_Change()
        SaveSetting macroName, "Convert", "M_dpi", M_dpi
        End Sub





Private Sub bCMYK_Click()
        SaveSetting macroName, "Convert", "bCMYK", IIf(bCMYK, "1", "0")
        End Sub
Private Sub bCMYKm_Click()
        SaveSetting macroName, "Convert", "bCMYKm", IIf(bCMYKm, "1", "0")
        End Sub
Private Sub bSpot_Click()
        SaveSetting macroName, "Convert", "bSpot", IIf(bSpot, "1", "0")
        End Sub
Private Sub b16col_Click()
        SaveSetting macroName, "Convert", "b16col", IIf(b16col, "1", "0")
        End Sub
Private Sub bBW_Click()
        SaveSetting macroName, "Convert", "bBW", IIf(bBW, "1", "0")
        End Sub
Private Sub bD_Click()
        SaveSetting macroName, "Convert", "bD", IIf(bD, "1", "0")
        End Sub
Private Sub bG_Click()
        SaveSetting macroName, "Convert", "bG", IIf(bG, "1", "0")
        End Sub
Private Sub bL_Click()
        SaveSetting macroName, "Convert", "bL", IIf(bL, "1", "0")
        End Sub
Private Sub bP_Click()
        SaveSetting macroName, "Convert", "bP", IIf(bP, "1", "0")
        End Sub
Private Sub bRGB_Click()
        SaveSetting macroName, "Convert", "bRGB", IIf(bRGB, "1", "0")
        End Sub







Private Sub bAAliasing_Click()
        SaveSetting macroName, "Convert", "bAAliasing", IIf(bAAliasing, "1", "0")
        End Sub
Private Sub bProfile_Click()
        SaveSetting macroName, "Convert", "bProfile", IIf(bProfile, "1", "0")
        End Sub
Private Sub bTransp_Click()
        SaveSetting macroName, "Convert", "bTransp", IIf(bTransp, "1", "0")
        End Sub
Private Sub bOverPrBlack_Click()
        SaveSetting macroName, "Convert", "bOverPrBlack", IIf(bOverPrBlack, "1", "0")
        End Sub
Private Sub tb_OverPrBlackLim_Change()
        SaveSetting macroName, "Convert", "tb_OverPrBlackLim", tb_OverPrBlackLim
        End Sub










Private Sub UserForm_Initialize()
        Dim s$
        
        Me.Height = 225: Me.Width = 407 ' 283
        myLoadPresets
        
        s = GetSetting(macroName, "Convert", "Pos")
        If Len(s) Then
            StartUpPosition = 0
            Me.Top = CSng(Split(s, " ")(1))
            Me.Left = CSng(Split(s, " ")(0))
        End If
        
        cb_ChaResol = (GetSetting(macroName, "Convert", "cb_ChaResol", "1") = "1")
        myDPIbox.AddItem "Auto (min/max)"
        myDPIbox.AddItem "User"
        myDPIbox.text = GetSetting(macroName, "Convert", "myDPIbox", "Auto (min/max)")
        
        cb_ColorType.List = Array("CMYK Color", "Grey Color", "RGB Color")
        cb_ColorType = GetSetting(macroName, "Convert", "cb_ColorType", "CMYK Color")
        
        c_OutlineWEdit = (GetSetting(macroName, "Convert", "c_OutlineWEdit", "0") = "1")
        cb_OutlineWEdit.List = Array("Enlarge to...", "Remove")
        cb_OutlineWEdit = GetSetting(macroName, "Convert", "cb_OutlineWEdit", "Enlarge to...")
        
        sLens2bit = (GetSetting(macroName, "Convert", "sLens2bit", "0") = "1")
        sBevelBit = (GetSetting(macroName, "Convert", "sBevelBit", "1") = "1")
        sContArtMedia = (GetSetting(macroName, "Convert", "sContArtMedia", "1") = "1")
        sDropShaToBit = (GetSetting(macroName, "Convert", "sDropShaToBit", "1") = "1")
        sDropShaMulty = (GetSetting(macroName, "Convert", "sDropShaMulty", "0") = "1")
        sBlendB = (GetSetting(macroName, "Convert", "sBlendB", "1") = "1")
        sContBr = (GetSetting(macroName, "Convert", "sContBr", "1") = "1")
        sDistirtToCur = (GetSetting(macroName, "Convert", "sDistirtToCur", "1") = "1")
        sEnvelToCur = (GetSetting(macroName, "Convert", "sEnvelToCur", "1") = "1")
        
        PowerToBit = (GetSetting(macroName, "Convert", "PowerToBit", "0") = "1")
        sSymbToShape = (GetSetting(macroName, "Convert", "sSymbToShape", "1") = "1")
        sDimSep = (GetSetting(macroName, "Convert", "sDimSep", "1") = "1")
        
        meshToBit = (GetSetting(macroName, "Convert", "meshToBit", "1") = "1")
        FtFillToBit = (GetSetting(macroName, "Convert", "FtFillToBit", "0") = "1")
        
        sOutScaleNo = (GetSetting(macroName, "Convert", "sOutScaleNo", "1") = "1")
        sMiterLimit = (GetSetting(macroName, "Convert", "sMiterLimit", "0") = "1")
        sMiterLimitValue = GetSetting(macroName, "Convert", "sMiterLimitValue", "45")
        
        sOverprint = (GetSetting(macroName, "Convert", "sOverprint", "1") = "1")
        sOverprintBlack = (GetSetting(macroName, "Convert", "sOverprintBlack", "0") = "1")
        
        sCMYK = (GetSetting(macroName, "Convert", "sCMYK", "0") = "1")
        sCMY = (GetSetting(macroName, "Convert", "sCMY", "0") = "1")
        sBW = (GetSetting(macroName, "Convert", "sBW", "0") = "1")
        sGr = (GetSetting(macroName, "Convert", "sGr", "0") = "1")
        sHLS = (GetSetting(macroName, "Convert", "sHLS", "1") = "1")
        sHSB = (GetSetting(macroName, "Convert", "sHSB", "1") = "1")
        sLAB = (GetSetting(macroName, "Convert", "sLAB", "1") = "1")
        sPanH = (GetSetting(macroName, "Convert", "sPanH", "1") = "1")
        sReg = (GetSetting(macroName, "Convert", "sReg", "1") = "1")
        sRGB = (GetSetting(macroName, "Convert", "sRGB", "1") = "1")
        sSpot = (GetSetting(macroName, "Convert", "sSpot", "1") = "1")
        sYIQ = (GetSetting(macroName, "Convert", "sYIQ", "1") = "1")
        scUserInk = (GetSetting(macroName, "Convert", "sUserInk", "1") = "1")
        sChangeUserColor = (GetSetting(macroName, "Convert", "sChangeUserColor", "0") = "1")
        
        PatternToBit = (GetSetting(macroName, "Convert", "PatternToBit", "1") = "1")
        TexturToBit = (GetSetting(macroName, "Convert", "TexturToBit", "1") = "1")
        PSFillToBit = (GetSetting(macroName, "Convert", "PSFillToBit", "1") = "1")
        
        tToCur = (GetSetting(macroName, "Convert", "tToCur", "1") = "1")
        
        bCrop = (GetSetting(macroName, "Convert", "bCrop", "1") = "1")
        bAngle0 = (GetSetting(macroName, "Convert", "bAngle0", "1") = "1")
        bOverprint = (GetSetting(macroName, "Convert", "bOverprint", "1") = "1")
        bLinkBr = (GetSetting(macroName, "Convert", "bLinkBr", "0") = "1")
        
        M_dpi = GetSetting(macroName, "Convert", "M_dpi", "600")
        CG_dpi = GetSetting(macroName, "Convert", "CG_dpi", "300")
        
        bCMYK = (GetSetting(macroName, "Convert", "bCMYK", "0") = "1")
        bCMYKm = (GetSetting(macroName, "Convert", "bCMYKm", "0") = "1")
        bSpot = (GetSetting(macroName, "Convert", "bSpot", "0") = "1")
        b16col = (GetSetting(macroName, "Convert", "b16col", "1") = "1")
        bBW = (GetSetting(macroName, "Convert", "bBW", "0") = "1")
        bD = (GetSetting(macroName, "Convert", "bD", "1") = "1")
        bG = (GetSetting(macroName, "Convert", "bG", "0") = "1")
        bL = (GetSetting(macroName, "Convert", "bL", "1") = "1")
        bP = (GetSetting(macroName, "Convert", "bP", "1") = "1")
        bRGB = (GetSetting(macroName, "Convert", "bRGB", "1") = "1")
        
        bAAliasing = (GetSetting(macroName, "Convert", "bAAliasing", "1") = "1")
        bProfile = (GetSetting(macroName, "Convert", "bProfile", "1") = "1")
        bOverPrBlack = (GetSetting(macroName, "Convert", "bOverPrBlack", "0") = "1")
        tb_OverPrBlackLim = GetSetting(macroName, "Convert", "tb_OverPrBlackLim", "95")
        bTransp = (GetSetting(macroName, "Convert", "bTransp", "1") = "1")
        
        
        If sOverprint = 0 Then sOverprintBlack.Enabled = False
        cmPresVerCaption.Caption = "Preset version: " & convPresVer
        End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectEtched
        cm_addPres.SpecialEffect = fmSpecialEffectEtched
        cm_DelPres.SpecialEffect = fmSpecialEffectEtched
        cm_presSave.SpecialEffect = fmSpecialEffectEtched
        End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        SaveSetting macroName, "Convert", "Pos", Left & " " & Top
        End Sub
