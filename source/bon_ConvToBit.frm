VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bon_ConvToBit 
   Caption         =   "Convert to Bitmap"
   ClientHeight    =   4110
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   2751
   OleObjectBlob   =   "bon_ConvToBit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "bon_ConvToBit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub grWork_CheckBox_Click()
        SaveSetting macroName, "conv2bit", "grWork", IIf(grWork_CheckBox, "1", "0")
        End Sub



Private Sub cb_presets_Change()
        Select Case cb_presets.text
        Case "Preset 1": loadPresetsAtStart 1
        Case "Preset 2": loadPresetsAtStart 2
        Case "Preset 3": loadPresetsAtStart 3
        End Select
        End Sub
        
        
Private Sub cb_colorMode_Change()
        Select Case cb_colorMode.text
        Case "Grayscale": ch_Overptint.Enabled = False: ch_Dithered.Enabled = True
        Case "CMYKColor": ch_Dithered.Enabled = False: ch_Overptint.Enabled = True
        Case "RGBColor": ch_Dithered.Enabled = False: ch_Overptint.Enabled = False
        End Select
        End Sub
Private Sub ch_Overptint_Click()
        If ch_Overptint.value = True Then _
        tx_OverprintLimit.Enabled = True Else tx_OverprintLimit.Enabled = False
        End Sub


Private Sub cm_save_Click()
        Dim s$
        Select Case cb_presets.text
        Case "Preset 1"
            s = cb_dpi.value & "|" & cb_colorMode.text & "|" & ch_Dithered.value & "|" & _
            ch_Profile.value & "|" & ch_Overptint.value & "|" & tx_OverprintLimit.value & "|" & _
            ch_Aliasing.value & "|" & ch_Transparent.value
            SaveSetting macroName, "conv2bit", "Preset1", s
        Case "Preset 2"
            s = cb_dpi.value & "|" & cb_colorMode.text & "|" & ch_Dithered.value & "|" & _
            ch_Profile.value & "|" & ch_Overptint.value & "|" & tx_OverprintLimit.value & "|" & _
            ch_Aliasing.value & "|" & ch_Transparent.value
            SaveSetting macroName, "conv2bit", "Preset2", s
        Case "Preset 3"
            s = cb_dpi.value & "|" & cb_colorMode.text & "|" & ch_Dithered.value & "|" & _
            ch_Profile.value & "|" & ch_Overptint.value & "|" & tx_OverprintLimit.value & "|" & _
            ch_Aliasing.value & "|" & ch_Transparent.value
            SaveSetting macroName, "conv2bit", "Preset3", s
        End Select
        cb_presets_Change
        End Sub
Private Sub cm_save_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_save.SpecialEffect = fmSpecialEffectSunken
        End Sub
Private Sub cm_save_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_save.SpecialEffect = fmSpecialEffectRaised
        End Sub
        
        
        
Private Sub cm_exit_Click()
        Unload Me
        End Sub
Private Sub cm_exit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_exit.SpecialEffect = fmSpecialEffectSunken
        End Sub
Private Sub cm_exit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_exit.SpecialEffect = fmSpecialEffectRaised
        End Sub
        
        



Private Sub UserForm_Initialize()
        grWork_CheckBox = (GetSetting(macroName, "conv2bit", "grWork", "1") = "1")
        cb_presets.List = Array("Preset 1", "Preset 2", "Preset 3")
        cb_presets.text = GetSetting(macroName, "conv2bit", "Presets", "Preset 1")
        
        cb_dpi.AddItem "300"
        cb_dpi.AddItem "200"
        cb_dpi.AddItem "150"
        cb_dpi.AddItem "100"
        cb_dpi.AddItem "96"
        cb_dpi.AddItem "72"
        cb_dpi = "300"
        
        cb_colorMode.List = Array("Grayscale", "CMYKColor", "RGBColor")
        cb_colorMode = "CMYKColor"
        
        ch_Profile = True
        tx_OverprintLimit = "95"
        ch_Aliasing = True
        
        loadPresetsAtStart 1
        ch_Overptint_Click
        End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        cm_save.SpecialEffect = fmSpecialEffectEtched
        cm_exit.SpecialEffect = fmSpecialEffectEtched
        End Sub
        

Private Sub loadPresetsAtStart(l&)
        Dim s$, a$()
        
        Select Case l
        Case 1
            s = GetSetting(macroName, "conv2bit", "Preset1")
            If s = "" Then prLoadStat.Caption = "False": Exit Sub
            a = Split(s, "|")
            cb_dpi.value = a(0)
            cb_colorMode.value = a(1)
            ch_Dithered.value = a(2)
            ch_Profile.value = a(3)
            ch_Overptint.value = a(4)
            tx_OverprintLimit.value = a(5)
            ch_Aliasing.value = a(6)
            ch_Transparent.value = a(7)
            prLoadStat.Caption = "True"
        Case 2
            s = GetSetting(macroName, "conv2bit", "Preset2")
            If s = "" Then prLoadStat.Caption = "False": Exit Sub
            a = Split(s, "|")
            cb_dpi.value = a(0)
            cb_colorMode.value = a(1)
            ch_Dithered.value = a(2)
            ch_Profile.value = a(3)
            ch_Overptint.value = a(4)
            tx_OverprintLimit.value = a(5)
            ch_Aliasing.value = a(6)
            ch_Transparent.value = a(7)
            prLoadStat.Caption = "True"
        Case 3
            s = GetSetting(macroName, "conv2bit", "Preset3")
            If s = "" Then prLoadStat.Caption = "False": Exit Sub
            a = Split(s, "|")
            cb_dpi.value = a(0)
            cb_colorMode.value = a(1)
            ch_Dithered.value = a(2)
            ch_Profile.value = a(3)
            ch_Overptint.value = a(4)
            tx_OverprintLimit.value = a(5)
            ch_Aliasing.value = a(6)
            ch_Transparent.value = a(7)
            prLoadStat.Caption = "True"
        End Select
        End Sub


