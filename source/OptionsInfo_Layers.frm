VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionsInfo_Layers 
   Caption         =   "Layers Options"
   ClientHeight    =   2175
   ClientLeft      =   42
   ClientTop       =   434
   ClientWidth     =   2100
   OleObjectBlob   =   "OptionsInfo_Layers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionsInfo_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmVisible_Click()
        'If cmVisible.Value = True Then cmNVisible.Value = False
        SaveSetting macroName, "Convert", "cmVisible", IIf(cmVisible, "1", "0")
        End Sub
Private Sub cmPrint_Click()
        'If cmPrint.Value = True Then cmNPrint.Value = False
        SaveSetting macroName, "Convert", "cmPrint", IIf(cmPrint, "1", "0")
        End Sub
Private Sub cmEnable_Click()
        'If cmPrint.Value = True Then cmNEnable.Value = False
        SaveSetting macroName, "Convert", "cmEnable", IIf(cmEnable, "1", "0")
        End Sub
        
Private Sub cmNVisible_Click()
        'If cmNVisible.Value = True Then cmVisible.Value = False
        SaveSetting macroName, "Convert", "cmNVisible", IIf(cmNVisible, "1", "0")
        End Sub
'Private Sub cmNPrint_Click()
'        If cmNPrint.Value = True Then cmPrint.Value = False
'        SaveSetting macroName, "Convert", "cmNPrint", IIf(cmNPrint, "1", "0")
'        End Sub
'Private Sub cmNEnable_Click()
'        If cmNEnable.Value = True Then cmEnable.Value = False
'        SaveSetting macroName, "Convert", "cmNEnable", IIf(cmNEnable, "1", "0")
'        End Sub


Private Sub Command_OK_Click()
        Unload Me
        End Sub
Private Sub Command_OK_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectSunken
        End Sub
Private Sub Command_OK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectRaised
        End Sub
        
        
Private Sub UserForm_Initialize()
        cmVisible = (GetSetting(macroName, "Convert", "cmVisible", "1") = "1")
        cmPrint = (GetSetting(macroName, "Convert", "cmPrint", "1") = "1")
        cmEnable = (GetSetting(macroName, "Convert", "cmEnable", "1") = "1")
        cmNVisible = (GetSetting(macroName, "Convert", "cmNVisible", "0") = "1")
        'cmNPrint = (GetSetting(macroName, "Convert", "cmNPrint", "0") = "1")
        'cmNEnable = (GetSetting(macroName, "Convert", "cmNEnable", "0") = "1")
        End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectEtched
        End Sub
