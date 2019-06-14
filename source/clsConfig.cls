VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clrFind As New Color
Public clrReplace As New Color
Public choosing As Boolean
Private Const sDefColor$ = "CMYK,USER,0,0,0,100,0,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0" '138A0000000000000064000000000000



Private Sub Class_Initialize()
  clrFind.CopyAssign CreateColor(sDefColor)
  clrReplace.CopyAssign CreateColor(sDefColor)
  choosing = False
End Sub


Public Sub Load()
  SetFindStr GetSetting(macroName, "ColorReplacer", "Find", sDefColor)
  SetReplaceStr GetSetting(macroName, "ColorReplacer", "Replace", sDefColor)
End Sub
    

Public Sub Save()
  SaveSetting macroName, "ColorReplacer", "Find", clrFind.ToString
  SaveSetting macroName, "ColorReplacer", "Replace", clrReplace.ToString
End Sub
    
Public Sub SaveToList()
  Dim c&: c = GetSetting(macroName, "ColorReplacer", "ColorListCount", 0) + 1
  SaveSetting macroName, "ColorReplacer", "ColorList" & c, clrFind.ToString & "|" & clrReplace.ToString
  SaveSetting macroName, "ColorReplacer", "ColorListCount", c
End Sub


Public Sub SetFindStr(val$)
  clrFind.CopyAssign CreateColor(val)
End Sub

Public Sub SetReplaceStr(val$)
  clrReplace.CopyAssign CreateColor(val)
End Sub
