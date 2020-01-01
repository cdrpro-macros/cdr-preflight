VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionInfo_Form 
   Caption         =   "Options"
   ClientHeight    =   7425
   ClientLeft      =   42
   ClientTop       =   406
   ClientWidth     =   6090
   OleObjectBlob   =   "OptionInfo_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionInfo_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Idx As Collection







Private Sub myOutlineWidthMin_Change()
        Dim s$
        s = myOutlineWidthMin.text
        s = Replace(s, ",", ".")
        myOutlineWidthMin.value = s
        End Sub


Private Sub Command_OK_Click()
        SaveSetting macroName, sREGAPPOPT, "OperatingMode", myOperatingMode
        SaveSetting macroName, sREGAPPOPT, "cb_Unit", cb_Unit
        SaveSetting macroName, sREGAPPOPT, "LimitInList", cm_limitInList
        SaveSetting macroName, sREGAPPOPT, "TransparencyForm", IIf(oi_Transparency, "1", "0")
        
        SaveSetting macroName, sREGAPPOPT, "TILFill", myTILFill
        'SaveSetting macroName, sREGAPPOPT, "TILOutline", myTILOutline
        SaveSetting macroName, sREGAPPOPT, "myMinColor", myMinColor
        SaveSetting macroName, sREGAPPOPT, "cb_showColor", IIf(cb_showColor, "1", "0")
        SaveSetting macroName, sREGAPPOPT, "myDpiMin", myDpiMin
        SaveSetting macroName, sREGAPPOPT, "myDpiMax", myDpiMax
        
        SaveSetting macroName, sREGAPPOPT, "OutlineWidthMin", myOutlineWidthMin
        SaveSetting macroName, sREGAPPOPT, "NodesCount", tb_NodesCount
        SaveSetting macroName, sREGAPPOPT, "SmalFontPt", tb_SmalFontPt
        SaveSetting macroName, sREGAPPOPT, "SmalFontColor", tb_SmalFontColor
        
        'SaveSetting macroName, sREGAPPOPT, "ErrLogSave", IIf(cb_ErrLogSave, "1", "0")
        Unload Me
        Select Case GetSetting(macroName, sREGAPPOPT, "cb_Unit", "millimeters")
        Case "millimeters": myUnitWork = cdrMillimeter
        Case "points": myUnitWork = cdrPoint
        Case Else: MsgBox "No Unit   ", vbCritical, "Warning"
        End Select
'        CdrPreflight_start
        End Sub
Private Sub Command_OK_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectSunken
        End Sub
Private Sub Command_OK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectRaised
        End Sub














Private Sub UserForm_Initialize()
        myOperatingMode.AddItem "Default"
        myOperatingMode.AddItem "In select"
        myOperatingMode.AddItem "On Active page"
        myOperatingMode.AddItem "On Master pages"
        myOperatingMode.AddItem "All page"
        myOperatingMode = GetSetting(macroName, sREGAPPOPT, "OperatingMode", "Default")
        cb_Unit.List = Array("millimeters", "points")
        cb_Unit = GetSetting(macroName, sREGAPPOPT, "cb_Unit", "millimeters")
        cm_limitInList = GetSetting(macroName, sREGAPPOPT, "LimitInList", "0")
        oi_Transparency = GetSetting(macroName, sREGAPPOPT, "TransparencyForm", "1")
        
        myTILFill = GetSetting(macroName, sREGAPPOPT, "TILFill", "280")
        'myTILOutline = GetSetting(macroName, sREGAPPOPT, "TILOutline", "280")
        myMinColor = GetSetting(macroName, sREGAPPOPT, "myMinColor", "10")
        cb_showColor = GetSetting(macroName, sREGAPPOPT, "cb_showColor", "0")
        myDpiMin = GetSetting(macroName, sREGAPPOPT, "myDpiMin", "250")
        myDpiMax = GetSetting(macroName, sREGAPPOPT, "myDpiMax", "320")
        
        myOutlineWidthMin = GetSetting(macroName, sREGAPPOPT, "OutlineWidthMin", "0.0762")
        tb_NodesCount = GetSetting(macroName, sREGAPPOPT, "NodesCount", "8000")
        tb_SmalFontPt = GetSetting(macroName, sREGAPPOPT, "SmalFontPt", "6")
        tb_SmalFontColor = GetSetting(macroName, sREGAPPOPT, "SmalFontColor", "12")
        
        'cb_ErrLogSave = GetSetting(macroName, sREGAPPOPT, "ErrLogSave", "0")
        
        myLoadPresetsList
        myLoadPresets
        End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Command_OK.SpecialEffect = fmSpecialEffectEtched
        End Sub







Private Sub cm_CheckAll_Click(): cm_Check True: End Sub
Private Sub cm_UncheckAll_Click(): cm_Check False: End Sub
Private Sub cm_Check(b As Boolean)
  Dim i&
  For i = 0 To List2.ListCount - 1
    If List2.selected(i) <> b Then List2.selected(i) = b
  Next
End Sub







'====================================================================================
'=================================      Load List      ==============================
'====================================================================================
Private Sub myLoadPresetsList()
  Dim c&, a$()
  If myResumeErr Then On Error Resume Next
  a = Split(cdrMyType, "|")
  
  Set Idx = New Collection
  
  For c = 1 To UBound(a) + 1
    Select Case a(c - 1)
      Case "Problem outlines", "li1", "li2", "li3", "li4", "li5", "li6"
        'Do Nothing
        Idx.Add IGNORE_ITEM, "i" & c
      Case Else
        List2.AddItem a(c - 1)
        Idx.Add CStr(List2.ListCount - 1), "i" & c
    End Select
  Next
  End Sub


'====================================================================================
'=================================     Load Preset     ==============================
'====================================================================================
Private Sub myLoadPresets()
  Dim c&, i&, presName$
  If myResumeErr Then On Error Resume Next
  c = GetSetting(macroName, sREGAPPOPT, "PresetsCount", 0)
  For i = 1 To c
    presName = GetSetting(macroName, sREGAPPOPT, "Presets" & i & "Name")
    If presName <> "" Then cb_presList.AddItem i & "| " & presName
  Next i
End Sub
        
Private Sub cb_presList_Change()
  If myResumeErr Then On Error Resume Next
  If cb_presList.SelLength = 0 Then Exit Sub
  
  Dim a$()
  a = Split(cb_presList.SelText, "|")
  a = Split(GetSetting(macroName, sREGAPPOPT, "Presets" & a(0)), "|")
  
  Dim i&, p$()
  For i = 1 To UBound(a) - 1
    p = Split(a(i), "-")
    If Idx.Item("i" & i) <> IGNORE_ITEM Then
      List2.selected(CLng(Idx.Item("i" & i))) = p(0)
    End If
  Next i
End Sub




'====================================================================================
'=================================     Save Preset     ==============================
'====================================================================================
Private Sub cm_presSave_Click()
  If cb_presList.SelText = "" Then
    MsgBox "Need to select any preset before.", vbCritical, macroName & " " & macroVersion
    Exit Sub
  End If
  Dim msg&
  msg = MsgBox("Are you sure you want to save a preset?   ", vbQuestion + vbOKCancel, macroName & " " & macroVersion)
  If msg <> 1 Then Exit Sub
  
  Dim c1&, i&, a$()
  For c1 = 0 To cb_presList.ListCount - 1 Step 1
      If cb_presList.SelText = cb_presList.List(c1) Then
          a = Split(cb_presList.SelText, "|")
          i = CLng(a(0))
          Exit For
      End If
  Next c1
  
  Dim strPres$
  strPres = ctcPresVer & "|"
  
  Dim i2&
  For i2 = 1 To Idx.Count
    If Idx(i2) = IGNORE_ITEM Then
      strPres = strPres & "0|"
    Else
      If List2.selected(Idx(i2)) Then strPres = strPres & "1|" Else strPres = strPres & "0|"
    End If
  Next
  
  SaveSetting macroName, sREGAPPOPT, "Presets" & i, strPres
End Sub
        
'====================================================================================
'==================================     Add Preset     ==============================
'====================================================================================
Private Sub cm_presAdd_Click()
  If myResumeErr Then On Error Resume Next
  Dim c&: c = CLng(GetSetting(macroName, sREGAPPOPT, "PresetsCount", 0)) + 1
  
  Dim strPresN$: strPresN = InputBox("Name for Preset", "Name...")
  If strPresN = "" Then Exit Sub
  
  Dim strPres$
  strPres = ctcPresVer & "|"
  
  Dim i&
  For i = 1 To Idx.Count
    If Idx(i) = IGNORE_ITEM Then
      strPres = strPres & "0|"
    Else
      If List2.selected(Idx(i)) Then strPres = strPres & "1|" Else strPres = strPres & "0|"
    End If
  Next
  
  SaveSetting macroName, sREGAPPOPT, "Presets" & c, strPres
  SaveSetting macroName, sREGAPPOPT, "Presets" & c & "Name", strPresN
  
  cb_presList.AddItem c & "| " & strPresN
  cb_presList.text = c & "| " & strPresN
  SaveSetting macroName, sREGAPPOPT, "PresetsCount", c
End Sub

'====================================================================================
'==================================     Del Preset     ==============================
'====================================================================================
Private Sub cm_presDel_Click()
  If cb_presList.SelText = "" Then
    MsgBox "Need to select any preset before.", vbCritical, macroName & " " & macroVersion
    Exit Sub
  End If
  
  Dim msg&
  msg = MsgBox("Are you sure you want to delete a preset?   ", vbQuestion + vbOKCancel, macroName & " " & macroVersion)
  If msg <> 1 Then Exit Sub
  
  Dim i&, c&, c1&, i2&, a$()
  If myResumeErr Then On Error Resume Next

  If cb_presList.SelLength = 0 Then Exit Sub
  For c1 = 0 To cb_presList.ListCount - 1 Step 1
      If cb_presList.SelText = cb_presList.List(c1) Then
          a = Split(cb_presList.SelText, "|")
          i = CLng(a(0))
          Exit For
      End If
  Next c1

  cb_presList.Clear
  c = CLng(GetSetting(macroName, sREGAPPOPT, "PresetsCount", 0))
  If i < c Then
      For i2 = i + 1 To c Step 1
          SaveSetting macroName, sREGAPPOPT, "Presets" & i, _
          GetSetting(macroName, sREGAPPOPT, "Presets" & i2)
          SaveSetting macroName, sREGAPPOPT, "Presets" & i & "Name", _
          GetSetting(macroName, sREGAPPOPT, "Presets" & i2 & "Name")
          i = i + 1
      Next i2
      DeleteSetting macroName, sREGAPPOPT, "Presets" & c
      DeleteSetting macroName, sREGAPPOPT, "Presets" & c & "Name"
  Else
      DeleteSetting macroName, sREGAPPOPT, "Presets" & i
      DeleteSetting macroName, sREGAPPOPT, "Presets" & i & "Name"
  End If
  
  SaveSetting macroName, sREGAPPOPT, "PresetsCount", c - 1
  myLoadPresets
End Sub
        
        
'====================================================================================

Private Sub sl_SH_Click()
  DoItemsCheck 3, 11
End Sub
Private Sub sl_EF_Click()
  DoItemsCheck 12, 23
End Sub
Private Sub sl_TX_Click()
  DoItemsCheck 24, 29
End Sub
Private Sub sl_FO_Click()
  DoItemsCheck 30, 43
End Sub
Private Sub sl_CL_Click()
  DoItemsCheck 44, 61
End Sub
Private Sub sl_BT_Click()
  DoItemsCheck 62, 80
End Sub

Private Sub DoItemsCheck(s&, e&)
  Dim i&
  For i = s To e
    If List2.selected(i) Then List2.selected(i) = False Else List2.selected(i) = True
  Next
End Sub
