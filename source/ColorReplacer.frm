VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorReplacer 
   Caption         =   "Color Replacer"
   ClientHeight    =   8040
   ClientLeft      =   42
   ClientTop       =   434
   ClientWidth     =   5985
   OleObjectBlob   =   "ColorReplacer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cfg As New clsConfig



Private Sub cm_Close_Click()
  If Not cfg.choosing Then Unload Me
End Sub


Private Sub UserForm_Initialize()
  LoadList
  cfg.Load
  UpdateFind
  UpdateReplace
End Sub
    
    
        
Private Sub LoadList()
  Dim i&, a$()
  For i = 1 To GetSetting(macroName, "ColorReplacer", "ColorListCount", 0)
    a = Split(GetSetting(macroName, "ColorReplacer", "ColorList" & i), "|")
    cfg.SetFindStr a(0)
    cfg.SetReplaceStr a(1)
    AddToList cfg.clrFind, cfg.clrReplace
  Next i
End Sub
        
Private Sub cm_add_Click()
  AddToList cfg.clrFind, cfg.clrReplace
  cfg.Save
  cfg.SaveToList
End Sub
Private Sub cm_Del_Click()
  If ListFind.text = "" Then Exit Sub
  
  Dim msg&
  msg = MsgBox("Are you sure you want to delete?   ", vbQuestion + vbOKCancel, macroName & " " & macroVersion)
  If msg <> 1 Then Exit Sub
  
  Dim i&: i = ListFind.ListIndex + 1
  
  Dim c&, i2&
  c = CLng(GetSetting(macroName, "ColorReplacer", "ColorListCount", 0))
  If i < c Then
      For i2 = i + 1 To c Step 1
          SaveSetting macroName, "ColorReplacer", "ColorList" & i, _
          GetSetting(macroName, "ColorReplacer", "ColorList" & i2)
          i = i + 1
      Next i2
      DeleteSetting macroName, "ColorReplacer", "ColorList" & c
  Else
      DeleteSetting macroName, "ColorReplacer", "ColorList" & i
  End If
  SaveSetting macroName, "ColorReplacer", "ColorListCount", c - 1
  
  ListFind.Clear
  ListReplace.Clear
  LoadList
End Sub
        
        
Private Sub AddToList(c1 As Color, c2 As Color)
  ListFind.AddItem c1.Name(True)
  ListReplace.AddItem c2.Name(True)
End Sub




Private Sub ListFind_Click()
  ListSelectItem Me.ListFind, Me.ListReplace
End Sub

Private Sub ListReplace_Click()
  ListSelectItem Me.ListReplace, Me.ListFind
End Sub

Private Sub ListSelectItem(lf As ListBox, ls As ListBox)
  If lf.text = "" Then Exit Sub
  Dim i&: i = lf.ListIndex
  ls.selected(i) = True
  
  Dim a$()
  a = Split(GetSetting(macroName, "ColorReplacer", "ColorList" & (i + 1)), "|")
  cfg.SetFindStr a(0)
  cfg.SetReplaceStr a(1)
  UpdateFind
  UpdateReplace
  
End Sub







Private Sub cmFind_Click()
    cfg.choosing = True: Me.Enabled = False
    If cfg.clrFind.UserAssignEx() Then UpdateFind
    Me.Enabled = True: cfg.choosing = False
    End Sub
Private Sub cmReplace_Click()
    cfg.choosing = True: Me.Enabled = False
    If cfg.clrReplace.UserAssignEx() Then UpdateReplace
    Me.Enabled = True: cfg.choosing = False
    End Sub
    

Private Sub UpdateFind()
    UpdateColor cfg.clrFind, cmFind, lblFind
    End Sub
Private Sub UpdateReplace()
    UpdateColor cfg.clrReplace, cmReplace, lblReplace
    End Sub
Private Sub UpdateColor(c As Color, btn As MSForms.label, lbl As MSForms.label)
    Dim sName As String
    Dim cRGB As New Color
    cRGB.CopyAssign c
    cRGB.ConvertToRGB
    btn.BackColor = RGB(cRGB.RGBRed, cRGB.RGBGreen, cRGB.RGBBlue)
    sName = " " & c.Name & " (" & c.Name(True) & ")"
    lbl.Caption = sName
    End Sub
    
    
Private Sub cmSwap_Click()
    Dim co As New Color
    co.CopyAssign cfg.clrFind
    cfg.clrFind.CopyAssign cfg.clrReplace
    cfg.clrReplace.CopyAssign co
    UpdateFind
    UpdateReplace
    End Sub
