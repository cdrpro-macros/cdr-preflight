VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uShowTIL 
   Caption         =   "Total Ink Limit"
   ClientHeight    =   2400
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   5075
   OleObjectBlob   =   "uShowTIL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uShowTIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private inMode As Boolean, srTarget As ShapeRange
#If VBA7 Then
  'PtrSafe
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
  Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long 'Used to see if PP is open
  Private Declare PtrSafe Function ShowWindow Lib "USER32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#Else
  Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
  Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long 'Used to see if PP is open
  Private Declare Function ShowWindow Lib "USER32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If

Private Const OPTSTR$ = "ShowTIL"
Private Const SW_HIDE& = 0
Private Const SW_MINIMIZE& = 6




Private Sub UserForm_Initialize()
    Me.Height = 44.25
    Me.Width = 124.5
    
    Dim s$: s = GetSetting(macroName, OPTSTR, "Pos")
    If Len(s) Then
      StartUpPosition = 0
      Me.Left = CSng(Split(s, " ")(0))
      Me.Top = CSng(Split(s, " ")(1))
    End If
    
    cbLimit.AddItem "240"
    cbLimit.AddItem "260"
    cbLimit.AddItem "280"
    cbLimit.AddItem "300"
    cbLimit.AddItem "320"
    cbLimit.AddItem "340"
    cbLimit.text = GetSetting(macroName, OPTSTR, "Limit", "280")
    
    cbAccuracy.AddItem "10"
    cbAccuracy.AddItem "20"
    cbAccuracy.AddItem "30"
    cbAccuracy.AddItem "50"
    cbAccuracy.AddItem "80"
    cbAccuracy.AddItem "100"
    cbAccuracy.AddItem "150"
    cbAccuracy.text = GetSetting(macroName, OPTSTR, "Accuracy", "20")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveSetting macroName, OPTSTR, "Pos", Left & " " & Top
End Sub

Private Sub cbClose_Click()
    If inMode Then Call DoDelMask(False)
    SaveSetting macroName, OPTSTR, "Limit", cbLimit.text
    SaveSetting macroName, OPTSTR, "Accuracy", cbAccuracy.text
    Unload Me
End Sub

Private Function DoDelMask(DoSel As Boolean)
    ActivePage.Layers("sMask").Delete
    ActivePage.Layers("sEditMode").Delete
    ActiveDocument.ClearSelection
    If DoSel Then srTarget.CreateSelection
    Set srTarget = Nothing
    inMode = False
End Function


Private Sub cbShow_Click()
    If inMode Then Call DoDelMask(True)
    Call ShowTILOverride
End Sub


Private Sub ShowTILOverride()
    If ActiveSelectionRange.Count = 0 Then Exit Sub
    
    sProgress.Width = 0.5
    Me.Height = 54
    'DoEvents
    
    Dim s As Shape, ex As ExportFilter, p$, undLvl&
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrTopLeft
    
    undLvl = 1
    Set srTarget = New ShapeRange
    Set srTarget = ActiveSelectionRange
    
    If ActiveSelectionRange.Count = 1 Then
        Set s = ActiveSelectionRange(1)
        If s.Type <> cdrBitmapShape Then Exit Sub
    Else
        Set s = ActiveSelectionRange.ConvertToBitmapEx(cdrCMYKColorImage, False, False, 150, cdrNormalAntiAliasing, True, True, 95)
        ActiveDocument.ClearSelection
        s.AddToSelection
        undLvl = undLvl + 1
    End If
    
    Dim w&, h&
    Dim resLim#, newRes&
    resLim = CDbl(cbAccuracy.text)
    
    Dim sW#, sh#, sX#, sY#, oPix#
    
    With s.Bitmap
        If .SizeHeight > .SizeWidth Then newRes = CLng(resLim * .ResolutionY / .SizeHeight) Else newRes = CLng(resLim * .ResolutionX / .SizeWidth)
        .Resample .SizeWidth * newRes / .ResolutionX, .SizeHeight * newRes / .ResolutionY, True, newRes, newRes
        
        s.GetPosition sX, sY
        s.GetSize sW, sh
        oPix = sW / s.Bitmap.SizeWidth
        
        sProgress.Width = 2
        'DoEvents

        p = Environ("Temp") & "\" & s.StaticID & ".cpt"
        Set ex = ActiveDocument.ExportBitmap(p, cdrCPT, cdrSelection, cdrCMYKColorImage, _
            .SizeWidth, .SizeHeight, .ResolutionX, .ResolutionY, cdrNoAntiAliasing, False, _
            False, True, False, cdrCompressionNone)
    End With
    ex.Finish
    ActiveDocument.unDo undLvl
    
    sProgress.Width = 5
    'DoEvents
    
    Dim PPisOpened As Boolean, PP As Object, pDoc As Object
    PPisOpened = WindowIsOpen("PhotoPaint")
    
'    If PPisOpened = False Then _
'      Shell """" & SetupPath & "Programs\CorelPP.exe"" -DDE -NoUI"
    
    Set PP = CreateObject("CorelPHOTOPAINT.Application")
    Set pDoc = PP.OpenDocument(p)
    
    Dim hwnd&
    Do: Sleep 5: hwnd = FindWindow("PhotoPaint", vbNullString): DoEvents: Loop While hwnd = 0
    ShowWindow hwnd, SW_MINIMIZE 'SW_HIDE

    '===============
    Dim col As New Collection
    w = pDoc.SizeWidth
    h = pDoc.SizeHeight

    Dim lim&: lim = GetV(cbLimit.text)
    Dim intCount&
    intCount = w * h

    Dim p1&, p2&, c1&, c2&, c3&, c4&, model& '3 = cmyk / 5 = rgb
    
    For p1 = 0 To w - 1
        For p2 = 0 To h - 1
            PP.CorelScript.GetPixelColor p1, p2, model, c1, c2, c3, c4
            If (c1 + c2 + c3 + c4) >= lim Then col.Add "" & p1 & "x" & p2
            sProgress.Width = ((h * p1 + p2) * 100 / intCount) + 5
            'DoEvents
        Next
    Next

    pDoc.Close
    VBA.Kill p
    Set pDoc = Nothing
    If PPisOpened = False Then PP.Quit
    Set PP = Nothing

    '===============
    Optimization = True
    EventsEnabled = False
    
    Dim BGLayer As Layer, sBG As Shape
    Set BGLayer = ActivePage.CreateLayer("sMask")
    Set sBG = BGLayer.CreateRectangle2(0, 0, 7000, 7000)
    sBG.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
    sBG.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, 0)
    sBG.Outline.SetNoOutline
    
    With sBG.Transparency
        .ApplyUniformTransparency 25
        .AppliedTo = cdrApplyToFill
        .MergeMode = cdrMergeNormal
    End With
    
    Dim EMLayer As Layer
    Set EMLayer = ActivePage.CreateLayer("sEditMode")
    EMLayer.Activate
    
    If col.Count > 0 Then
        CreateMask col, sX, sY, oPix
    End If
    
    inMode = True
    ActiveDocument.ClearSelection
    
    EventsEnabled = True
    Optimization = False
    Application.CorelScript.RedrawScreen
    Refresh
    
    Me.Height = 44.25
    'DoEvents
End Sub

Private Sub CreateMask(col As Collection, sX#, sY#, oPix#)
    Dim v As Variant, a$(), rec As Shape, x#, y#
    Dim sr As New ShapeRange
    For Each v In col
        a = Split(v, "x", , vbTextCompare)
        x = sX + (CLng(a(0)) * oPix)
        y = sY - (CLng(a(1)) * oPix)
        Set rec = ActiveLayer.CreateRectangle(x, y, x + oPix, y - oPix)
        rec.Outline.SetNoOutline
        rec.Fill.ApplyUniformFill CreateCMYKColor(0, 100, 100, 0)
        sr.Add rec
    Next
    sr.Combine
End Sub


'Convert To 100
Private Function GetC(c&) As Long
    GetC = CLng(c * 100 / 255)
End Function

'Convert To 255
Private Function GetV(c&) As Long
    GetV = CLng(c * 255 / 100)
End Function

Private Function WindowIsOpen(pstrWindow As String) As Boolean
    WindowIsOpen = (FindWindow(pstrWindow, vbNullString) <> 0)
End Function
