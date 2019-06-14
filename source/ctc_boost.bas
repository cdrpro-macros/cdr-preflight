Attribute VB_Name = "ctc_boost"
Option Explicit

Private MSGBADVER$
Public MSGKEYINFO$, MSGOTHERCOMP$




Public Sub boostStart(Optional ByVal unDo$ = "")
    On Error Resume Next
    If unDo <> "" Then ActiveDocument.BeginCommandGroup unDo
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    End Sub
Public Sub boostFinish(Optional ByVal endUndoGroup% = False)
    On Error Resume Next
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    Application.CorelScript.RedrawScreen
    If endUndoGroup Then ActiveDocument.EndCommandGroup
    Refresh
    End Sub
    



Public Function IsDo(Optional ByVal ChDoc As Boolean = True) As Boolean
    IsDo = False
    Call GetLengStr
    
    If ChVer Then
      Dim needVer$, msg$
      needVer = "X7"
      msg = Replace(MSGBADVER, "$CURRENT", Replace(Version, "Version ", "", , , vbTextCompare), , , vbTextCompare)
      msg = Replace(msg, "$NEED", needVer, , , vbTextCompare)
      MsgBox msg, vbCritical, macroName
      Exit Function
    End If
    
    If ChDoc Then If ActiveDocument Is Nothing Then Exit Function
    IsDo = True
End Function

Private Function ChVer() As Boolean
    ChVer = CBool(VersionMajor < 17)
End Function

Private Sub GetLengStr()
    If UILanguage = cdrRussian Then
        MSGBADVER = "Макрос не совместим с Вашей версией CorelDRAW!" & vbCr & "Ваша версия: $CURRENT" & vbCr & "Минимальная необходимая версия: $NEED"
        MSGKEYINFO = "Для того что бы воспользоваться этой функцией," & vbCr & _
            "необходимо купить регистрационный ключ на сайте " & myWebSite & "!"
        MSGOTHERCOMP = "Регистрационный ключ не подходит для этого компьютера."
    Else
        MSGBADVER = "Macros do not work with this version of CorelDRAW!" & vbCr & "Current version: $CURRENT" & vbCr & "The minimum required version: $NEED"
        MSGKEYINFO = "If you want to use this function" & vbCr & _
            "you have to buy registration key on " & myWebSite & "!"
        MSGOTHERCOMP = "Your registration key does not correspond to this computer."
    End If
End Sub

