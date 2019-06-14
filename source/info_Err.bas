Attribute VB_Name = "info_Err"
Option Explicit

Public errStr$

Public Function myErrLogWr()
  Dim hF&, s$
  hF = FreeFile()
  s = Interaction.Environ("AppData") + "\Corel\" & "SanM_CtC.err"
  Open s For Append As #hF
  Print #hF, "Corel: " & Version
  Print #hF, "Date: " & DateTime.Date$ & " (" & DateTime.time$ & ")"
  Print #hF, "FileName: " & ActiveDocument.Name
  Print #hF, "ErrorCount: " & errCount
  Print #hF, errStr
  Close hF
End Function
        
Public Function myConvErrLogWr()
  Dim hF&, s$
  hF = FreeFile()
  s = Interaction.Environ("AppData") + "\Corel\" & "SanM_CtC_conv.err"
  Open s For Append As #hF
  Print #hF, "Corel: " & Version
  Print #hF, "Date: " & DateTime.Date$ & " (" & DateTime.time$ & ")"
  Print #hF, "FileName: " & ActiveDocument.Name
  Print #hF, "ErrorCount: " & errCount
  Print #hF, errStr
  Close hF
End Function
