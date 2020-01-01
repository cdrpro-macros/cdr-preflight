Attribute VB_Name = "ctc_docker"
Option Explicit



Function ExportList(lst$)
  Dim pth$
  pth = CorelScriptTools.GetFolder("C:\", "Select Folder")
  If pth = "" Then Exit Function
  If Right(pth, 1) <> "\" Then pth = pth & "\"
  
  Dim objRegEx As Object
  Set objRegEx = CreateObject("VBScript.RegExp")
  objRegEx.IgnoreCase = True
  objRegEx.MultiLine = True
  objRegEx.Global = True
  objRegEx.Pattern = "(<P.+?""><A.+?"">|<P.+?"">)(.+?)<.+"
  lst = objRegEx.Replace(lst, "$2<br />")
  
  Dim hF&, i&
  hF = FreeFile()
  pth = pth & ActiveDocument.Name & "_" & "cdrPreflightList.html"
  
  Open pth For Append As #hF
    Print #hF, "<html><head>"
    Print #hF, "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1251"" />"
    Print #hF, "<title>Preflight (" & ActiveDocument.Name & ")</title></head><body>"
    
    Print #hF, "<p>Document: " & ActiveDocument.Name & "<br />"
    Print #hF, "Date: " & VBA.Date$ & "<br />"
    Print #hF, "Time: " & VBA.time$ & "</p>"
    Print #hF, "<p><strong>" & lst & "</strong></p>"
    
    Print #hF, "</body></head>"
  Close hF
  
  MsgBox "File " & pth & " saved!", vbInformation, macroName
End Function




Function cmRefresh(sPreset$) As String
  If myResumeErr Then On Error Resume Next
  
  If IsDo(False) = False Then Exit Function
  If ActiveDocument Is Nothing Then cmRefresh = "<p class=""alert"">Need open a document before!</p>": Exit Function

  'Сканируем документ
  CdrPreflight_start
  
  'Пишем лог
  If (GetSetting(macroName, sREGAPPOPT, "ErrLogSave", "0")) = "1" Then If errCount > 0 Then myErrLogWr

  Dim pr$()
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  If sPreset = "_none" Then
    pr = Split(myDefPreset, "|")
  Else
    pr = Split(GetSetting(macroName, sREGAPPOPT, "Presets" & sPreset), "|")
  End If
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  
  scolSpot = scolSpot + scolPan
  
  
  Dim a$(), c&, a2$(), sHTML$
  a = Split(cdrMyType, "|")
  sHTML = ""

  For c = 1 To UBound(a) + 1

    a2 = Split(pr(c), "-")
    If a2(0) <> "0" Then
      '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      Select Case a(c - 1)
        
        'Layers ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "Layers not visible": If lVis > 0 Then _
          sHTML = sHTML & prCreateItemHtml("lNotVIS", a(c - 1), CStr(lVis))
          
        Case "Layers not printable": If lPrint > 0 Then _
          sHTML = sHTML & prCreateItemHtml("lNotPr", a(c - 1), CStr(lPrint))

        Case "Layers not editable": If lEdit > 0 Then _
          sHTML = sHTML & prCreateItemHtml("lNotEd", a(c - 1), CStr(lEdit))
        
        
        'Shapes ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "OLE shape": If list_OLE.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sOLE", a(c - 1), CStr(list_OLE.Count))
          
        Case "BarCode shape": If sBarCode > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sBarCode", a(c - 1), CStr(sBarCode))
          
        Case "EPS shape": If list_EPS.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sEPS", a(c - 1), CStr(list_EPS.Count))
        
        Case "Symbol shape": If list_symbol.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sSymbol", a(c - 1), CStr(list_symbol.Count))
        
        Case "Perfect shape": If list_PerfSh.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sPerfect", a(c - 1), CStr(list_PerfSh.Count))
        
        Case "Powerclip shape": If list_PoweClip.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sPC", a(c - 1), CStr(list_PoweClip.Count))
        
        Case "Curves not closed": If list_noCloseCur.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sNClCur", a(c - 1), CStr(list_noCloseCur.Count))
        
        Case "Nodes count": If sCurNod > 0 Then _
          sHTML = sHTML & prCreateItemHtml("sNCCount", a(c - 1), CStr(sCurNod))
        
        Case "Shape (Nodes > " & GetSetting(macroName, sREGAPPOPT, "NodesCount", "8000") & ")"
          If list_NodesMax.Count > 0 Then sHTML = sHTML & prCreateItemHtml("sNCMAX", a(c - 1), CStr(list_NodesMax.Count))
        
        
        'Effects ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "Transparency effect": If list_EffTransparency.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eTr", a(c - 1), CStr(list_EffTransparency.Count))
        
        Case "Lens effect": If list_EffLens.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eLens", a(c - 1), CStr(list_EffLens.Count))
        
        Case "Blend effect": If list_EffBlend.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eBlend", a(c - 1), CStr(list_EffBlend.Count))
        
        Case "Drop Shadow effect": If list_EffShadow.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eDrSh", a(c - 1), CStr(list_EffShadow.Count))
        
        Case "Contour effect": If list_EffContour.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eContour", a(c - 1), CStr(list_EffContour.Count))
        
        Case "Control Path effect": If list_ControlPath.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eCPath", a(c - 1), CStr(list_ControlPath.Count))
        
        Case "Bevel effect": If list_EffBevel.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eBevel", a(c - 1), CStr(list_EffBevel.Count))
        
        Case "Artistic Media": If list_EffArtisticMedia.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eAMedia", a(c - 1), CStr(list_EffArtisticMedia.Count))
        
        Case "Extrude effect": If list_EffExtrude.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eExtrude", a(c - 1), CStr(list_EffExtrude.Count))
        
        Case "Envelope effect": If list_EffEnvelope.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eEnvelope", a(c - 1), CStr(list_EffEnvelope.Count))
        
        Case "Distortion effect": If list_EffDistortion.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("eDistort", a(c - 1), CStr(list_EffDistortion.Count))
        
        Case "Perspective effect": If list_EffPerspective.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("ePersp", a(c - 1), CStr(list_EffPerspective.Count))
          
        
        'Text ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "Text": If list_Text.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("tText", a(c - 1), CStr(list_Text.Count))
        
        Case "Overflow Text": If list_TextOver.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("tOver", a(c - 1), CStr(list_TextOver.Count))
        
        Case "Text to path": If tcOnP > 0 Then _
          sHTML = sHTML & prCreateItemHtml("tPath", a(c - 1), CStr(tcOnP))
        
        Case "Small font size (" & GetSetting(macroName, sREGAPPOPT, "SmalFontPt", "6") & "pt)"
          If list_txtSmalPt.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("tFSmal", a(c - 1), CStr(list_txtSmalPt.Count))
            
        Case "Composite color in text (" & GetSetting(macroName, sREGAPPOPT, "SmalFontColor", "12") & "pt)"
          If list_txtSmalCol.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("tSColor", a(c - 1), CStr(list_txtSmalCol.Count))
            
        Case "Table shape": If list_Table.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("tTable", a(c - 1), CStr(list_Table.Count))
        
        
        'Fill ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "Uniform fill": If shuf > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fUniform", a(c - 1), CStr(shuf))
        
        Case "Fountain fill": If list_FountainFill.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fFountain", a(c - 1), CStr(list_FountainFill.Count))
        
        Case "Mid-point inequal of 50": If sFonFillMP > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fMP", a(c - 1), CStr(sFonFillMP))
        
        Case "Hatch fill": If shhf > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fHatch", a(c - 1), CStr(shhf))
        
        Case "Pattern fill": If list_fillPattern.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fPat", a(c - 1), CStr(list_fillPattern.Count))
        
        Case "Postscript fill": If list_fillPS.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fPS", a(c - 1), CStr(list_fillPS.Count))
        
        Case "Texture fill": If list_fillTexture.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fTex", a(c - 1), CStr(list_fillTexture.Count))
        
        Case "Mesh fill": If list_fillMesh.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fMesh", a(c - 1), CStr(list_fillMesh.Count))
        
        Case "No fill": If shnf > 0 Then _
          sHTML = sHTML & prCreateItemHtml("fNO", a(c - 1), CStr(shnf))
        
        
        'Outline ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "Outline": If sOuLineN > 0 Then _
          sHTML = sHTML & prCreateItemHtml("oCount", a(c - 1), CStr(sOuLineN))
        
        Case "Enhanced Outline": If sOuLineEnh > 0 Then _
          sHTML = sHTML & prCreateItemHtml("oECount", a(c - 1), CStr(sOuLineEnh))
        
        Case "Outline (scale with image)": If list_OutLineScal.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("oScale", a(c - 1), CStr(list_OutLineScal.Count))
        
        'Case "Problem outlines": If list_OutlineProbl.Count > 0 Then _
        '  sHTML = sHTML & prCreateItemHtml("oProblem", a(c - 1), CStr(list_OutlineProbl.Count))
        
        Case "Outline Width <= " & GetSetting(macroName, sREGAPPOPT, "OutlineWidthMin", "0.0762")
          If list_OutlineMin.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("oSmall", a(c - 1), CStr(list_OutlineMin.Count))
        
        
        Case "Overprints fill/Outline": If oPrinf + oPrino > 0 Then _
          sHTML = sHTML & prCreateItemHtml("oOverprint", a(c - 1), oPrinf & "/" & oPrino)
          
        'Colors ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        Case "CMYK color": If scolCMYK > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cCMYK", a(c - 1), CStr(scolCMYK))
            
        Case "CMY color": If scolCMY > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cCMY", a(c - 1), CStr(scolCMY))
        
        Case "Spot (PANTONE...)": If scolSpot > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cSPOT", a(c - 1), CStr(scolSpot))
        
        Case "PANTONE Hex": If scolPanH > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cHEX", a(c - 1), CStr(scolPanH))
        
        Case "RGB color": If scolRGB > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cRGB", a(c - 1), CStr(scolRGB))
        
        Case "LAB color": If scolLab > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cLAB", a(c - 1), CStr(scolLab))
        
        Case "HSB color": If scolHSB > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cHSB", a(c - 1), CStr(scolHSB))
        
        Case "HLS color": If scolHLS > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cHLS", a(c - 1), CStr(scolHLS))
        
        Case "YIQ color": If scolYIQ > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cYIQ", a(c - 1), CStr(scolYIQ))
        
        Case "Black and White": If scolBW > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cBW", a(c - 1), CStr(scolBW))
        
        Case "Gray color": If scolGray > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cGRAY", a(c - 1), CStr(scolGray))
        
        Case "Registration color": If scolReg > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cREG", a(c - 1), CStr(scolReg))
        
        Case "Mixed color": If scolMix > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cMIX", a(c - 1), CStr(scolMix))
        
        Case "Multichannel color": If scolMulti > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cMUL", a(c - 1), CStr(scolMulti))
        
        Case "User Ink color": If scolUserInk > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cUSER", a(c - 1), CStr(scolUserInk))
        
        Case "Color Control (min " & GetSetting(macroName, sREGAPPOPT, "myMinColor", "10") & ")": If sColorSmalLim > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cControl", a(c - 1), CStr(sColorSmalLim))
        
        Case "TIL > " & GetSetting(macroName, sREGAPPOPT, "TILFill", "280"): If uColorTIL300 > 0 Then _
          sHTML = sHTML & prCreateItemHtml("cTIL", a(c - 1), CStr(uColorTIL300))
        
        Case "CMYK 400": If scolCMYK100 > 0 Then _
          sHTML = sHTML & prCreateItemHtml("c400", a(c - 1), CStr(scolCMYK100))
        
        
        'Bitmap ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "Bitmap": If list_Allbit.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bCount", a(c - 1), CStr(list_Allbit.Count))
        
        Case "Bitmap < " & GetSetting(macroName, sREGAPPOPT, "myDpiMin", "250") & " dpi"
            If list_MinDPI.Count > 0 Then _
            sHTML = sHTML & prCreateItemHtml("bDPImin", a(c - 1), CStr(list_MinDPI.Count))
            
        Case "Bitmap > " & GetSetting(macroName, sREGAPPOPT, "myDpiMax", "320") & " dpi"
            If list_MaxDPI.Count > 0 Then _
            sHTML = sHTML & prCreateItemHtml("bDPImax", a(c - 1), CStr(list_MaxDPI.Count))
            
        Case "Angle bitmap inequal of 0": If list_BitRot.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bRot", a(c - 1), CStr(list_BitRot.Count))
        
        Case "Unproportional bitmaps": If list_BitXY.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bUnProp", a(c - 1), CStr(list_BitXY.Count))
        
        Case "Crop bitmap On": If list_BitCrop.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bCrop", a(c - 1), CStr(list_BitCrop.Count))
        
        Case "Bitmap link": If list_BitLink.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bLink", a(c - 1), CStr(list_BitLink.Count))
        
        Case "Bitmap transparency": If list_BitTr.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bTr", a(c - 1), CStr(list_BitTr.Count))
        
        Case "Bitmap overprints": If list_BitOverpr.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bOverprint", a(c - 1), CStr(list_BitOverpr.Count))
        
        
        
        'Bitmap Modes ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Case "Gray bitmap": If list_BitGr.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bGray", a(c - 1), CStr(list_BitGr.Count))
        
        Case "Black and White bitmap": If list_BitBW.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bBW", a(c - 1), CStr(list_BitBW.Count))
        
        Case "CMYK color bitmap": If list_BitCMYK.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bCMYK", a(c - 1), CStr(list_BitCMYK.Count))
        
        Case "Duotone bitmap": If list_BitDuo.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bDuoT", a(c - 1), CStr(list_BitDuo.Count))
        
        Case "RGB bitmap": If list_BitRGB.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bRGB", a(c - 1), CStr(list_BitRGB.Count))
        
        Case "CMYK Multichannel bitmap": If bCMYKm > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bCMYKm", a(c - 1), CStr(bCMYKm))
        
        Case "LAB bitmap": If list_BitLAB.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bLAB", a(c - 1), CStr(list_BitLAB.Count))
        
        Case "Paletted bitmap": If list_BitPal.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bPAL", a(c - 1), CStr(list_BitPal.Count))
        
        Case "16 colors bitmap": If b16 > 0 Then _
          sHTML = sHTML & prCreateItemHtml("b16", a(c - 1), CStr(b16))
        
        Case "Spot MultiChannel bitmap": If list_BitDevN.Count > 0 Then _
          sHTML = sHTML & prCreateItemHtml("bDevN", a(c - 1), CStr(list_BitDevN.Count))
            
      End Select
      '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End If

  Next 'c
  
  If Len(sHTML) = 0 Then cmRefresh = "Nothing" Else cmRefresh = sHTML
End Function



Function cmOptions()
  If IsDo(False) = False Then Exit Function
  OptionInfo_Form.Show
End Function

Function cmConvOpt()
  If IsDo(False) = False Then Exit Function
  option_Form.Show
End Function



Function ShowAbout() As String
  Dim s$, sColor$, sDColor$
  Select Case ColorManager.ColorEngine
    Case 0: s = "Kodak"
    Case 1: s = "Microsoft ICM"
    Case 2: s = "Microsoft WCS"
    Case 3: s = "Adobe"
    Case 4: s = "None"
    Case 5: s = "LCMS"
    Case Else: s = "..."
  End Select
  sColor = ""
  sDColor = ""
  If VersionMajor > 14 Then
    sColor = "<br />" & _
    ColorManager.DefaultColorContext.RGBColorProfile.Name & "<br />" & _
    ColorManager.DefaultColorContext.CMYKColorProfile.Name & "<br />" & _
    ColorManager.DefaultColorContext.GrayscaleColorProfile.Name
    
    If Not ActiveDocument Is Nothing Then
      sDColor = "<br />" & "<br />" & "Active Document Profiles:" & "<br />" & _
      ActiveDocument.ColorContext.RGBColorProfile.Name & "<br />" & _
      ActiveDocument.ColorContext.CMYKColorProfile.Name & "<br />" & _
      ActiveDocument.ColorContext.GrayscaleColorProfile.Name
    End If
  End If
  
  ShowAbout = "<p class=""about-info"">" & "Version: " & macroVersion & " " & macroModifyDate & "<br />" & _
    "Copyright " & Chr(169) & " " & macroCopyright & "<br />" & _
    "<br />" & _
    "<a href=""http://" & myWebSite & """ target=""_blank"">CDRPRO.RU</a>" & _
    "<a href=""mailto:" & myEmail & """>" & myEmail & "</a>" & _
    "<br />" & _
    "CorelDRAW " & Version & "<br />" & _
    SetupPath & "<br />" & _
    "<br />" & _
    "Color Engine: " & s & sColor & sDColor & "</p>"
End Function


Function cmLoadPresets() As String
  Dim s$: s = "<select id=""presetsListSel"" onchange=""cmChangePreset()""><option value=""_none"">Default</option>"
  Dim c&, i&, presName$, selected$
  Dim s2$: s2 = ""
  
  On Error GoTo myEnd
  c = GetSetting(macroName, sREGAPPOPT, "PresetsCount", 0)
  selected = GetSetting(macroName, sREGAPPOPT, "PresetsLast", "0")
  For i = 1 To c
    presName = GetSetting(macroName, sREGAPPOPT, "Presets" & i & "Name")
    If presName <> "" Then
      If CLng(selected) = i Then
        s2 = s2 & "<option value=""" & i & """ selected=""selected"">" & presName & "</option>"
      Else
        s2 = s2 & "<option value=""" & i & """>" & presName & "</option>"
      End If
    End If
  Next i
  cmLoadPresets = s & s2 & "</select>"
  Exit Function
myEnd:
  MsgBox "Critical error in LoadPresets. Please, contact to the developer.", vbCritical, macroName & " " & macroVersion
  cmLoadPresets = s & "</select>"
End Function

Function cmChangePreset(id$)
  SaveSetting macroName, sREGAPPOPT, "PresetsLast", IIf(id = "_none", "0", id)
End Function
















'====================================================================================
'============================                          ==============================
'====================================================================================
Private Function prCreateItemHtml(typeID$, label$, value$) As String
  If myResumeErr Then On Error Resume Next
  Dim lb$, lbv$

  lbv = label & " = " & value
  
  If label = "====================" Then
    prCreateItemHtml = "<p class=""separator""> </p>"
  Else
    If isClicked(typeID) Then
      lb = "<a href=""#"" onClick=""LoadList2('" & typeID & "', '" & label & "');"">" & lbv & "</a>"
    Else
      lb = lbv
    End If
    prCreateItemHtml = "<p class=""" & typeID & """ title=""" & lbv & """>" & lb & "</p>"
  End If
End Function


Private Function isClicked(typeID$) As Boolean
  Select Case typeID
    Case "lNotVIS", "lNotPr", "lNotEd", "sBarCode", "sNCCount", "tPath", "fUniform", "fMP", "fHatch", "fNO", "oCount", "oECount", "bCMYKm", "b16": isClicked = False
    Case Else: isClicked = True
  End Select
End Function



















Function LoadList2(typeID$) As String
  Dim srNew As ShapeRange
  Set srNew = New ShapeRange
  
  srObj.RemoveAll
  Select Case typeID
  
    Case "sOLE"
      Set srObj = list_OLE.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "sEPS"
      Set srObj = list_EPS.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "sSymbol"
      Set srObj = list_symbol.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "sPerfect"
      Set srObj = list_PerfSh.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "sPC"
      Set srObj = list_PoweClip.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "sNClCur"
      Set srObj = list_noCloseCur.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "sNCMAX"
      Set srObj = list_NodesMax.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    '===================================================
    
    Case "eTr"
      Set srObj = list_EffTransparency.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eLens"
      Set srObj = list_EffLens.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eBlend"
      Set srObj = list_EffBlend.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eDrSh"
      Set srObj = list_EffShadow.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eContour"
      Set srObj = list_EffContour.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eCPath"
      Set srObj = list_ControlPath.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eBevel"
      Set srObj = list_EffBevel.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eAMedia"
      Set srObj = list_EffArtisticMedia.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eExtrude"
      Set srObj = list_EffExtrude.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eEnvelope"
      Set srObj = list_EffEnvelope.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "eDistort"
      Set srObj = list_EffDistortion.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    Case "ePersp"
      Set srObj = list_EffPerspective.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    '===================================================
    
    Case "tText"
      Set srObj = list_Text.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "tOver"
      Set srObj = list_TextOver.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "tFSmal"
      Set srObj = list_txtSmalPt.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "tSColor"
      Set srObj = list_txtSmalCol.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "tTable"
      Set srObj = list_Table.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    '===================================================
      
    Case "fFountain"
      Set srObj = list_FountainFill.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "fPat"
      Set srObj = list_fillPattern.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "fPS"
      Set srObj = list_fillPS.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "fTex"
      Set srObj = list_fillTexture.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "fMesh"
      Set srObj = list_fillMesh.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    '===================================================
    
    Case "oScale"
      Set srObj = list_OutLineScal.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "oSmall"
      Set srObj = list_OutlineMin.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
    
    '===================================================
    
    Case "oOverprint"
      OverprintFillView2 list_CanFillOutline, srNew
      Set srObj = srNew.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    '===================================================
    
    Case "cCMYK": LoadList2 = CreateShapeRange2(typeID, cdrColorCMYK)
    Case "cCMY": LoadList2 = CreateShapeRange2(typeID, cdrColorCMY)
      
    Case "cSPOT"
      SpotView list_CanFillOutline, srNew
      Set srObj = srNew.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "cHEX": LoadList2 = CreateShapeRange2(typeID, cdrColorPantoneHex)
    Case "cRGB": LoadList2 = CreateShapeRange2(typeID, cdrColorRGB)
    Case "cLAB": LoadList2 = CreateShapeRange2(typeID, cdrColorLab)
    Case "cHSB": LoadList2 = CreateShapeRange2(typeID, cdrColorHSB)
    Case "cHLS": LoadList2 = CreateShapeRange2(typeID, cdrColorHLS)
    Case "cYIQ": LoadList2 = CreateShapeRange2(typeID, cdrColorYIQ)
    Case "cBW": LoadList2 = CreateShapeRange2(typeID, cdrColorBlackAndWhite)
    Case "cGRAY": LoadList2 = CreateShapeRange2(typeID, cdrColorGray)
    Case "cREG": LoadList2 = CreateShapeRange2(typeID, cdrColorRegistration)
    Case "cMIX": LoadList2 = CreateShapeRange2(typeID, cdrColorMixed)
    Case "cMUL": LoadList2 = CreateShapeRange2(typeID, cdrColorMultiChannel)
    Case "cUSER": LoadList2 = CreateShapeRange2(typeID, cdrColorUserInk)
    
    Case "cControl"
      Set srObj = list_ColorSmalLim.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "cTIL"
      san_TILView list_CanFillOutline, srNew
      Set srObj = srNew.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "c400"
      san_CMYK100View list_CanFillOutline, srNew
      Set srObj = srNew.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    '===================================================
    
    Case "bCount"
      Set srObj = list_Allbit.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bDPImin"
      Set srObj = list_MinDPI.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bDPImax"
      Set srObj = list_MaxDPI.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bRot"
      Set srObj = list_BitRot.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bUnProp"
      Set srObj = list_BitXY.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bCrop"
      Set srObj = list_BitCrop.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bLink"
      Set srObj = list_BitLink.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bTr"
      Set srObj = list_BitTr.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bOverprint"
      Set srObj = list_BitOverpr.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    '===================================================
    
    Case "bGray"
      Set srObj = list_BitGr.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bBW"
      Set srObj = list_BitBW.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bCMYK"
      Set srObj = list_BitCMYK.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bDuoT"
      Set srObj = list_BitDuo.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bRGB"
      Set srObj = list_BitRGB.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    'Case "bCMYKm"
    
    Case "bLAB"
      Set srObj = list_BitLAB.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case "bPAL"
      Set srObj = list_BitPal.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    'Case "b16"
    
    Case "bDevN"
      Set srObj = list_BitDevN.ReverseRange
      LoadList2 = DrawItemsForSecondList(typeID)
      
    Case Else
      LoadList2 = "<p>...</p>"
    
  End Select
End Function




Private Function CreateShapeRange2(typeID$, cmdl As cdrColorType) As String
  Dim srNew As ShapeRange
  Set srNew = New ShapeRange
  Call san_ColorModeView(list_CanFillOutline, srNew, cmdl)
  Set srObj = srNew.ReverseRange
  CreateShapeRange2 = DrawItemsForSecondList(typeID)
End Function




Private Function DrawItemsForSecondList(typeID$) As String
  Dim c&, s As Shape, n$, ef As Effect, fc As FountainColor, fc2&, til&
  Dim html$, htmItm$, stl$
  html = "": fc2 = 0
  
  myUnit = myDoc.Unit
  myDoc.Unit = myUnitWork
  
  For c = 1 To srObj.Count
    Set s = srObj(c)
    
    htmItm = ""
    
    Select Case typeID
      
      Case "sOLE": n = s.OLE.ProgID & " - " & s.Layer.Page.Name
      Case "sEPS": n = " EPS File - " & s.Layer.Page.Name
      Case "sSymbol": n = " Symbol - " & s.Layer.Page.Name
      Case "sPerfect": n = " Perfect Shape - " & s.Layer.Page.Name
      Case "sPC": n = myShapeType(s) & " (" & s.PowerClip.Shapes.Count & " obj)" & " - " & s.Layer.Page.Name
      Case "sNClCur": n = " Curve (" & Round(s.Outline.Width, 3) & ") - " & s.Layer.Page.Name
      Case "sNCMAX": n = " Curve" & " - " & s.Layer.Page.Name
      
      '=======================================================================
      
      Case "eTr"
        If s.Transparency.Type <> cdrNoTransparency Then
          Select Case s.Transparency.Type
            Case cdrUniformTransparency: n = " Uniform " & s.Transparency.Uniform
            Case cdrFountainTransparency: n = " Fountain"
            Case cdrPatternTransparency: n = " Pattern"
            Case cdrTextureTransparency: n = " Texture"
          End Select
        Else
          n = ""
        End If
        
        If s.CanHaveFill And s.Fill.Type = cdrFountainFill Then
          Dim ffColor As FountainColor
          For Each ffColor In s.Fill.Fountain.Colors
            If ffColor.Opacity < 255 Then
              If n = "" Then n = "Fill Opacity" Else n = n & "|Fill Opacity"
              Exit For
            End If
          Next
        End If
        
        n = n & " - " & s.Layer.Page.Name
            
      Case "eLens"
        For Each ef In s.Effects
          If ef.Type = cdrLens Then
              Select Case ef.Lens.Type
              Case cdrLensBrighten: n = " Brighten"
              Case cdrLensColorAdd: n = " ColorAdd"
              Case cdrLensColorLimit: n = " ColorLimit"
              Case cdrLensCustomColorMap: n = " ColorMap"
              Case cdrLensFishEye: n = " FishEye"
              Case cdrLensHeatMap: n = " HeatMap"
              Case cdrLensInvert: n = " Invert"
              Case cdrLensMagnify: n = " Magnify"
              Case cdrLensTintedGrayscale: n = " TintedGray"
              Case cdrLensWireframe: n = " Wireframe"
              End Select
          End If
        Next ef
        n = n & " - " & s.Layer.Page.Name
        
      Case "eBlend": n = " Blend - " & s.Layer.Page.Name
      Case "eDrSh": n = " Drop Shadow - " & s.Layer.Page.Name
      Case "eContour": n = " Contour - " & s.Layer.Page.Name
      Case "eCPath": n = " Control Path - " & s.Layer.Page.Name
      Case "eBevel": n = " Bevel - " & s.Layer.Page.Name
      Case "eAMedia": n = " Artistic Media - " & s.Layer.Page.Name
      Case "eExtrude": n = " Extrude - " & s.Layer.Page.Name
      Case "eEnvelope": n = " Envelope - " & s.Layer.Page.Name
      Case "eDistort": n = " Distortion - " & s.Layer.Page.Name
      Case "ePersp": n = " Perspective - " & s.Layer.Page.Name
      
      '=======================================================================
      
      Case "tText"
        If s.text.IsArtisticText Then n = " Artistic" Else n = " Paragraph"
        Select Case s.text.Story.LanguageID
          Case cdrRussian: n = n + " (RUS)"
          Case cdrEnglishUS: n = n + " (US)"
          Case cdrEnglishUK: n = n + " (UK)"
          'Case Else: n = n + " (...)"
        End Select
        n = n & " - " & s.Layer.Page.Name
        
      Case "tOver": n = " Text " & " - " & s.Layer.Page.Name
      Case "tFSmal"
        If s.text.IsArtisticText Then n = " Artistic" Else n = " Paragraph"
        n = n & mySmallFontSize(s)
      Case "tSColor"
        If s.text.IsArtisticText Then n = " Artistic" Else n = " Paragraph"
        n = n & mySmallFontColor(s)
      Case "tTable": n = " Table" & " - " & s.Layer.Page.Name
      
      '=======================================================================
      
      Case "fFountain": n = myShapeType(s) & " (" & CStr(s.Fill.Fountain.Colors.Count) & " colors)" & " - " & s.Layer.Page.Name
      Case "fPat": n = myShapeType(s) & " - " & s.Layer.Page.Name
      Case "fPS": n = myShapeType(s) & " - " & s.Layer.Page.Name
      Case "fTex": n = myShapeType(s) & " (" & s.Fill.Texture.Resolution & " dpi)" & " - " & s.Layer.Page.Name
      Case "fMesh": n = " Mesh fill object" & " - " & s.Layer.Page.Name
      
      '=======================================================================
      
      Case "oScale": n = " Outline " & "(" & Round(s.Outline.Width, 3) & ")" & " - " & s.Layer.Page.Name
      Case "oSmall": n = " Outline " & "(" & Round(s.Outline.Width, 3) & ")" & " - " & s.Layer.Page.Name
      
      '=======================================================================
      
      Case "oOverprint"
        n = myShapeType(s)
        If s.CanHaveFill And s.OverprintFill Then n = n + " (F)"
        If s.CanHaveOutline And s.OverprintOutline Then n = n + " (O)"
        n = Replace(n, "F) (O", "F & O")
        n = n & " - " & s.Layer.Page.Name

      '=======================================================================
      
      Case "cCMYK": htmItm = GetStrForColor(s, c, cdrColorCMYK)
      Case "cCMY": htmItm = GetStrForColor(s, c, cdrColorCMY)
      Case "cSPOT": htmItm = GetStrForColor2(s, c, cdrColorPantone, cdrColorSpot)
      Case "cHEX": htmItm = GetStrForColor(s, c, cdrColorPantoneHex)
      Case "cRGB": htmItm = GetStrForColor(s, c, cdrColorRGB)
      Case "cLAB": htmItm = GetStrForColor(s, c, cdrColorLab)
      Case "cHSB": htmItm = GetStrForColor(s, c, cdrColorHSB)
      Case "cHLS": htmItm = GetStrForColor(s, c, cdrColorHLS)
      Case "cYIQ": htmItm = GetStrForColor(s, c, cdrColorYIQ)
      Case "cBW": n = myShapeType(s): n = n & myShapeVectorColor(s, cdrColorBlackAndWhite)
      Case "cGRAY": htmItm = GetStrForColor(s, c, cdrColorGray)
      Case "cREG": n = myShapeType(s): n = n & myShapeVectorColor(s, cdrColorRegistration)
      Case "cMIX": n = myShapeType(s): n = n & myShapeVectorColor(s, cdrColorMixed)
      Case "cMUL": n = myShapeType(s): n = n & myShapeVectorColor(s, cdrColorMultiChannel)
      Case "cUSER": htmItm = GetStrForColor(s, c, cdrColorUserInk)
      
      '=======================================================================
      
      Case "cControl"
        If s.CanHaveFill Then
          Select Case s.Fill.Type
            Case cdrUniformFill
              If s.Fill.UniformColor.Type = cdrColorCMYK Then
                If scanColorSmLim2(s.Fill.UniformColor) = True Then
                  stl = GetColorClass2(s.Fill.UniformColor)
                  n = "F: " & s.Fill.UniformColor.Name(True)
                  htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
                End If
              End If
            Case cdrFountainFill
              For Each fc In s.Fill.Fountain.Colors
                If fc.Color.Type = cdrColorCMYK Then
                  If scanColorSmLim2(fc.Color) = True Then
                    stl = GetColorClass2(fc.Color)
                    n = "FF: " & fc.Color.Name(True)
                    htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
                  End If
                End If
              Next
          End Select
        End If
        
        If s.CanHaveOutline Then
          If s.Outline.Type <> cdrNoOutline Then
            If s.Outline.Color.Type = cdrColorCMYK Then
              If scanColorSmLim2(s.Outline.Color) = True Then
                stl = GetColorClass2(s.Outline.Color)
                n = "O: " & s.Outline.Color.Name(True)
                htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
              End If
            End If
          End If
        End If
        
      '=======================================================================
           
      Case "cTIL"
        til = GetSetting(macroName, sREGAPPOPT, "TILFill", "280")
        If s.CanHaveFill Then
          Select Case s.Fill.Type
            Case cdrUniformFill
              If s.Fill.UniformColor.Type = cdrColorCMYK Then
                If s.Fill.UniformColor.CMYKCyan + s.Fill.UniformColor.CMYKMagenta + s.Fill.UniformColor.CMYKYellow + s.Fill.UniformColor.CMYKBlack >= til Then
                  stl = GetColorClass2(s.Fill.UniformColor)
                  n = "F: " & s.Fill.UniformColor.Name(True) & " %" & (s.Fill.UniformColor.CMYKCyan + s.Fill.UniformColor.CMYKMagenta + s.Fill.UniformColor.CMYKYellow + s.Fill.UniformColor.CMYKBlack)
                  htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
                End If
              End If
            Case cdrFountainFill
              For Each fc In s.Fill.Fountain.Colors
                If fc.Color.Type = cdrColorCMYK Then
                  If fc.Color.CMYKCyan + fc.Color.CMYKMagenta + fc.Color.CMYKYellow + fc.Color.CMYKBlack >= til Then
                    stl = GetColorClass2(fc.Color)
                    n = "FF: " & fc.Color.Name(True) & " %" & (fc.Color.CMYKCyan + fc.Color.CMYKMagenta + fc.Color.CMYKYellow + fc.Color.CMYKBlack)
                    htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
                  End If
                End If
              Next
          End Select
        End If
      
        If s.CanHaveOutline Then
          If s.Outline.Color.Type = cdrColorCMYK Then
            If s.Outline.Color.CMYKCyan + s.Outline.Color.CMYKMagenta + s.Outline.Color.CMYKYellow + s.Outline.Color.CMYKBlack >= til Then
              stl = GetColorClass2(s.Outline.Color)
              n = "O: " & s.Outline.Color.Name(True) & " %" & (s.Outline.Color.CMYKCyan + s.Outline.Color.CMYKMagenta + s.Outline.Color.CMYKYellow + s.Outline.Color.CMYKBlack)
              htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
            End If
          End If
        End If
       
      '=======================================================================
            
      Case "c400"
        If s.CanHaveFill Then
          Select Case s.Fill.Type
            Case cdrUniformFill
              If s.Fill.UniformColor.Type = cdrColorCMYK Then
                With s.Fill.UniformColor
                  If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then
                    stl = GetColorClass2(s.Fill.UniformColor)
                    n = "F: " & .Name(True)
                    htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
                  End If
                End With
              End If
            Case cdrFountainFill
              For Each fc In s.Fill.Fountain.Colors
                If fc.Color.Type = cdrColorCMYK Then
                  With fc.Color
                    If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then
                      stl = GetColorClass2(fc.Color)
                      n = "FF: " & .Name(True)
                      htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
                    End If
                  End With
                End If
              Next
          End Select
        End If
      
        If s.CanHaveOutline Then
          If s.Outline.Color.Type = cdrColorCMYK Then
            With s.Outline.Color
              If .CMYKCyan > 0 And .CMYKMagenta > 0 And .CMYKYellow > 0 And .CMYKBlack > 0 Then
                stl = GetColorClass2(s.Outline.Color)
                n = "O: " & .Name(True)
                htmItm = htmItm & "<p title=""" & n & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
              End If
            End With
          End If
        End If
        
      '=======================================================================
        
      Case "bCount": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bDPImin": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bDPImax": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bRot": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bUnProp": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bCrop": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bLink": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bTr": n = myBitmapType(s): n = n & myBitmapResol(s)
      Case "bOverprint": n = myBitmapType(s): n = n & myBitmapResol(s)
      
      '=======================================================================
      
      Case "bGray": n = myBitmapResol(s)
      Case "bBW": n = myBitmapResol(s)
      Case "bCMYK": n = myBitmapResol(s)
      Case "bDuoT": n = myBitmapResol(s)
      Case "bRGB": n = myBitmapResol(s)
      'Case "bCMYKm": n = myBitmapResol(s)
      Case "bLAB": n = myBitmapResol(s)
      Case "bPAL": n = myBitmapResol(s)
      'Case "b16": n = myBitmapResol(s)
      Case "bDevN": n = myBitmapResol(s)
        
    End Select
    
    If htmItm = "" Then
      html = html & "<p title=""" & n & """><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
    Else
      html = html & htmItm
    End If
    
  Next
  
  myDoc.Unit = myUnit
  DrawItemsForSecondList = html
End Function


'for color
Private Function GetStrForColor(s As Shape, c&, clr As cdrColorType) As String
  If myResumeErr Then On Error Resume Next
  Dim htm$, n$, tt$, fc As FountainColor, stl$
  htm = ""
  
  If s.CanHaveFill Then
    Select Case s.Fill.Type
      Case cdrUniformFill
        If s.Fill.UniformColor.Type = clr Then
          stl = GetColorClass2(s.Fill.UniformColor)
          If clr = cdrColorPantoneHex Then
            n = "F: " & Replace(s.Fill.UniformColor.Name(False), "Hexachrome", "Hex", , , vbTextCompare)
            tt = n
          Else
            n = "F: " & s.Fill.UniformColor.Name(True)
            tt = s.Fill.UniformColor.Name(False) & " (" & s.Fill.UniformColor.Name(True) & ")"
          End If
          htm = htm & "<p title=""" & tt & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
        End If
      Case cdrFountainFill
        For Each fc In s.Fill.Fountain.Colors
            If fc.Color.Type = clr Then
              stl = GetColorClass2(fc.Color)
              If clr = cdrColorPantoneHex Then
                n = "FF: " & Replace(fc.Color.Name(False), "Hexachrome", "Hex", , , vbTextCompare)
                tt = n
              Else
                n = "FF: " & fc.Color.Name(True)
                tt = fc.Color.Name(False) & " (" & fc.Color.Name(True) & ")"
              End If
              htm = htm & "<p title=""" & tt & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
            End If
        Next
    End Select
  End If

  If s.CanHaveOutline Then
    If s.Outline.Color.Type = clr Then
      stl = GetColorClass2(s.Outline.Color)
      If clr = cdrColorPantoneHex Then
        n = "O: " & Replace(s.Outline.Color.Name(False), "Hexachrome", "Hex", , , vbTextCompare)
        tt = n
      Else
        n = "O: " & s.Outline.Color.Name(True)
        tt = s.Outline.Color.Name(False) & " (" & s.Outline.Color.Name(True) & ")"
      End If
      htm = htm & "<p title=""" & tt & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
    End If
  End If

  GetStrForColor = htm
End Function


'for spot
Private Function GetStrForColor2(s As Shape, c&, clr1 As cdrColorType, clr2 As cdrColorType) As String
  If myResumeErr Then On Error Resume Next
  Dim htm$, n$, n2$, fc As FountainColor, stl$
  htm = ""
  
  If s.CanHaveFill Then
    Select Case s.Fill.Type
      Case cdrUniformFill
        If s.Fill.UniformColor.Type = clr1 Or s.Fill.UniformColor.Type = clr2 Then
          stl = GetColorClass2(s.Fill.UniformColor)
          n = "F: " & s.Fill.UniformColor.Name(False)
          n2 = n & " / " & s.Fill.UniformColor.Name(True)
          n = n & " %" & s.Fill.UniformColor.Tint
          htm = htm & "<p title=""" & n2 & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
        End If
      Case cdrFountainFill
        For Each fc In s.Fill.Fountain.Colors
            If fc.Color.Type = clr1 Or fc.Color.Type = clr2 Then
             stl = GetColorClass2(fc.Color)
             n = "FF: " & fc.Color.Name(False)
             n2 = n & " / " & fc.Color.Name(True)
             n = n & " %" & fc.Color.Tint
             htm = htm & "<p title=""" & n2 & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
            End If
        Next
    End Select
  End If

  If s.CanHaveOutline Then
    If s.Outline.Color.Type = clr1 Or s.Outline.Color.Type = clr2 Then
      stl = GetColorClass2(s.Outline.Color)
      n = "O: " & s.Outline.Color.Name(False)
      n2 = n & " / " & s.Outline.Color.Name(True)
      n = n & " %" & s.Outline.Color.Tint
      htm = htm & "<p title=""" & n2 & """ class=""NoIco""><b style=""" & stl & """ class=""clr""> </b><a href=""#"" onClick=""SelectShape('" & c & "');"">" & n & "</a></p>"
    End If
  End If

  GetStrForColor2 = htm
End Function


Private Function myShapeType(s As Shape) As String
  If myResumeErr Then On Error Resume Next
  Select Case s.Type
    Case 1: myShapeType = "Rectangle"
    Case 2: myShapeType = "Ellipse"
    Case 3: myShapeType = "Curve Shape"
    Case 4: myShapeType = "Polygon Shape"
    Case 5: myShapeType = "Bitmap"
    Case 6: myShapeType = "Text"
    Case 19: myShapeType = "Connector"
    Case 20: myShapeType = "MeshFill"
    Case 26: myShapeType = "Perfect Shape"
  End Select
  If myShapeType = "" Then myShapeType = "Shape"
End Function

Private Function myShapeVectorColor$(s As Shape, cType As cdrColorType)
  If myResumeErr Then On Error Resume Next
  Dim n$, fc As FountainColor, fc2&: fc2 = 0: n = ""
  If s.CanHaveFill Then
      Select Case s.Fill.Type
      Case cdrUniformFill: If s.Fill.UniformColor.Type = cType Then n = n + " (F)"
      Case cdrFountainFill
          For Each fc In s.Fill.Fountain.Colors
              If fc.Color.Type = cType Then fc2 = fc2 + 1
          Next
          If fc2 > 0 Then n = n & " (FF " & fc2 & "c.)"
      End Select
  End If
  '=====================
  If s.CanHaveOutline Then
      If s.Outline.Color.Type = cType Then n = n + " (O)"
  End If
  '=====================
  n = Replace(n, "F) (O", "F & O")
  n = Replace(n, "c.) (O", "c. & O")
  myShapeVectorColor = n & " - " & s.Layer.Page.Name
End Function

Private Function GetColorClass2(c As Color) As String
  GetColorClass2 = "background-color:" & GetColorValue(c) & ";"
End Function

Private Function GetColorValue(c As Color) As String
  Dim cl As New Color
  cl.CopyAssign c
  If cl.Type <> cdrColorRGB Then cl.ConvertToRGB
  GetColorValue = "RGB(" & cl.RGBRed & ", " & cl.RGBGreen & ", " & cl.RGBBlue & ")"
End Function


Private Function scanColorSmLim2(c As Color) As Boolean
  Dim cl&: cl = GetSetting(macroName, sREGAPPOPT, "myMinColor", "10")
  If c.CMYKCyan > 0 And c.CMYKCyan < cl Then scanColorSmLim2 = True: Exit Function
  If c.CMYKMagenta > 0 And c.CMYKMagenta < cl Then scanColorSmLim2 = True: Exit Function
  If c.CMYKYellow > 0 And c.CMYKYellow < cl Then scanColorSmLim2 = True: Exit Function
  If c.CMYKBlack > 0 And c.CMYKBlack < cl Then scanColorSmLim2 = True
End Function






Function CountList2() As String
  CountList2 = CStr(srObj.Count)
End Function


Function SelectShapeByItem(i&)
  If myResumeErr Then On Error Resume Next
  Dim k&: k = GetKeySt
  Select Case k
    Case 0, 1
      If srObj.Item(i).Layer.Master = False Then srObj.Item(i).Layer.Page.Activate
      ActiveDocument.ClearSelection
      srObj.Item(i).CreateSelection
      Application.Refresh
      If k = 1 Then Call SelectShapeAndZoom(i)
    Case 3
      Call SelectAllShapes(i)
  End Select
End Function

Private Sub SelectShapeAndZoom(i&)
  Const MARGIN = 100
  Dim w#, h#, x#, y#, vaX#, vaY#, vaW#, vaH#
  If ActiveWindow Is Nothing Then Beep: Exit Sub
  If myResumeErr Then On Error Resume Next
  ActiveWindow.ActiveView.GetViewArea vaX, vaY, vaW, vaH
  With srObj.Item(i)
      .GetBoundingBox x, y, w, h
      x = x + (w / 2): y = y + (h / 2)
      w = w * (100 + MARGIN) / 100
      h = h * (100 + MARGIN) / 100
      ActiveWindow.ActiveView.SetViewArea x - w / 2, y - h / 2, w, h
  End With
End Sub

Private Sub SelectAllShapes(i&)
  If myResumeErr Then On Error Resume Next
  Dim s As Shape, sr3 As New ShapeRange, p As Page
  Set p = srObj.Item(i).Layer.Page

  ActiveDocument.ClearSelection
  sr3.RemoveAll
  For Each s In srObj
      If s.Layer.Page.Name = p.Name Then sr3.Add s
  Next s
  sr3.CreateSelection
  ActiveWindow.ActiveView.ToFitShapeRange sr3
End Sub


Private Function GetKeySt() As Long
  If (GetKeyState(vbKeyShift) And &HFF80) <> 0 = True Then
    GetKeySt = 2
  ElseIf (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then
    GetKeySt = 1
  ElseIf (GetKeyState(vbKeyA) And &HFF80) <> 0 = True Then
    GetKeySt = 3 'A
  Else
    GetKeySt = 0
  End If
End Function






'====================================================================================
'==============    Определение цветовой модели Битмапа для Листа     ================
'====================================================================================
Private Function myBitmapType$(s As Shape)
  If myResumeErr Then On Error Resume Next
  Select Case s.Bitmap.Mode
    Case cdrRGBColorImage: myBitmapType = "RGB"
    Case cdrPalettedImage: myBitmapType = "PAL"
    Case cdrLABImage: myBitmapType = "LAB"
    Case cdrGrayscaleImage: myBitmapType = "GRAY"
    Case cdrDuotoneImage: myBitmapType = "DUO"
    Case cdrCMYKMultiChannelImage: myBitmapType = "CMYKm"
    Case cdrCMYKColorImage: myBitmapType = "CMYK"
    Case cdrBlackAndWhiteImage: myBitmapType = "BW"
    Case cdr16ColorsImage: myBitmapType = "16c"
    Case cdrRGBMultiChannelImage: myBitmapType = "RGBm"
    Case cdrSpotMultiChannelImage: myBitmapType = "MultiCh"
  End Select
End Function
        
Private Function myBitmapResol$(s As Shape)
  If myResumeErr Then On Error Resume Next
  myBitmapResol = myBitmapResol & " (" & s.Bitmap.ResolutionX & " x " & s.Bitmap.ResolutionY & " dpi)"
  myBitmapResol = myBitmapResol & " - " & s.Layer.Page.Name
End Function



