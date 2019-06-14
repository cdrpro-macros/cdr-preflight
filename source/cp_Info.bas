Attribute VB_Name = "cp_Info"
Option Explicit


Sub ShowCdrPreflightDocker()
  Application.FrameWork.ShowDocker "34695a15-b045-1b43-96a4-6e5eee9679c7"
End Sub

Function InitializeCdrPreflight()
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  cdrMyType = "Layers not visible|Layers not printable|Layers not editable|li1|OLE shape|BarCode shape|EPS shape|Symbol shape|Perfect shape|Powerclip shape|Curves not closed|Nodes count" & _
  "|Shape (Nodes > " & GetSetting(macroName, sREGAPPOPT, "NodesCount", "8000") & ")|li2|Transparency effect|Lens effect|Blend effect|Drop Shadow effect|Contour effect|Control Path effect|Bevel effect" & _
  "|Artistic Media|Extrude effect|Envelope effect|Distortion effect|Perspective effect|li3|Text|Overflow Text|Text to path|Small font size (" & GetSetting(macroName, sREGAPPOPT, "SmalFontPt", "6") & "pt)" & _
  "|Composite color in text (" & GetSetting(macroName, sREGAPPOPT, "SmalFontColor", "12") & "pt)|Table shape|li4|Uniform fill|Fountain fill|Mid-point inequal of 50|Hatch fill|Pattern fill|Postscript fill|Texture fill|Mesh fill" & _
  "|No fill|Outline|Enhanced Outline|Outline (scale with image)|Problem outlines|Outline Width <= " & GetSetting(macroName, sREGAPPOPT, "OutlineWidthMin", "0.0762") & "|Overprints fill/Outline|li5|CMYK color|CMY color|Spot (PANTONE...)|PANTONE Hex|" & _
  "RGB color|LAB color|HSB color|HLS color|YIQ color|Black and White|Gray color|Registration color|Mixed color|Multichannel color|User Ink color|Color Control " & "(min " & GetSetting(macroName, sREGAPPOPT, "myMinColor", "10") & ")" & _
  "|TIL > " & GetSetting(macroName, sREGAPPOPT, "TILFill", "280") & "|CMYK 400|li6|Bitmap|Bitmap < " & GetSetting(macroName, sREGAPPOPT, "myDpiMin", "250") & " dpi|Bitmap > " & _
  GetSetting(macroName, sREGAPPOPT, "myDpiMax", "320") & " dpi|Angle bitmap inequal of 0|Unproportional bitmaps|Crop bitmap On|Bitmap link|Bitmap transparency|Bitmap overprints|Gray bitmap|Black and White bitmap|CMYK color bitmap|Duotone bitmap|RGB bitmap|CMYK Multichannel bitmap|LAB bitmap|Paletted bitmap|16 colors bitmap|Spot MultiChannel bitmap"
End Function




'====================================================================================
'==============================         Запуск ИНФО        ==========================
'====================================================================================
Function CdrPreflight_start()
  Dim myShapeRange As ShapeRange
  Set myShapeRange = New ShapeRange
  
  Dim myOperatingMode$
  myOperatingMode = GetSetting(macroName, sREGAPPOPT, "OperatingMode", "Default")
  
  If cdrMyType = "" Then InitializeCdrPreflight
  
  Set myDoc = ActiveDocument
  Set myOldPage = ActivePage
  'Set myActiveLayer = ActiveLayer
  Set myMasterPage = myDoc.MasterPage
  
  myUnit = myDoc.Unit
  myDoc.Unit = myUnitWork
  
  Select Case myOperatingMode
    Case "Default": infoDefaultMode
    Case "In select"
        Set myShapeRange = ActiveSelectionRange
        If myShapeRange.Count > 0 Then
        infoSelectMode ActiveSelectionRange, True
        ElseIf myShapeRange.Count = 0 Then
        myBeforeWork
        'InfoForm.Show 0
        End If
    Case "On Active page": infoActivePageMode
    Case "On Master pages": infoMasterPageMode True
    Case "All page": infoAllPageMode True
  End Select
  
  myDoc.Unit = myUnit
  
End Function












'====================================================================================
'===================         Режим по умолчанию (Default)        ====================
'====================================================================================
Private Sub infoDefaultMode()
  Dim myShapeRange As ShapeRange, mySelShepes As ShapeRange
  Set myShapeRange = New ShapeRange
  Set mySelShepes = New ShapeRange
  
  Set myShapeRange = ActiveSelectionRange
  Set mySelShepes = ActiveSelectionRange
  
  myBeforeWork
  
  If myShapeRange.Count = 0 Then
      boostStart "INFO"
      Application.Status.BeginProgress "Get INFO...", False
      
      myLayerScan
      myFindShapesMasterCount
      myFindShapesCount
      
      infoMasterPageMode False
      infoAllPageMode False
      
      Application.Status.EndProgress
      boostFinish endUndoGroup:=True
      myDoc.unDo
      myDoc.ClearSelection
  Else
  infoSelectMode myShapeRange, False
  myDoc.ClearSelection: mySelShepes.CreateSelection
  End If
  myAfterWork
End Sub



'====================================================================================
'=======================     Режим в выделенном (Select)     ========================
'====================================================================================
Private Sub infoSelectMode(myShapeRange As ShapeRange, myLevelWork As Boolean)
  Dim myShapeRange2 As ShapeRange, mySelShepes As ShapeRange
  Set myShapeRange2 = New ShapeRange
  Set mySelShepes = New ShapeRange
  
  If myLevelWork Then myBeforeWork: Set mySelShepes = myShapeRange
  
  boostStart "INFO"
  Application.Status.BeginProgress "Get INFO...", False
  
  myCountS = myShapeRange.Count
  infoDocStart2 myShapeRange
  
  Application.Status.EndProgress
  boostFinish endUndoGroup:=True
  myDoc.unDo
  
  If myLevelWork Then myAfterWork: myDoc.ClearSelection: mySelShepes.CreateSelection
End Sub



'====================================================================================
'=============================     Режим Master Page     ============================
'====================================================================================
Private Sub infoMasterPageMode(myLevelWork As Boolean)
  Dim l As Layer
  
  If myLevelWork Then myBeforeWork: _
  boostStart "INFO": myLayerScan: myFindShapesMasterCount: _
  Application.Status.BeginProgress "Get INFO...", False
  
  myMasterPage.UnlockAllShapes
  For Each l In myMasterPage.Layers
      If l.IsGuidesLayer = False And l.IsGridLayer = False Then _
      If l.Master = True Then infoDocStart2 l.Shapes.All
  Next l
  
  If myLevelWork Then Application.Status.EndProgress: _
  boostFinish endUndoGroup:=True: myDoc.unDo: myDoc.ClearSelection: myAfterWork
End Sub



'====================================================================================
'==============================     Режим All Page     ==============================
'====================================================================================
Private Sub infoAllPageMode(myLevelWork As Boolean)
  Dim l As Layer, c&, lc&
  
  If myLevelWork Then myBeforeWork: _
  boostStart "INFO": myLayerScan: myFindShapesCount: _
  Application.Status.BeginProgress "Get INFO...", False
  
  lc = myOldPage.index
  For Each l In myOldPage.Layers
      If l.IsSpecialLayer = False Then infoDocStart2 l.Shapes.All
  Next l
  
  For c = myDoc.Pages.Count To 1 Step -1
      If c <> lc Then
      Set myPage = myDoc.Pages(c)
      myPage.Activate
      myPage.UnlockAllShapes
      For Each l In myPage.Layers
          If l.IsSpecialLayer = False Then infoDocStart2 l.Shapes.All
      Next l
      End If
  Next c
  
  myOldPage.Activate
  
  If myLevelWork Then Application.Status.EndProgress: _
  boostFinish endUndoGroup:=True: myDoc.unDo: myDoc.ClearSelection: myAfterWork
End Sub



'====================================================================================
'============================     Режим ActivePage     ==============================
'====================================================================================
Private Sub infoActivePageMode()
  Dim l As Layer
  
  myBeforeWork
  boostStart "INFO"
  myLayerScan
  Application.Status.BeginProgress "Get INFO...", False
  
  myOldPage.UnlockAllShapes
  
  For Each l In myOldPage.Layers
      If l.IsSpecialLayer = False Then myCountS = myCountS + l.Shapes.Count
  Next l
  
  For Each l In myOldPage.Layers
      If l.IsSpecialLayer = False Then infoDocStart2 l.Shapes.All
  Next l
  
  Application.Status.EndProgress
  boostFinish endUndoGroup:=True
  myDoc.unDo
  myDoc.ClearSelection
  myAfterWork
End Sub

