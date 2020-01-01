Attribute VB_Name = "ctc_var"
Option Explicit

'#If VBA7 Then
  Public Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
'#Else
'  Public Declare Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
'#End If

Public Const myResumeErr As Boolean = False

Public Const macroName$ = "CdrPreflight"
Public Const macroVersion$ = "6.3.2"
Public Const macroModifyDate$ = "(01.01.2020)"
Public Const macroCopyright$ = "2006-2020 by Sanich"
Public Const sREGAPPOPT$ = "General"

Public Const convPresVer& = 3 'Версия пресета
Public Const ctcPresVer& = 1 'Версия пресета для проверки

Public Const myWebSite$ = "cdrpro.ru"
Public Const myEmail$ = "info@cdrpro.ru"

Public Const IGNORE_ITEM = "==="

Public cdrMyType$ 'Типы для списка
Public Const myDefPreset$ = "1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|"


Public _
  myDoc As Document, _
  myUnit As cdrUnit, _
  myUnitWork As cdrUnit, _
  myPage As Page, _
  myOldPage As Page, _
  myMasterPage As Page, _
  myLayer As Layer, _
  myCountS&, mySc&


'======Для Листа======================
Public srObj As New ShapeRange
Public myFindProblemType$

    
Public _
    list_OLE As ShapeRange, _
    list_EPS As ShapeRange, _
    list_symbol As ShapeRange, _
    list_Dimension As ShapeRange, _
    list_PerfSh As ShapeRange, _
    list_PoweClip As ShapeRange, _
    list_noCloseCur As ShapeRange, _
    list_NodesMax As ShapeRange
    
Public _
    list_EffTransparency As ShapeRange, _
    list_EffBlend As ShapeRange, _
    list_EffShadow As ShapeRange, _
    list_EffLens As ShapeRange, _
    list_EffContour As ShapeRange, _
    list_ControlPath As ShapeRange, _
    list_EffBevel As ShapeRange, _
    list_EffArtisticMedia As ShapeRange, _
    list_EffExtrude As ShapeRange, _
    list_EffEnvelope As ShapeRange, _
    list_EffDistortion As ShapeRange, _
    list_EffPerspective As ShapeRange

Public _
    list_Text As ShapeRange, _
    list_TextOver As ShapeRange, _
    list_Table As ShapeRange, _
    list_txtSmalPt As ShapeRange, _
    list_txtSmalCol As ShapeRange
    
Public _
    list_FountainFill As ShapeRange, _
    list_fillPattern As ShapeRange, _
    list_fillPS As ShapeRange, _
    list_fillTexture As ShapeRange, _
    list_fillMesh As ShapeRange, _
    list_OutlineMin As ShapeRange, _
    list_OutlineProbl As ShapeRange, _
    list_OutLineScal As ShapeRange

Public _
    list_CanFillOutline As ShapeRange, _
    list_CanFill As ShapeRange, _
    list_CanOutline As ShapeRange, _
    list_ColorSmalLim As ShapeRange

Public _
    list_Allbit As ShapeRange, _
    list_MinDPI As ShapeRange, _
    list_MaxDPI As ShapeRange, _
    list_BitRot As ShapeRange, _
    list_BitXY As ShapeRange, _
    list_BitCrop As ShapeRange, _
    list_BitLink As ShapeRange, _
    list_BitTr As ShapeRange, _
    list_BitOverpr As ShapeRange
    
Public _
    list_BitCMYK As ShapeRange, _
    list_BitRGB As ShapeRange, _
    list_BitBW As ShapeRange, _
    list_BitDuo As ShapeRange, _
    list_BitGr As ShapeRange, _
    list_BitLAB As ShapeRange, _
    list_BitDevN As ShapeRange, _
    list_BitPal As ShapeRange
'=====================================




Public _
    errCount&
Public _
    lEdit&, _
    lVis&, _
    lPrint&, _
    tcOnP&
Public _
    b16&, _
    bCMYKm&, _
    bitFill&, _
    bitOutl&
Public _
    sBarCode&
Public _
    shhf&, _
    shnf&, _
    shuf&, _
    sFonFillMP&, _
    uColorTIL&, _
    uColorTIL300&, _
    sColorSmalLim&
Public _
    scolBW&, _
    scolCMY&, _
    scolCMYK&, _
    scolCMYK100&, _
    scolGray&, _
    scolHLS&, _
    scolHSB&, _
    scolLab&, _
    scolMix&, _
    scolMulti&, _
    scolPan&, _
    scolPanH&, _
    scolReg&, _
    scolRGB&, _
    scolSpot&, _
    scolUserInk&, _
    scolYIQ&
Public _
    sOuLineN&, _
    sOuLineEnh&, _
    oPrinf&, _
    oPrino&
Public _
    sCurNod&






'==============================================================================================
'==============================================================================================

Public Function myVarNull()
    errCount = 0
    errStr = ""
    
    Set list_OLE = New ShapeRange
    Set list_EPS = New ShapeRange
    Set list_symbol = New ShapeRange
    Set list_Dimension = New ShapeRange
    Set list_PerfSh = New ShapeRange
    Set list_PoweClip = New ShapeRange
    Set list_noCloseCur = New ShapeRange
    Set list_NodesMax = New ShapeRange
    
    Set list_EffTransparency = New ShapeRange
    Set list_EffBlend = New ShapeRange
    Set list_EffShadow = New ShapeRange
    Set list_EffLens = New ShapeRange
    Set list_EffContour = New ShapeRange
    Set list_ControlPath = New ShapeRange
    Set list_EffBevel = New ShapeRange
    Set list_EffArtisticMedia = New ShapeRange
    Set list_EffExtrude = New ShapeRange
    Set list_EffEnvelope = New ShapeRange
    Set list_EffDistortion = New ShapeRange
    Set list_EffPerspective = New ShapeRange
    
    Set list_Text = New ShapeRange
    Set list_TextOver = New ShapeRange
    Set list_Table = New ShapeRange
    Set list_txtSmalPt = New ShapeRange
    Set list_txtSmalCol = New ShapeRange
    
    Set list_FountainFill = New ShapeRange
    Set list_fillPattern = New ShapeRange
    Set list_fillPS = New ShapeRange
    Set list_fillTexture = New ShapeRange
    Set list_fillMesh = New ShapeRange
    Set list_OutlineMin = New ShapeRange
    Set list_OutLineScal = New ShapeRange
    Set list_OutlineProbl = New ShapeRange
    
    Set list_CanFillOutline = New ShapeRange
    Set list_CanFill = New ShapeRange
    Set list_CanOutline = New ShapeRange
    Set list_ColorSmalLim = New ShapeRange
    
    Set list_Allbit = New ShapeRange
    Set list_MinDPI = New ShapeRange
    Set list_MaxDPI = New ShapeRange
    Set list_BitRot = New ShapeRange
    Set list_BitXY = New ShapeRange
    Set list_BitCrop = New ShapeRange
    Set list_BitLink = New ShapeRange
    Set list_BitTr = New ShapeRange
    Set list_BitOverpr = New ShapeRange
    
    Set list_BitCMYK = New ShapeRange
    Set list_BitRGB = New ShapeRange
    Set list_BitBW = New ShapeRange
    Set list_BitDuo = New ShapeRange
    Set list_BitGr = New ShapeRange
    Set list_BitLAB = New ShapeRange
    Set list_BitDevN = New ShapeRange
    Set list_BitPal = New ShapeRange
    
    
    list_CanFillOutline.RemoveAll
    list_CanFill.RemoveAll
    list_CanOutline.RemoveAll
    list_ColorSmalLim.RemoveAll
    
    list_Allbit.RemoveAll
    list_MinDPI.RemoveAll
    list_MaxDPI.RemoveAll
    list_BitCrop.RemoveAll
    list_BitLink.RemoveAll
    list_BitXY.RemoveAll
    list_BitRot.RemoveAll
    list_BitTr.RemoveAll
    list_BitOverpr.RemoveAll
    list_BitCMYK.RemoveAll
    list_BitRGB.RemoveAll
    list_BitBW.RemoveAll
    list_BitDuo.RemoveAll
    list_BitGr.RemoveAll
    list_BitLAB.RemoveAll
    list_BitDevN.RemoveAll
    list_BitPal.RemoveAll
    
    list_Text.RemoveAll
    list_Table.RemoveAll
    list_txtSmalPt.RemoveAll
    list_txtSmalCol.RemoveAll
    
    list_EPS.RemoveAll
    list_PerfSh.RemoveAll
    list_NodesMax.RemoveAll
    list_PoweClip.RemoveAll
    
    list_FountainFill.RemoveAll
    list_fillMesh.RemoveAll
    list_fillTexture.RemoveAll
    list_fillPattern.RemoveAll
    list_fillPS.RemoveAll
    list_OutlineMin.RemoveAll
    list_OutLineScal.RemoveAll
    list_OutlineProbl.RemoveAll
    list_OLE.RemoveAll
    list_noCloseCur.RemoveAll
    list_TextOver.RemoveAll
    list_symbol.RemoveAll
    list_Dimension.RemoveAll
    
    list_EffBlend.RemoveAll
    list_EffShadow.RemoveAll
    list_EffLens.RemoveAll
    list_EffTransparency.RemoveAll
    list_EffContour.RemoveAll
    list_ControlPath.RemoveAll
    list_EffBevel.RemoveAll
    list_EffArtisticMedia.RemoveAll
    list_EffPerspective.RemoveAll
    list_EffExtrude.RemoveAll
    list_EffEnvelope.RemoveAll
    list_EffDistortion.RemoveAll
    

    
    b16 = 0: bCMYKm = 0: bitFill = 0: bitOutl = 0
    
    tcOnP = 0
    
    sBarCode = 0
    shhf = 0: shnf = 0: shuf = 0
    
    uColorTIL = 0: uColorTIL300 = 0
    scolBW = 0: scolCMY = 0: scolCMYK = 0: scolCMYK100 = 0: scolGray = 0: scolHLS = 0: scolHSB = 0
    scolLab = 0: scolMix = 0: scolMulti = 0: scolPan = 0: scolPanH = 0
    scolReg = 0: scolRGB = 0: scolSpot = 0: scolUserInk = 0: scolYIQ = 0
    
    sColorSmalLim = 0
    sFonFillMP = 0
    sOuLineN = 0: sOuLineEnh = 0
    oPrinf = 0: oPrino = 0
    
    lEdit = 0: lVis = 0: lPrint = 0
    
    sCurNod = 0
    
    Select Case GetSetting(macroName, sREGAPPOPT, "cb_Unit", "millimeters")
    Case "millimeters": myUnitWork = cdrMillimeter
    Case "points": myUnitWork = cdrPoint
    Case Else: MsgBox "No Unit   ", vbCritical, "Warning"
    End Select
    End Function



Public Function myVarNothing()
    Set srObj = Nothing
    
    Set list_OLE = Nothing
    Set list_EPS = Nothing
    Set list_symbol = Nothing
    Set list_PerfSh = Nothing
    Set list_Dimension = Nothing
    Set list_PoweClip = Nothing
    Set list_noCloseCur = Nothing
    Set list_NodesMax = Nothing
    
    Set list_EffTransparency = Nothing
    Set list_EffBlend = Nothing
    Set list_EffShadow = Nothing
    Set list_EffLens = Nothing
    Set list_EffContour = Nothing
    Set list_ControlPath = Nothing
    Set list_EffBevel = Nothing
    Set list_EffArtisticMedia = Nothing
    Set list_EffExtrude = Nothing
    Set list_EffEnvelope = Nothing
    Set list_EffDistortion = Nothing
    Set list_EffPerspective = Nothing
    
    Set list_Text = Nothing
    Set list_TextOver = Nothing
    Set list_Table = Nothing
    Set list_txtSmalPt = Nothing
    Set list_txtSmalCol = Nothing
    
    Set list_FountainFill = Nothing
    Set list_fillPattern = Nothing
    Set list_fillPS = Nothing
    Set list_fillTexture = Nothing
    Set list_fillMesh = Nothing
    Set list_OutlineMin = Nothing
    Set list_OutLineScal = Nothing
    Set list_OutlineProbl = Nothing
    
    Set list_CanFillOutline = Nothing
    Set list_CanFill = Nothing
    Set list_CanOutline = Nothing
    Set list_ColorSmalLim = Nothing
    
    Set list_Allbit = Nothing
    Set list_MinDPI = Nothing
    Set list_MaxDPI = Nothing
    Set list_BitRot = Nothing
    Set list_BitXY = Nothing
    Set list_BitCrop = Nothing
    Set list_BitLink = Nothing
    Set list_BitTr = Nothing
    Set list_BitOverpr = Nothing
    
    Set list_BitCMYK = Nothing
    Set list_BitRGB = Nothing
    Set list_BitBW = Nothing
    Set list_BitDuo = Nothing
    Set list_BitGr = Nothing
    Set list_BitLAB = Nothing
    Set list_BitDevN = Nothing
    Set list_BitPal = Nothing
    End Function
