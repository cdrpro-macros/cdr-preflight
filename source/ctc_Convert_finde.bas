Attribute VB_Name = "ctc_Convert_finde"
Option Explicit

Private myFindeShRange As ShapeRange, myShType As cdrShapeType


'====================================================================================
'========================     Find Shapes for Converter     =========================
'====================================================================================
Public Sub myFindeShapes(myType$)
        Dim l As Layer, sr As ShapeRange, lc&, c&
        Set sr = New ShapeRange
        If myResumeErr Then On Error Resume Next
        'По мастеру =================================
        myMasterPage.UnlockAllShapes
        For Each l In myMasterPage.Layers
            If l.IsGuidesLayer = False And l.IsGridLayer = False Then
                If l.Master = True Then
                    If l.Editable And l.Visible And l.Printable Then sr.AddRange l.Shapes.All
                End If
            End If
        Next l
        'По страницам ===============================
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        lc = myOldPage.index
        For Each l In myOldPage.Layers
            If l.IsSpecialLayer = False Then _
                If l.Editable And l.Visible And l.Printable Then sr.AddRange l.Shapes.All
        Next l
        For c = myDoc.Pages.Count To 1 Step -1
            If c <> lc Then
            Set myPage = myDoc.Pages(c)
            myPage.Activate: myPage.UnlockAllShapes
            For Each l In myPage.Layers
                If l.IsSpecialLayer = False Then _
                    If l.Editable And l.Visible And l.Printable Then sr.AddRange l.Shapes.All
            Next l
            End If
        Next c
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        For Each myPage In myDoc.Pages
'            myPage.Activate: myPage.UnlockAllShapes
'            For Each l In myPage.Layers
'                If l.IsSpecialLayer = False Then sr.AddRange l.Shapes.All
'            Next l
'        Next myPage
        myFindeShapesEd sr, myType
        End Sub
Public Sub myFindeShapesEd(sr As ShapeRange, myType$)
        Set myFindeShRange = New ShapeRange
        
        Select Case myType
            Case "OLE_EPS_Symb"
                myFindeShRange.RemoveAll
                myShType = cdrOLEObjectShape: myFindeShapes2 sr
                Set list_OLE = myFindeShRange.ReverseRange
                
                myFindeShRange.RemoveAll
                myShType = cdrEPSShape: myFindeShapes2 sr
                Set list_EPS = myFindeShRange.ReverseRange
                
                myFindeShRange.RemoveAll
                myShType = cdrSymbolShape: myFindeShapes2 sr
                Set list_symbol = myFindeShRange.ReverseRange
                
                myFindeShRange.RemoveAll
                myShType = cdrLinearDimensionShape: myFindeShapes2 sr
                Set list_Dimension = myFindeShRange.ReverseRange
                
            Case "PowerClip"
                myFindeShRange.RemoveAll
                myFindeShapesPow sr
                Set list_PoweClip = myFindeShRange.ReverseRange
                
            Case "Effect"
                myFindeEff sr
                
                myFindeShRange.RemoveAll
                myShType = cdrArtisticMediaGroupShape: myFindeShapes2 sr
                Set list_EffArtisticMedia = myFindeShRange.ReverseRange

                myFindeShRange.RemoveAll
                myShType = cdrCustomEffectGroupShape: myFindeShapes2 sr
                Set list_EffBevel = myFindeShRange.ReverseRange
                
            Case "Txt_Bit"
                myFindeShRange.RemoveAll
                myShType = cdrTextShape: myFindeShapes2 sr
                Set list_Text = myFindeShRange.ReverseRange
                
                myFindeShRange.RemoveAll
                myShType = cdrBitmapShape: myFindeShapes2 sr
                Set list_Allbit = myFindeShRange.ReverseRange
                
            Case "Fill_Outline"
                myFindeFillOutline sr
                
                myFindeShRange.RemoveAll
                myShType = cdrMeshFillShape: myFindeShapes2 sr
                Set list_fillMesh = myFindeShRange.ReverseRange
        End Select
        End Sub
        
        
        
        
        
'====================================================================================
'===========================     Find Shapes По s.Type    ===========================
'====================================================================================
Private Sub myFindeShapes2(sr As ShapeRange)
        Dim s As Shape
        If myResumeErr Then On Error Resume Next
        On Error GoTo myErr
        For Each s In sr
            If s.Type = cdrGroupShape Then
                myFindeShapes2 s.Shapes.All
            Else
                If s.Type = myShType Then myFindeShRange.Add s
            End If
            If Not s.PowerClip Is Nothing Then myFindeShapes2 s.PowerClip.Shapes.All
        Next s
        Exit Sub
myErr:
        If Err.Number = -2147467259 Then
            Err.Clear
            Resume Next
        End If
        End Sub
'====================================================================================
'============================     Find Fill & Outline    ============================
'====================================================================================
Private Sub myFindeFillOutline(sr As ShapeRange)
        Dim s As Shape
        If myResumeErr Then On Error Resume Next
        On Error GoTo myErr
        For Each s In sr
            If s.Type = cdrGroupShape Then
                myFindeFillOutline s.Shapes.All
            Else
                If s.CanHaveFill Then If s.Fill.Type <> cdrNoFill Then list_CanFill.Add s
                If s.CanHaveOutline Then If s.Outline.Type <> cdrNoOutline Then list_CanOutline.Add s
                If s.CanHaveOutline Then If s.Outline.Type = cdrNoOutline Then list_OutlineProbl.Add s
            End If
            If Not s.PowerClip Is Nothing Then myFindeFillOutline s.PowerClip.Shapes.All
        Next s
        Exit Sub
myErr:
        If Err.Number = -2147467259 Then
            Err.Clear
            Resume Next
        End If
        End Sub
'====================================================================================
'===============================     Find PowerClip    ==============================
'====================================================================================
Private Sub myFindeShapesPow(sr As ShapeRange)
        Dim s As Shape
        If myResumeErr Then On Error Resume Next
        For Each s In sr
            If s.Type = cdrGroupShape Then myFindeShapesPow s.Shapes.All
            If Not s.PowerClip Is Nothing Then myFindeShRange.Add s
        Next s
        End Sub
'====================================================================================
'============================     Find Effect Objects    ============================
'====================================================================================
Private Sub myFindeEff(sr As ShapeRange)
        Dim s As Shape, ef As Effect
        If myResumeErr Then On Error Resume Next
        On Error GoTo myErr
        For Each s In sr
            If s.Type = cdrGroupShape Then
                myFindeEff s.Shapes.All
            Else
                If s.Effects.Count > 0 Then
                For Each ef In s.Effects
                    Select Case ef.Type
                    Case cdrDropShadow: list_EffShadow.Add s
                    Case cdrDistortion: list_EffDistortion.Add s
                    Case cdrEnvelope: list_EffEnvelope.Add s
                    Case cdrBlend: list_EffBlend.Add ef.Blend.StartShape
                    Case cdrContour: list_EffContour.Add s
                    Case cdrLens: list_EffLens.Add s
                    End Select
                Next ef
                End If
            End If
            
            If Not s.PowerClip Is Nothing Then myFindeEff s.PowerClip.Shapes.All
        Next s
        Exit Sub
myErr:
        If Err.Number = -2147467259 Then
            Err.Clear
            Resume Next
        End If
        End Sub

