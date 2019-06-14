Attribute VB_Name = "ctc_functions"
Option Explicit

'====================================================================================
'==================================     До и после     ==============================
'====================================================================================
Public Function myBeforeWork()
        On Error Resume Next
        myVarNull
        myCountS = 0: mySc = 0
        End Function
Public Function myAfterWork()
        On Error Resume Next
        'effLens = effLens - shTr
        Dim mySel As Boolean
        If ActiveSelectionRange.Count = 0 Then mySel = True
        list_EffLens.RemoveRange list_EffTransparency
        If mySel Then myDoc.ClearSelection
        'InfoForm.Show 0
        End Function
'====================================================================================
'================================     Счётчик Шейпов     ============================
'====================================================================================
Public Function myFindShapesCount()
        Dim l As Layer
        For Each myPage In myDoc.Pages
        For Each l In myPage.Layers
        myCountS = myCountS + l.FindShapes.Count
        Next l
        Next myPage
        End Function
Public Function myFindShapesMasterCount()
        Dim l As Layer
        For Each l In myMasterPage.Layers
        If l.IsGuidesLayer = False And l.IsGridLayer = False Then _
        If l.Master = True Then myCountS = myCountS + l.FindShapes.Count
        Next l
        End Function
'====================================================================================
'===========================     Проверка слоёв в ИНФО     ==========================
'====================================================================================
Public Function myLayerScan()
        For Each myPage In myDoc.Pages
        For Each myLayer In myPage.Layers
            If myLayer.IsSpecialLayer = False Then
                If myLayer.Editable = False Then lEdit = lEdit + 1: myLayer.Editable = True
                If myLayer.Visible = False Then lVis = lVis + 1: myLayer.Visible = True
                If myLayer.Printable = False Then lPrint = lPrint + 1: myLayer.Printable = True
            End If
        Next myLayer
        Next myPage

        For Each myLayer In myMasterPage.Layers
            If myLayer.IsGuidesLayer = False And myLayer.IsGridLayer = False Then
            If myLayer.Master = True Then
                If myLayer.Editable = False Then lEdit = lEdit + 1: myLayer.Editable = True
                If myLayer.Visible = False Then lVis = lVis + 1: myLayer.Visible = True
                If myLayer.Printable = False Then lPrint = lPrint + 1: myLayer.Printable = True
            End If
            End If
        Next myLayer
        End Function
'====================================================================================
'==================     Обработка слоёв перед конвертированием     ==================
'====================================================================================
Public Function myLayerEnable()
        Dim s$, sd$: s = "": sd = ""
        
        'Для слоёв на страницах
        For Each myPage In myDoc.Pages
        For Each myLayer In myPage.Layers
            If myLayer.IsSpecialLayer = False Then
                s = s & myLayerEnable2(myLayer)
            End If
        Next myLayer
        Next myPage
        
        'Для мастер слоёв
        For Each myLayer In myMasterPage.Layers
            If myLayer.IsGuidesLayer = False And myLayer.IsGridLayer = False Then
            If myLayer.Master = True Then
                If myLayer.IsSpecialLayer = False Then
                   s = s & myLayerEnable2(myLayer)
                End If
            End If
            End If
            
            'Вписать работу для десктопа
            If myLayer.IsDesktopLayer Then
                sd = myLayerEnable2(myLayer)
            End If
            
        Next myLayer
        If s <> "" Then s = Left(s, Len(s) - 1)
        If sd <> "" Then sd = Left(sd, Len(sd) - 1)
        SaveSetting macroName, "Convert", "LayersVPE", s
        SaveSetting macroName, "Convert", "LayersVPEd", sd
        End Function
Private Function myLayerEnable2(myLayer As Layer) As String
        Dim s$, ss$: s = "": ss = ""
        If conv_cmEnable Then If myLayer.Editable = False Then myLayer.Editable = True
        If conv_cmVisible Then If myLayer.Visible = False Then myLayer.Visible = True
        If conv_cmPrint Then If myLayer.Printable = False Then myLayer.Printable = True
        
        If conv_cmNVisible Then
            If myLayer.Visible = False Then myLayer.Visible = True: s = myLayer.Name & ",0" Else s = myLayer.Name & ",1"
            If myLayer.Printable = False Then myLayer.Printable = True: s = s & ",0" Else s = s & ",1"
            If myLayer.Editable = False Then myLayer.Editable = True: s = s & ",0|" Else s = s & ",1|"
        End If
            
        myLayerEnable2 = ss & s: s = ""
        End Function
        
Public Function myLayerDesable()
        Dim s$, a$(), b$(), d$()
        s = GetSetting(macroName, "Convert", "LayersVPE", "")
        If s = "" Then Exit Function
        a = Split(s, "|")
        'Вписать проверку на ошибки
        Dim i&
        For i = 0 To UBound(a)
            b = Split(a(i), ",")
            For Each myPage In myDoc.Pages
                For Each myLayer In myPage.Layers
                
                    If myLayer.IsSpecialLayer = False Then
                        If myLayer.Name = b(0) Then
                            myLayer.Visible = CBool(b(1))
                            myLayer.Printable = CBool(b(2))
                            myLayer.Editable = CBool(b(3))
                        End If
                    End If
                    
                Next myLayer
            Next myPage
    
            For Each myLayer In myMasterPage.Layers
                If myLayer.IsGuidesLayer = False And myLayer.IsGridLayer = False Then
                If myLayer.Master = True Then
                    If myLayer.IsSpecialLayer = False Then
                    If myLayer.Name = b(0) Then
                        myLayer.Visible = CBool(b(1))
                        myLayer.Printable = CBool(b(2))
                        myLayer.Editable = CBool(b(3))
                    End If
                    End If
                End If
                End If
                    
                'Вписать работу для десктопа
                If myLayer.IsDesktopLayer Then
                    d = Split(GetSetting(macroName, "Convert", "LayersVPEd", ""), ",")
                    myLayer.Visible = CBool(d(1))
                    myLayer.Printable = CBool(d(2))
                    myLayer.Editable = CBool(d(3))
                End If
                
            Next myLayer
        Next
        End Function
