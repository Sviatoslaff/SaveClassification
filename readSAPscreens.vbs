Sub ProcessArticle(article, session)
	On Error Resume Next
    
    session.findById("wnd[0]").maximize
    
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = article
    session.findById("wnd[0]/tbar[1]/btn[5]").press

    'Choose Basic Data 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = True

	If Err.Number <> 0 Then										'Если вкладкм нет, то ошибка
		Exit Sub												'Выходим из процелуры 
	End If		

    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB7:SAPLMGD1:2033/btnPUSH_GRUNDDATENTEXT").press
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/tblSAPLMGD1TC_LONGTEXT/btnSELE[0,0]").setFocus
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/tblSAPLMGD1TC_LONGTEXT/btnSELE[0,0]").press
    BD1_Text = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell").Text
    Art_Name = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").Text
    
    pressF3()
    pressF3()
    
    'Choose Classification
    session.findById("wnd[0]/tbar[1]/btn[5]").press                                                 'кнопка Select Views
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = False
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = False
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).selected = False
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = True
    
    '<GuiTextField Id="/app/con[1]/ses[0]/wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]" Name="MSICHTAUSW-DYTXT" Text="Classification"/>
    If session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]", False).Text = "Classification" Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press                                             'OK
        
        'session.findById("wnd[0]/usr/btn%#AUTOTEXT004").press                                      'кнопка выбора класса
        
        'поиск нужного класса в таблице классов
        Set elem = Nothing
        winCount = session.findById("wnd[1]/usr").Children.Count
        Dim myArr(100, 1)
        Dim maxX
        Dim arrElement(200,10)
        For i = 1 To winCount
            winElement = session.findById("wnd[1]/usr").Children(i - 1).Text
            If winElement = "Z01" Then
                Set elem = session.findById("wnd[1]/usr").Children(i - 1)
                Exit For
            End If
        Next
        If Not elem Is Nothing Then
            elem.setFocus
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/btnRCTMS-LISTE").press
            'обработка окна со списком классификации, получение массива с индексами значений
            iCount = session.findById("wnd[1]/usr").Children.Count
            y = 0
            For i = 1 To iCount
                idElem = session.findById("wnd[1]/usr").Children(i - 1).Id
                txtElem = session.findById("wnd[1]/usr").Children(i - 1).Text
                If txtElem <> "" Then
                    lenElem = Len(idElem)
                    leftElem = InStrRev(idElem, "[")
                    IndexElem = Mid(idElem, leftElem + 1)
                    IndexElem = Left(IndexElem, Len(IndexElem) - 1)
                    arrElem = Split(IndexElem, ",")
                    x = CInt(arrElem(1))
                    If x = maxX Then
                        y = y + 1
                    Else
                        y = 0
                    End If
                    arrElement(x,y) = txtElem
                    maxX = x
                End If
            Next
			session.findById("wnd[1]/tbar[0]/btn[12]").press
            'составление массива с парой параметр - значение
            For i = 6 To maxX
                myArr(i - 6, 0) = arrElement(i,0)
                txtElem = ""
                For j = 1 To 10
                    If txtElem = "" Then
                        txtElem = arrElement(i,j)
                    Else
                        txtElem = txtElem & " " & arrElement(i,j)
                    End If
                Next
                myArr(i - 6, 1) = txtElem
            Next
            'запись значений в БД
            ClassCode = "Z01"
            ClassName = "Classification"            
        Else
            ClassCode = "NtF"
            ClassName = "Class Z01 Not Found"
			session.findById("wnd[1]/tbar[0]/btn[12]").press
        End If

    Else
        ClassCode = "NoC"
        ClassName = "The Article has no Classification"
		session.findById("wnd[1]/tbar[0]/btn[12]").press					'выход из окна с ракурсами
    End If
	MaxX = MaxX - 6
    Call InsertArticle(article, Art_Name, BD1_Text, ClassCode, ClassName, myArr, MaxX )
    pressF3()
End Sub

'To avoid using error handling you can use:
'If Not session.findById("wnd[1]", False) Is Nothing Then
'    session.findById("wnd[1]").setFocus
'End If