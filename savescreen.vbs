sPath = "C:\VBScript"
screen1 = "screen1.xml"
screen2 = "screen2.xml"

startTransaction("MM03")

session.findById("wnd[0]").maximize

session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "M098-741100" '"10417"
session.findById("wnd[0]/tbar[1]/btn[5]").press
'Basic Data 1
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB7:SAPLMGD1:2033/btnPUSH_GRUNDDATENTEXT").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/tblSAPLMGD1TC_LONGTEXT/btnSELE[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/tblSAPLMGD1TC_LONGTEXT/btnSELE[0,0]").press
BD1_Text = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell").Text
Art_Name = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").Text

pressF3()
pressF3()

'Classification
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true

'<GuiTextField Id="/app/con[1]/ses[0]/wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]" Name="MSICHTAUSW-DYTXT" Text="Classification"/>
If session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]", False).Text = "Classification" Then
	session.findById("wnd[1]/tbar[0]/btn[0]").press											'OK

	session.findById("wnd[0]/usr/btn%#AUTOTEXT004").press									'кнопка выбора класса
	
	'поиск нужного класса в таблице классов
	Set elem = Nothing
	winCount = session.findById("wnd[1]/usr").Children.Count
	Dim myArr(100, 1) 
	Dim maxX
	Dim arrElement(200,10)
	For i = 1 To winCount
		winElement = session.findById("wnd[1]/usr").Children(i-1).Text
		If winElement = "001" Then
			Set elem = session.findById("wnd[1]/usr").Children(i-1)
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
			idElem = session.findById("wnd[1]/usr").Children(i-1).Id
			txtElem = session.findById("wnd[1]/usr").Children(i-1).Text
			If txtElem <> "" Then
				lenElem = len(idElem)
				leftElem = InStrRev(idElem, "[")
				IndexElem = Mid(idElem, leftElem + 1)
				IndexElem = Left(IndexElem, Len(IndexElem) - 1)
				arrElem = Split(IndexElem, ",")
				x = CInt(arrElem(1))
				if x = maxX Then
					y = y + 1
				else 
					y = 0
				end if
				arrElement(x,y) = txtElem
				maxX = x
			End If			
		Next
		'составление массива с парой параметр - значение
		For i = 6 To maxX
			myArr(i-6, 0) = arrElement(i,0)
			txtElem = ""
			For j = 1 To 10
				If txtElem = "" Then
					txtElem = arrElement(i,j)
				Else	
					txtElem = txtElem & " " &  arrElement(i,j)
				End If	
			Next
			myArr(i-6, 1) = txtElem
		Next
		'запись значений в БД

	Else
		MsgBox "Class not found"
		' СДЕЛАТЬ ОБРАБОТКУ - КЛАСС НЕ НАЙДЕН
	End If

Else 
	MsgBox "No classification"
	' СДЕЛАТЬ ОБРАБОТКУ - КЛАССИФИКАЦИИ НЕТ
End If

filepath = sPath & "\" & screen1


	Dim currentNode

	Set xmlParser = CreateObject("Msxml2.DOMDocument")

	' Создание объявления XML
	xmlParser.appendChild(xmlParser.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'"))

	If Not IsObject(application) Then
	Set SapGuiAuto  = GetObject("SAPGUI")
	Set application = SapGuiAuto.GetScriptingEngine
	End If
	If Not IsObject(connection) Then
	Set connection = application.Children(0)
	End If
	If Not IsObject(session) Then
	Set session    = connection.Children(0)
	End If
	If IsObject(WScript) Then
	WScript.ConnectObject session,     "on"
	WScript.ConnectObject application, "on"
	End If

	' Максимизируем окно SAP
'	session.findById("wnd[0]").maximize

	'enumeration "wnd[0]"
	enumeration "wnd[1]/usr"

	MsgBox "Finished!", vbSystemModal Or vbInformation



Sub enumeration(SAPRootElementId)

	Set SAPRootElement = session.findById(SAPRootElementId)
	
	'Создание корневого элемента
	Set XMLRootNode = xmlParser.appendChild(xmlParser.createElement(SAPRootElement.Type))
	
	enumChildrens SAPRootElement, XMLRootNode
	
	'xmlParser.save("C:\VBScript\SAP_tree.xml")
	xmlParser.save("C:\VBScript\SAP_tree-001.xml")
End Sub

Sub enumChildrens(SAPRootElement, XMLRootNode) 
	For i = 0 To SAPRootElement.Children.Count - 1
		Set SAPChildElement = SAPRootElement.Children.ElementAt(i)
		
		' Создаем узел
		Set XMLSubNode = XMLRootNode.appendChild(xmlParser.createElement(SAPChildElement.Type))
		
		' Атрибут Name
		Set attrName = xmlParser.createAttribute("Name")
		attrName.Value = SAPChildElement.Name
		XMLSubNode.setAttributeNode(attrName)
		
		' Атрибут Text
		If (Len(SAPChildElement.Text) > 0) Then
			Set attrText = xmlParser.createAttribute("Text")
			attrText.Value = SAPChildElement.Text
			XMLSubNode.setAttributeNode(attrText)
		End If
		
		' Атрибут Id
		Set attrId = xmlParser.createAttribute("Id")
		attrId.Value = SAPChildElement.Id
		XMLSubNode.setAttributeNode(attrId)
		
		' Если текущий объект - контейнер, то перебираем дочерние элементы
		If (SAPChildElement.ContainerType) Then enumChildrens SAPChildElement, XMLSubNode
	Next
End Sub

pressF3()

MsgBox("The script finished.")

'To avoid using error handling you can use:
'If Not session.findById("wnd[1]", False) Is Nothing Then
'    session.findById("wnd[1]").setFocus
'End If