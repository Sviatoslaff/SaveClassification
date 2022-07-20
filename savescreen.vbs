sPath = "C:\VBScript"
screen1 = "screen1.xml"
screen2 = "screen2.xml"

startTransaction("MM03")

session.findById("wnd[0]").maximize

session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "M098-741100" '"10417"
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true

'<GuiTextField Id="/app/con[1]/ses[0]/wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]" Name="MSICHTAUSW-DYTXT" Text="Classification"/>
If session.findById("/app/con[1]/ses[0]/wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]", False).Text = "Classification" Then
	session.findById("wnd[1]/tbar[0]/btn[0]").press											'OK

	session.findById("/app/con[1]/ses[0]/wnd[0]/usr/btn%#AUTOTEXT004").press				'кнопка выбора класса
	
	'поиск нужного класса в таблице классов
	Set elem = Nothing
	winCount = session.findById("wnd[1]/usr").Children.Count
	For i=1 To winCount
		winElement = session.findById("wnd[1]/usr").Children(i-1).Text
		If winElement = "Z01" Then
			Set elem = session.findById("wnd[1]/usr").Children(i-1)
			Exit For
		End If
	Next
	If Not elem Is Nothing Then
		elem.setFocus
		session.findById("wnd[1]/tbar[0]/btn[0]").press
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
	session.findById("wnd[0]").maximize

	'enumeration "wnd[0]"
	enumeration "wnd[1]/usr"

	MsgBox "Finished!", vbSystemModal Or vbInformation



Sub enumeration(SAPRootElementId)

	Set SAPRootElement = session.findById(SAPRootElementId)
	
	'Создание корневого элемента
	Set XMLRootNode = xmlParser.appendChild(xmlParser.createElement(SAPRootElement.Type))
	
	enumChildrens SAPRootElement, XMLRootNode
	
	'xmlParser.save("C:\VBScript\SAP_tree.xml")
	xmlParser.save("C:\VBScript\SAP_tree-00.xml")
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