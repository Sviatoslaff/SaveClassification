
Sub saveScreenXML(filepath)

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
	enumeration "wnd[0]/usr"

	MsgBox "Finished!", vbSystemModal Or vbInformation

End Sub

Sub enumeration(SAPRootElementId)

	Set SAPRootElement = session.findById(SAPRootElementId)
	
	'Создание корневого элемента
	Set XMLRootNode = xmlParser.appendChild(xmlParser.createElement(SAPRootElement.Type))
	
	enumChildrens SAPRootElement, XMLRootNode
	
	'xmlParser.save("C:\VBScript\SAP_tree.xml")
	xmlParser.save(filepath)
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