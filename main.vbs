'Запрашиваем файл с номерами
'Set ArticlesExcel = selectExcel()

excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open (excelFile)

' Считаем, что в первой строке - заголовок
intRow = 2
' Цикл для каждого артикула
startTransaction("MM03")
Do Until ArticlesExcel.Cells(intRow,1).Value = ""
	article = ArticlesExcel.Cells(intRow, 1).Value
	Call ProcessArticle(article, session)								'вызов процедуры для чтения из SAP очередного материала
	intRow = intRow + 1
Loop

objWorkbook.Close False
ArticlesExcel.Quit
pressF3()

MsgBox "Script finished!", vbSystemModal Or vbInformation