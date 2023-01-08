Sub InsertArticle (article, Art_Name, BD1_Text, ClassCode, ClassName, myArr, MaxX)

    Dim Conn, rs, S
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
    "Dbq=C:\Projects\SULZER03\MATERIALS.accdb"
    Set rs = CreateObject("ADODB.Recordset")


    Set rs = Conn.Execute(" INSERT INTO Materials " _
     & "VALUES ('" & article _
     & "', '" & Art_Name _
     & "', '" & BD1_Text _
     & "', '" & ClassCode _
     & "', '" & ClassName _
     & "', '"  _
     & "')")

    If rs.ActiveConnection.Errors.Count >= 0 Then  
     For Each Err In rs.ActiveConnection.Errors  
        MsgBox "Error number: " & Err.Number & vbCr & _  
           Err.Description  
     Next
    End If 

    For i = 0 To MaxX-1
        'if myArr(i,0) = "" Then Exit For
        Set rs = Conn.Execute(" INSERT INTO Chars " _
        & "VALUES ('" & article _
        & "', '" & myArr(i, 0) _
        & "', '" & myArr(i, 1) _
        & "')")
    Next    

    If rs.State = adStateOpen Then
        rs.Close
    End If
    If Conn.State = adStateOpen Then
        Conn.Close
    End If
Exit Sub

End Sub



Sub SaveArticle (article, Art_Name, BD1_Text, ClassCode, ClassName, myArr, MaxX)

    'Set dbs = OpenDatabase("C:\Projects\SULZER031\MATERIALS.accdb")
    'Set rs = dbs.OpenRecordset("Materials")

    Dim AccApp
    Set AccApp = CreateObject("Access.Application")
    AccApp.OpenCurrentDatabase "C:\Projects\SULZER031\MATERIALS.accdb"
    AccApp.DoCMD.RunSQL("INSERT INTO Materials " _
    & "VALUES ('" & article _
    & "', '" & Art_Name _
    & "', '" & BD1_Text _
    & "', '" & ClassCode _
    & "', '" & ClassName _
    & "', '1"  _
    & "')")

    For i = 0 To MaxX-1
        'if myArr(i,0) = "" Then Exit For
        AccApp.DoCMD.RunSQL("INSERT INTO Chars " _
        & "VALUES ('" & article _
        & "', '" & myArr(i, 0) _
        & "', '" & myArr(i, 1) _
        & "')")
    Next    

Exit Sub



End Sub
