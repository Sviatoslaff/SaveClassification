Sub InsertArticle(article, Art_Name, BD1_Text, ClassCode, ClassName, myArr, MaxX)
    
    Dim Conn, rs, S
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open "PROVIDER=Microsoft.ACE.OLEDB.12.0;" & _
    "DATA SOURCE=C:\VBScript\SaveClassification\MATERIALS.accdb"
    Set rs = CreateObject("ADODB.Recordset")
    
    Set rs = Conn.Execute(" INSERT INTO Materials " _
     & "VALUES ('" & article _
     & "', '" & Art_Name _
     & "', '" & BD1_Text _
     & "', '" & ClassCode _
     & "', '" & ClassName _
     & "')")

    For i = 0 To MaxX
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
End Sub