Sub InsertArticle(article, Art_Name, BD1_Text, ClassCode, ClassName, myArr, MaxX)
    
    Dim Conn, rs, S
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open "PROVIDER=Microsoft.ACE.OLEDB.12.0;" & _
    "DATA SOURCE=C:\VBScript\SaveClassification\MATERIALS.accdb"
    Set rs = CreateObject("ADODB.Recordset")

    On Error GoTo Err_Execute  

    Set rs = Conn.Execute(" INSERT INTO Materials " _
     & "VALUES ('" & article _
     & "', '" & Art_Name _
     & "', '" & BD1_Text _
     & "', '" & ClassCode _
     & "', '" & ClassName _
     & "', '"  _
     & "')")

    For i = 0 To MaxX-1
        'if myArr(i,0) = "" Then Exit For
        Set rs = Conn.Execute(" INSERT INTO Chars " _
        & "VALUES ('" & article _
        & "', '" & myArr(i, 0) _
        & "', '" & myArr(i, 1) _
        & "')")
    Next    

    Err_Execute:  
    ' Notify user of any errors that result from executing the query  
    If rs.ActiveConnection.Errors.Count >= 0 Then  
       For Each Err In rs.ActiveConnection.Errors  
          MsgBox "Error number: " & Err.Number & vbCr & _  
             Err.Description  
       Next Err  
    End If 

    If rs.State = adStateOpen Then
        rs.Close
    End If
    If Conn.State = adStateOpen Then
        Conn.Close
    End If


End Sub
