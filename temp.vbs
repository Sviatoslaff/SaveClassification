Sub InsertMaterial() 
     
    Dim Conn, rs, S
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open "PROVIDER=Microsoft.ACE.OLEDB.12.0;" & _
                "DATA SOURCE=C:\Projects\SULZER02\MATERIALS.accdb"
    Set rs = CreateObject("ADODB.Recordset")

    Set rs = Conn.Execute(" INSERT INTO Materials " _ 
            & "VALUES ('V12-ADDED', 'V12-NAME', 'V-12 BASIC TEXT', '112', '112 CLASS')")
    If rs.State = adStateOpen then 
        rs.Close
    End If
    If Conn.State = adStateOpen then
        Conn.Close 
    End If    
End Sub

InsertMaterial()