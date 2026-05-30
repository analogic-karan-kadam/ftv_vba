Private Sub showRecord_Released()
    On Error GoTo ErrorHandler
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim r As Long
    Dim c As Long

    Set conn = New ADODB.Connection

    conn.Open _
    "Provider=MSOLEDBSQL;" & _
    "Data Source=AAPL-L170\SQLEXPRESS;" & _
    "Initial Catalog=VbaData;" & _
    "User ID=sa;" & _
    "Password=sa;" & _
    "TrustServerCertificate=yes;"

    sql = "SELECT * FROM Student"
    Set rs = conn.Execute(sql)

    LogDiagnosticsMessage "Query executed successfully"

    If rs.EOF Then
        LogDiagnosticsMessage "No records found"
        MsgBox "No records found."
        GoTo CleanUp
    End If

    dynamicTable.Clear

    dynamicTable.Rows = 1
    dynamicTable.Cols = rs.Fields.Count

    For c = 0 To rs.Fields.Count - 1
        dynamicTable.TextMatrix(0, c) = rs.Fields(c).Name
    Next c

    Do Until rs.EOF
        dynamicTable.Rows = dynamicTable.Rows + 1
        r = dynamicTable.Rows - 1
        For c = 0 To rs.Fields.Count - 1
            If rs.Fields(c).Name = "Logtime" Then
            dynamicTable.ColWidth(c) = 2500
            dynamicTable.TextMatrix(r, c) = _
            Format(DateAdd("n", 330, rs.Fields(c).Value), "yyyy-mm-dd hh:mm:ss")
            Else
                dynamicTable.TextMatrix(r, c) = rs.Fields(c).Value & ""
            
            End If
        Next c
        rs.MoveNext
    Loop
    LogDiagnosticsMessage "Data loaded successfully"

CleanUp:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If

    Set rs = Nothing
    Set conn = Nothing

    LogDiagnosticsMessage "Resources released"
    Exit Sub

ErrorHandler:

    LogDiagnosticsMessage "Error " & Err.Number & _
                          ": " & Err.Description
    MsgBox "Error: " & Err.Description
    Resume CleanUp
End Sub