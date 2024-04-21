    Public Function IsTableInDatabase(ByVal TableName As String) As Boolean

        Dim tbExists As DataTable
        Dim Sql As String
        Dim RetVal As Boolean

        'How to find if a table exists in a database
        'open a recordset with the following sql statement:
        'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
        'If the recordset it at eof then the table doesn't exist
        'if it has a record then the table does exist.

        On Error GoTo IsTableInDatabase_Error

        Sql = "SELECT name FROM sysobjects WHERE " &
            "xtype = 'U' " &
            "AND name = '" & TableName & "'"
        conDB()
        tbExists = GetTableData(Sql)



        If tbExists.Rows.Count > 0 Then 'There is no table <TableName> in database
            RetVal = True
        Else
            RetVal = False
        End If
        IsTableInDatabase = RetVal

        Exit Function

IsTableInDatabase_Error:

        Dim strES As String
        Dim intEL As Integer
        intEL = Erl()
        strES = Err.Description
        LogError("modDbDesign", "IsTableInDatabase", intEL, strES, Sql)
    End Function






'used log error function

 Public Sub LogError(ByVal ModuleName As String,
                 ByVal ProcedureName As String,
                 ByVal ErrorLineNumber As Integer,
                 ByVal ErrorDescription As String,
                 Optional ByVal SQLStatement As String = "",
                 Optional ByVal EventDesc As String = "")

     Dim Sql As String
     Dim MyMachineName As String
     Dim Vers As String
     Dim UID As String

     On Error Resume Next

     UID = AddTicks(UserName)

     SQLStatement = AddTicks(SQLStatement)

     ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "[MSSQL]")
     ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver]", "[SQL]")
     ErrorDescription = AddTicks(ErrorDescription)
     Dim location = Assembly.GetExecutingAssembly().Location
     Dim appName = Path.GetFileName(location)

     'Vers = Application.ProductVersion
     Vers = ""

     MyMachineName = System.Net.Dns.GetHostName

     Sql = "IF NOT EXISTS " &
             "    (SELECT * FROM ErrorLog WHERE " &
             "     ModuleName = '" & ModuleName & "' " &
             "     AND ProcedureName = '" & ProcedureName & "' " &
             "     AND ErrorLineNumber = '" & ErrorLineNumber & "' " &
             "     AND AppName = '" & appName & "' " &
             "     AND AppVersion = '" & Vers & "' ) " &
             "  INSERT INTO ErrorLog (" &
             "    ModuleName, ProcedureName, ErrorLineNumber, SQLStatement, " &
             "    ErrorDescription, UserName, MachineName, Eventdesc, AppName, AppVersion, EventCounter, eMailed) " &
             "  VALUES  ('" & ModuleName & "', " &
             "           '" & ProcedureName & "', " &
             "           '" & ErrorLineNumber & "', " &
             "           '" & SQLStatement & "', " &
             "           '" & ErrorDescription & "', " &
             "           '" & UID & "', " &
             "           '" & MyMachineName & "', " &
             "           '" & AddTicks(EventDesc) & "', " &
             "           '" & appName & "', " &
             "           '" & Vers & "', " &
             "           '1', '0') " &
   "ELSE "
     Sql = Sql & "  UPDATE ErrorLog " &
             "  SET SQLStatement = '" & SQLStatement & "', " &
             "  ErrorDescription = '" & ErrorDescription & "', " &
             "  MachineName = '" & MyMachineName & "', " &
             "  DateTime = getdate(), " &
             "  UserName = '" & UID & "', " &
             "  EventCounter = COALESCE(EventCounter, 0) + 1 " &
             "  WHERE ModuleName = '" & ModuleName & "' " &
             "  AND ProcedureName = '" & ProcedureName & "' " &
             "  AND ErrorLineNumber = '" & ErrorLineNumber & "' " &
             "  AND AppName = '" & appName & "' " &
             "  AND AppVersion = '" & Vers & "'"

     AddEditData(Sql)

 End Sub


'used gettable data function

Public Function GetTableData(ByVal strSQL As String) As DataTable


    GetTableData = Nothing
    Try
        conDB()
        cmd = New SqlCommand(strSQL, con)

        Dim adpt As New SqlDataAdapter(cmd)

        Dim l_dt As New DataTable

        adpt.Fill(l_dt)

        If Not IsNothing(l_dt) Then
            GetTableData = l_dt
        End If
    Catch ex As Exception
        Dim strES As String
        Dim intEL As Integer

        intEL = Erl()
        strES = ex.Message
        LogError("modFuntion", "GetTableData", intEL, strES)
    Finally
        cmd.Dispose()
        con.Close()
    End Try
End Function














