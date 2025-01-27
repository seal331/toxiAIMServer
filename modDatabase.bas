Attribute VB_Name = "modDatabase"
Option Explicit

Private Conn As ADODB.Connection

' Initialize the connection to the database
Public Function InitializeDatabase() As Boolean
    Dim strConnString As String
    
    On Error GoTo ErrorHandler
    
    ' Build the connection string via settings
    strConnString = BuildConnectionString()
    
    ' Log the connection string
    LogInformation "Database", "Initializing connection..."
    
    ' Connect to the database via the connection string
    Set Conn = New ADODB.Connection
    Conn.Open strConnString
    
    ' Log that we successfully connected
    LogInformation "Database", "Connection successful!"
    
    ' Return True and exit the initializer:
    InitializeDatabase = True
    Exit Function
    
ErrorHandler:
    LogError "Database", "Unable to connect to database: " & Err.Description
    InitializeDatabase = False
End Function

' Terminates the connection to the database
Public Sub TerminateDatabase()
    If Not Conn Is Nothing Then
        If Conn.State = adStateOpen Then Conn.Close
        Set Conn = Nothing
    End If
    
    LogInformation "Database", "Disconnected."
End Sub

' Execute a query without a return value (i.e. INSERT, UPDATE, DELETE)
Public Sub ExecuteNonQuery(ByVal strSQL As String)
    On Error GoTo ErrorHandler
    
    ' Ensure we have a open connection
    If Conn Is Nothing Or Conn.State <> adStateOpen Then
        Err.Raise vbObjectError, "modDatabase.ExecuteNonQuery", "Connection is not open"
    End If
    
    ' Execute the given SQL
    Conn.Execute strSQL
    
    ' Log the SQL query we just executed
    LogVerbose "Database", "Executed non-query `" & strSQL & "`"
    
    Exit Sub
    
ErrorHandler:
    LogError "Database", "Unable to execute non-query `" & strSQL & "`: " & Err.Description
End Sub

' Execute a query with a return value (i.e. SELECT)
Public Function ExecuteQuery(ByVal strSQL As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    ' Ensure we have a open connection
    If Conn Is Nothing Or Conn.State <> adStateOpen Then
        Err.Raise vbObjectError, "modDatabase.ExecuteQuery", "Connection is not open"
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = Conn.Execute(strSQL)
    Set ExecuteQuery = rst
    
    ' Log the SQL query we just executed
    LogVerbose "Database", "Executed query `" & strSQL & "`"
    
    Exit Function

ErrorHandler:
    LogError "Database", "Unable to execute query `" & strSQL & "`: " & Err.Description
End Function

' Execute a prepared query without a return value
Public Sub ExecutePreparedNonQuery(ByVal strSQL As String, ParamArray vntParams() As Variant)
    On Error GoTo ErrorHandler
    
    ' Ensure we have a open connection
    If Conn Is Nothing Or Conn.State <> adStateOpen Then
        Err.Raise vbObjectError, "modDatabase.ExecutePreparedNonQuery", "Connection is not open"
    End If
    
    Dim cmd As ADODB.Command
    Dim i As Long
    
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandText = strSQL
        .CommandType = adCmdText
        
        For i = LBound(vntParams) To UBound(vntParams)
            .Parameters.Append ConvertParameter(cmd, vntParams(i))
        Next i
        
        .Execute
    End With
    
    ' Log the SQL query we just executed
    LogVerbose "Database", "Executed prepared query `" & strSQL & "` with " & Join(vntParams, ", ")
    
    Exit Sub

ErrorHandler:
    LogError "Database", "Unable to execute prepared non-query `" & strSQL & "`: " & Err.Description
    LogError "Database", "Parameters: " & Join(vntParams, ", ")
End Sub

' Execute a prepared query with a return value
Public Function ExecutePreparedQuery(ByVal strSQL As String, ParamArray vntParams() As Variant) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    ' Ensure we have a open connection
    If Conn Is Nothing Or Conn.State <> adStateOpen Then
        Err.Raise vbObjectError, "modDatbase.ExecutedPreparedQuery", "Connection is not open"
    End If
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim i As Long
    
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandText = strSQL
        .CommandType = adCmdText
        
        For i = LBound(vntParams) To UBound(vntParams)
            .Parameters.Append ConvertParameter(cmd, vntParams(i))
        Next i
        
        Set rst = .Execute
        Set ExecutePreparedQuery = rst
    End With
    
    ' Log the SQL query we just executed
    LogVerbose "Database", "Executed prepared query `" & strSQL & "` with " & Join(vntParams, ", ")
    
    Exit Function
    
ErrorHandler:
    LogError "Database", "Unable to execute prepared query `" & strSQL & "`: " & Err.Description
    LogError "Database", "Parameters: " & Join(vntParams, ", ")
End Function

' Convert a Variant to a ADODB.Parameter
Public Function ConvertParameter(ByRef cmd As ADODB.Command, ByVal vnt As Variant) As ADODB.Parameter
    If IsNull(vnt) Or (IsByteArray(vnt) And IsBytesEmpty(vnt)) Then
        Set ConvertParameter = cmd.CreateParameter(, adBSTR, adParamInput, , Null)
        Exit Function
    End If
    
    Select Case VarType(vnt)
        Case vbArray + vbByte
            ' Byte array (e.g. binary data)
            Set ConvertParameter = cmd.CreateParameter(, adLongVarBinary, adParamInput, GetBytesLength(vnt), vnt)
        
        Case vbInteger, vbLong
            ' Integer and long
            Set ConvertParameter = cmd.CreateParameter(, adInteger, adParamInput, , vnt)
        
        Case vbSingle, vbDouble
            ' Floating-point numbers
            Set ConvertParameter = cmd.CreateParameter(, adDouble, adParamInput, , vnt)
        
        Case vbDate
            ' Date/Time
            Set ConvertParameter = cmd.CreateParameter(, adDBTimeStamp, adParamInput, , vnt)
            
        Case vbBoolean
            ' Boolean
            Set ConvertParameter = cmd.CreateParameter(, adBoolean, adParamInput, , vnt)
            
        Case Else
            Set ConvertParameter = cmd.CreateParameter(, adBSTR, adParamInput, , vnt)
    End Select
End Function

' Builds a connection string from the application's settings.
Private Function BuildConnectionString() As String
    Dim strDriver As String, strHost As String, strPort As String
    Dim strUserID As String, strPassword As String, strDbName As String
    
    strDriver = ValidateSetting(g_strDatabaseDriver, "Database driver must not be blank! Please edit it in the Settings menu.")
    strHost = ValidateSetting(g_strDatabaseHost, "Database host must not be blank! Please edit it in the Settings menu.")
    strPort = ValidateSetting(g_lngDatabasePort, "Database port must not be blank and must be numerical!", True)
    strUserID = ValidateSetting(g_strDatabaseUserID, "Database User ID must not be blank! Please edit it in the Settings menu.")
    strPassword = ValidateSetting(g_strDatabasePassword, "Database password must not be blank! Please edit it in the Settings menu.")
    strDbName = ValidateSetting(g_strDatabaseName, "Database name must not be blank! Please edit it in the Settings menu.")
    
    ' Construct the connection string
    BuildConnectionString = "Driver={" & strDriver & "};" & _
                            "Server=" & strHost & ";" & _
                            "Port=" & strPort & ";" & _
                            "User=" & strUserID & ";" & _
                            "Password=" & strPassword & ";" & _
                            "Database=" & strDbName & ";" & _
                            "Option=3;"
End Function

' Validates the setting by raising an error if empty and/or if specified, not numerical.
Private Function ValidateSetting(ByVal strValue As String, ByVal strErrorMessage As String, Optional ByVal blnNumerical As Boolean = False) As String
    If Trim(strValue) = "" Or (blnNumerical And Not IsNumeric(strValue)) Then
        Err.Raise vbObjectError, "modDatabase.ValidateSetting", strErrorMessage
    End If
    
    ValidateSetting = strValue
End Function
