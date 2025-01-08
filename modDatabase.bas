Attribute VB_Name = "modDatabase"
Option Explicit

Private Conn As ADODB.Connection

' Initialize the connection to the database
Public Function InitializeDatabase() As Boolean
    Dim strConnString As String
    
    On Error GoTo ErrorHandler
    
    ' Build the connection string via settings
    strConnString = BuildConnectionString
    
    ' Log the connection string
    LogInformation "Database", "Using connection string: " & strConnString
    
    ' Connect to the database via the connection string
    Set Conn = New ADODB.Connection
    Conn.Open strConnString
    
    ' Log that we successfully connected
    LogInformation "Database", "Connected!"
    
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
        If Conn.State = adStateOpen Then
            Conn.Close
        End If
        
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
    
    Dim RS As ADODB.Recordset
    Set RS = Conn.Execute(strSQL)
    Set ExecuteQuery = RS
    
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
    
    Dim oCmd As ADODB.Command, i As Long
    Set oCmd = New ADODB.Command
    
    With oCmd
        .ActiveConnection = Conn
        .CommandText = strSQL
        .CommandType = adCmdText
        
        For i = LBound(vntParams) To UBound(vntParams)
            .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, Len(vntParams(i)), vntParams(i))
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
    
    Dim oCmd As ADODB.Command, RS As ADODB.Recordset, i As Long
    Set oCmd = New ADODB.Command
    
    With oCmd
        .ActiveConnection = Conn
        .CommandText = strSQL
        .CommandType = adCmdText
        
        For i = LBound(vntParams) To UBound(vntParams)
            .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, Len(vntParams(i)), vntParams(i))
        Next i
        
        Set RS = .Execute
        Set ExecutePreparedQuery = RS
    End With
    
    ' Log the SQL query we just executed
    LogVerbose "Database", "Executed prepared query `" & strSQL & "` with " & Join(vntParams, ", ")
    
    Exit Function
    
ErrorHandler:
    LogError "Database", "Unable to execute prepared query `" & strSQL & "`: " & Err.Description
    LogError "Database", "Parameters: " & Join(vntParams, ", ")
End Function

'
Private Function BuildConnectionString() As String
    Dim strDriver As String, strHost As String, strPort As String
    Dim strUserID As String, strPassword As String, strDbName As String
    
    With AppSettings
        ' Read database settings:
        strDriver = ValidateSetting(.Database.Driver, "Database driver must not be blank! Please edit it in the Settings menu.")
        strHost = ValidateSetting(.Database.Host, "Database host must not be blank! Please edit it in the Settings menu.")
        strPort = ValidateSetting(CStr(.Database.Port), "Database port must not be blank and must be numerical!", True)
        strUserID = ValidateSetting(.Database.UserID, "Database User ID must not be blank! Please edit it in the Settings menu.")
        strPassword = ValidateSetting(.Database.Password, "Database password must not be blank! Please edit it in the Settings menu.")
        strDbName = ValidateSetting(.Database.Name, "Database name must not be blank! Please edit it in the Settings menu.")
    End With
    
    ' Construct the connection string
    BuildConnectionString = "Driver={" & strDriver & "};" & _
                            "Server=" & strHost & ";" & _
                            "Port=" & strPort & ";" & _
                            "User=" & strUserID & ";" & _
                            "Password=" & strPassword & ";" & _
                            "Database=" & strDbName & ";"
End Function

' Validates the setting by raising an error if empty and/or if specified, not numerical.
Private Function ValidateSetting(ByVal strValue As String, ByVal strErrorMessage As String, Optional ByVal blnNumerical As Boolean = False) As String
    If Trim(strValue) = "" Or (blnNumerical And Not IsNumeric(strValue)) Then
        Err.Raise vbObjectError, "modDatabase.ValidateSetting", strErrorMessage
    End If
    
    ValidateSetting = strValue
End Function
