Attribute VB_Name = "AdoUtilities"
'@Folder("ADODBDataAccess.Utils")
'@ModuleDescription("Common utility methods that extend the functionality of the API")
Option Explicit

Public Function SupportsTransactions(ByVal db As ADODB.Connection) As Boolean
    Const TRANSACTION_PROPERTY_NAME As String = "Transaction DDL"
    SupportsTransactions = ConnectionPropertyExists(db, TRANSACTION_PROPERTY_NAME)
    db.Errors.Clear
End Function

Public Function ConnectionPropertyExists(ByVal db As ADODB.Connection, ByVal fieldNameValue As String) As Boolean
    On Error Resume Next
    ConnectionPropertyExists = db.Properties(fieldNameValue)
    On Error GoTo 0
End Function

Public Function RecordsetPropertyExists(ByRef rs As ADODB.Recordset, ByVal fieldNameValue As String) As Boolean
    On Error Resume Next
    RecordsetPropertyExists = rs.Properties(fieldNameValue)
    On Error GoTo 0
End Function

Public Function IsRecordsetClosed(ByRef rs As ADODB.Recordset) As Boolean
    Dim result As Boolean
    result = False
    
    If Not rs Is Nothing Then
        result = (Not ((rs.State And ADODB.ObjectStateEnum.adStateOpen) = ADODB.ObjectStateEnum.adStateOpen))
    End If

    IsRecordsetClosed = result
End Function

Public Function IsRecordsetEmpty(ByRef rs As ADODB.Recordset) As Boolean
    IsRecordsetEmpty = (rs.EOF And rs.BOF)
End Function

Public Function RecordsetToDictionary(ByRef rs As ADODB.Recordset, ByVal keyFieldName As String) As Scripting.Dictionary
    On Error GoTo CleanFail
    Dim result As Scripting.Dictionary
    Set result = New Scripting.Dictionary
    result.Comparemode = vbTextCompare
    
    Dim arryTemp() As Variant
    ReDim arryTemp(rs.Fields.Count - 1)
    
    Dim varKey As Variant
    Dim i As Long
    Dim fld As ADODB.Field
    Do While Not rs.EOF
        varKey = rs.Fields(keyFieldName).value
    
        If Not result.Exists(varKey) Then
            
            For Each fld In rs.Fields
                arryTemp(i) = fld.value
                i = i + 1
            Next
            
            result(varKey) = arryTemp
            i = Empty
            
        End If
        
        rs.MoveNext
        
    Loop

    If Not ((rs.CursorType And adOpenForwardOnly) = adOpenForwardOnly) Then rs.MoveFirst
    
    Set RecordsetToDictionary = result

CleanExit:
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Public Function RecordsetToArray(ByRef rs As ADODB.Recordset, ByVal transpose As Boolean) As Variant
    On Error GoTo CleanFail
    If transpose Then
        Dim arryTemp() As Variant
        ReDim arryTemp(rs.RecordCount - 1, rs.Fields.Count - 1)

        Dim arrayRecords As Variant
        arrayRecords = rs.GetRows()
        
        Dim j As Long, i As Long
        For j = 0 To UBound(arrayRecords, 2)
            For i = 0 To UBound(arrayRecords, 1)
                arryTemp(j, i) = arrayRecords(i, j)
            Next i
        Next j
        
        RecordsetToArray = arryTemp
    
    Else
        RecordsetToArray = rs.GetRows()
    
    End If
    
CleanExit:
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Public Function FieldNamesToArray(ByRef rs As ADODB.Recordset, _
Optional ByVal makeProperCase As Boolean = False, _
Optional ByVal removeUnderscores As Boolean = False) As Variant
    On Error GoTo CleanFail
    Dim fieldNamesArray() As Variant
    ReDim fieldNamesArray(rs.Fields.Count - 1)
    
    Dim i As Long
    Dim tempName As String
    For i = 0 To rs.Fields.Count - 1
        tempName = rs.Fields.Item(i).name
        
        If makeProperCase Then tempName = Application.Proper(tempName)
        If removeUnderscores Then tempName = Replace(tempName, "_", " ")
        
        fieldNamesArray(i) = tempName
    Next i
    
    FieldNamesToArray = fieldNamesArray
    
CleanExit:
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Public Sub FieldNamesToRange(ByRef fieldNamesArray As Variant, ByRef destinationRange As Range, _
Optional ByVal horizontalOrientation As Boolean = True)
'fieldNamesArray must be a single dimensional array
'destinationRange should be like: Worksheet.Range(columnletter)
    Dim lowerBound As Long
    lowerBound = LBound(fieldNamesArray)
    
    Dim upperBound As Long
    upperBound = UBound(fieldNamesArray)
    
    Dim result As Variant
    If horizontalOrientation Then
        ReDim result(lowerBound To lowerBound, lowerBound To upperBound)
        
        Dim j As Long
        For j = lowerBound To upperBound
            result(lowerBound, j) = fieldNamesArray(j)
        Next j
        
        destinationRange.Resize(1, upperBound + 1).Value2 = result
    Else
        ReDim result(lowerBound To upperBound, lowerBound To lowerBound)
        
        Dim i As Long
        For i = lowerBound To upperBound
            result(i, lowerBound) = fieldNamesArray(i)
        Next i
        
        destinationRange.Resize(upperBound).Value2 = result
    End If
End Sub

Public Function GetFieldValueOrDefault(ByVal rs As ADODB.Recordset, ByVal fieldNameValue As String) As Variant
    On Error Resume Next
    GetFieldValueOrDefault = rs.Fields(fieldNameValue).value
    On Error GoTo 0
End Function

Public Function NullIf(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If VBA.IsNull(value) Then
        NullIf = replacement
    Else
        NullIf = value
    End If
End Function

Public Function IIfIsNullOrEmpty(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If Len(Trim$(Replace(Replace(Replace(value, " ", vbNullString), VBA.Chr$(10), vbNullString), VBA.Chr$(13), vbNullString))) _
                Or VBA.IsNull(value) _
        Then
        IIfIsNullOrEmpty = replacement
    Else
        IIfIsNullOrEmpty = value
    End If
End Function

Public Function Coalesce(ParamArray args()) As Variant
    Dim result As Variant
    result = Null
    
    Dim i As Long
    For i = LBound(args) To UBound(args)
        If Not VBA.IsNull(args(i)) Then
            result = args(i)
            Exit For
        End If
    Next i

    Coalesce = result
End Function

Public Function ExecuteWorkSheetQuery(ByVal db As ADODB.Connection, ByVal workSheetNameValue As String, _
Optional fieldNamesValue As String = "*", _
Optional ByVal joinClause As String = vbNullString, _
Optional ByVal predicateExpression As String = vbNullString, _
Optional ByVal orderByExpression As String = vbNullString) As ADODB.Recordset

    predicateExpression = IIf((predicateExpression <> vbNullString And InStr(UCase$(predicateExpression), "WHERE") = 0), "WHERE " & predicateExpression, predicateExpression)
    
    orderByExpression = IIf((orderByExpression <> vbNullString And InStr(UCase$(orderByExpression), "ORDER BY") = 0), "ORDER BY " & orderByExpression, orderByExpression)
        
    Dim queryString As String
    queryString = "SELECT " & fieldNamesValue & vbNewLine & _
                  "FROM [" & workSheetNameValue & "$] " & vbNewLine & _
                  joinClause & vbNewLine & _
                  predicateExpression & vbNewLine & _
                  orderByExpression
                  
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
    Set .ActiveConnection = db
        .CommandText = queryString
        .commandType = adCmdText
    End With
        
    Dim results As ADODB.Recordset
    Set results = New ADODB.Recordset
    results.CursorType = adOpenKeyset
    results.LockType = adLockOptimistic
    
    Set results = cmd.Execute()
    
    Set ExecuteWorkSheetQuery = results
End Function


Public Function QuickExecuteWorkSheetQuery(ByVal workBookFilePath As String, ByVal workSheetNameValue As String, _
Optional fieldNamesValue As String = "*", _
Optional ByVal joinClause As String = vbNullString, _
Optional ByVal predicateExpression As String = vbNullString, _
Optional ByVal orderByExpression As String = vbNullString) As ADODB.Recordset

    Dim workBookConnectionString As String
    workBookConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & workBookFilePath & _
                               ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1'"
                               
    predicateExpression = IIf((predicateExpression <> vbNullString And InStr(UCase$(predicateExpression), "WHERE") = 0), "WHERE " & predicateExpression, predicateExpression)
    
    orderByExpression = IIf((orderByExpression <> vbNullString And InStr(UCase$(orderByExpression), "ORDER BY") = 0), "ORDER BY " & orderByExpression, orderByExpression)
        
    Dim queryString As String
    queryString = "SELECT " & fieldNamesValue & vbNewLine & _
                  "FROM [" & workSheetNameValue & "$] " & vbNewLine & _
                  joinClause & vbNewLine & _
                  predicateExpression & vbNewLine & _
                  orderByExpression
    
    Dim db As ADODB.Connection
    Set db = New ADODB.Connection
    Set db = CreateAdoConnection(workBookConnectionString, adUseClient)
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
    Set .ActiveConnection = db
        .CommandText = queryString
        .commandType = adCmdText
    End With
        
    Dim results As ADODB.Recordset
    Set results = New ADODB.Recordset
    results.CursorType = adOpenKeyset
    results.LockType = adLockOptimistic
    
    Set results = cmd.Execute()
    
    Set QuickExecuteWorkSheetQuery = results
End Function

Public Function CreateAdoConnection(ByVal connString As String, ByVal cursorLocationValue As ADODB.CursorLocationEnum) As ADODB.Connection
    Dim result As ADODB.Connection
    Set result = New ADODB.Connection
    
    result.CursorLocation = cursorLocationValue  'must set before opening
    result.Open connString
    
    Set CreateAdoConnection = result
End Function
