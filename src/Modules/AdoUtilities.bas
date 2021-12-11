Attribute VB_Name = "AdoUtilities"
'@Folder("ADODBDataAccess.Utils")
'@ModuleDescription("Common utility methods that extend the functionality of the API")
Option Explicit

Private Const TRANSACTION_DDL_ADODB_CONNECTION_PROPERTY_NAME As String = "Transaction DDL"

Public Function GetFieldValueOrDefault(ByVal rs As ADODB.Recordset, ByVal fieldNameOrIndex As Variant) As Variant
    DbErrors.GuardExpression Not (VarType(fieldNameOrIndex) = vbInteger Or VarType(fieldNameOrIndex) = vbLong Or VarType(fieldNameOrIndex) = vbString), _
                             ThisWorkbook.VBProject.name & "AdoUtilities.GetFieldValueOrDefault", "'fieldNameOrIndex' must be of Type 'Integer', 'Long', or 'String'"
    On Error Resume Next
    GetFieldValueOrDefault = rs.Fields(fieldNameOrIndex).value
    On Error GoTo 0
End Function

Public Function IIfIsNullOrEmpty(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If IsNullOrEmpty(value) Then
        IIfIsNullOrEmpty = replacement
        
    Else
        IIfIsNullOrEmpty = value
        
    End If
End Function
		
Public Function IIfIsNullOrWhiteSpace(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If IsNullOrWhiteSpace(value) Then
        IIfIsNullOrWhiteSpace = replacement
        
    Else
        IIfIsNullOrWhiteSpace = value
        
    End If
End Function

Public Function IIfIsEmptyOrWhitespace(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If IsEmptyOrWhitespace(value) Then
        IIfIsEmptyOrWhitespace = replacement
        
    Else
        IIfIsEmptyOrWhitespace = value
        
    End If
End Function

Public Function IIfIsEmptyValue(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If IsEmptyValue(value) Then
        IIfIsEmptyValue = replacement
        
    Else
        IIfIsEmptyValue = value
        
    End If
End Function

Public Function IsNullOrEmpty(ByVal value As Variant) As Boolean
    If VBA.IsNull(value) Then
        IsNullOrEmpty = True
        
        Exit Function
        
    End If
    
    IsNullOrEmpty = IsEmptyValue(value)
End Function

Public Function IsNullOrWhiteSpace(ByVal value As Variant) As Boolean
    If VBA.IsNull(value) Then
        IsNullOrWhiteSpace = True
        
        Exit Function
        
    End If
    
    IsNullOrWhiteSpace = (LenB(RemoveNullOrWhiteSpace(value)) = 0)
End Function

Public Function IsEmptyOrWhitespace(ByVal value As Variant) As Boolean
    If IsEmptyValue(value) Then 
        IsEmptyOrWhitespace = True
        
        Exit Function
        
    End If
    
    IsEmptyOrWhitespace = (LenB(RemoveNullOrWhiteSpace(value)) = 0)
End Function

Public Function RemoveNullOrWhiteSpace(ByVal value As Variant) As Variant
    If IsNull(value) Then
        RemoveNullOrWhiteSpace = vbNullString
        
        Exit Function
        
    End If

    If IsEmptyValue(value) Then 
        RemoveNullOrWhiteSpace = vbNullString
        
        Exit Function
        
    End If
                                
    RemoveNullOrWhiteSpace = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(value, Chr$(0), vbNullString), Chr$(9), vbNullString), Chr$(10), vbNullString), Chr$(11), vbNullString), Chr$(12), vbNullString), Chr$(13), vbNullString), Chr$(14), vbNullString), Chr$(160), vbNullString), " ", vbNullString))
End Function

Public Function IsEmptyValue(ByVal value As Variant) As Boolean 
    If IsEmpty(value) Then
        IsEmptyValue = True
        
        Exit Function
        
    End If
    
    If value = vbNullString Then
        IsEmptyValue = True
        
        Exit Function
        
    End If
    
    If value = "" Then
        IsEmptyValue = True
        
        Exit Function
        
    End If
			
    IsEmptyValue = (LenB(value) = 0)
End Function 
							
Public Function NullIf(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If VBA.IsNull(value) Then
        NullIf = replacement
    
    Else
        NullIf = value
    
    End If
End Function

Public Function Coalesce(ParamArray args() As Variant) As Variant
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

Public Function IsRecordsetClosed(ByVal rs As ADODB.Recordset) As Boolean
    Dim result As Boolean
    result = False
    
    If Not rs Is Nothing Then
        result = (Not ((rs.State And ADODB.ObjectStateEnum.adStateOpen) = ADODB.ObjectStateEnum.adStateOpen))
    
    End If

    IsRecordsetClosed = result
End Function

Public Function IsRecordsetOpen(ByVal rs As ADODB.Recordset) As Boolean
    Dim result As Boolean
    result = False
    
    If Not rs Is Nothing Then
        result = ((rs.State And ADODB.ObjectStateEnum.adStateOpen) = ADODB.ObjectStateEnum.adStateOpen)
    
    End If

    IsRecordsetOpen = result
End Function

Public Function IsRecordsetEmpty(ByVal rs As ADODB.Recordset) As Boolean
    IsRecordsetEmpty = (rs.EOF And rs.BOF)
End Function

Public Function SupportsTransactions(ByVal db As ADODB.Connection) As Boolean
    SupportsTransactions = ConnectionPropertyExists(db, TRANSACTION_DDL_ADODB_CONNECTION_PROPERTY_NAME)
End Function

Public Function ConnectionPropertyExists(ByVal db As ADODB.Connection, ByVal fieldName As String) As Boolean
    Dim errorCount As Long
    errorCount = db.Errors.Count
    
    On Error Resume Next
    ConnectionPropertyExists = db.Properties(fieldName)
    On Error GoTo 0
    
    'must clear errors from connection
    If db.Errors.Count > errorCount Then
        db.Errors.Clear
        
    End If
End Function

Public Function RecordsetPropertyExists(ByRef rs As ADODB.Recordset, ByVal fieldName As String) As Boolean
    Dim errorCount As Long
    errorCount = rs.ActiveConnection.Errors.Count
    
    On Error Resume Next
    RecordsetPropertyExists = rs.Properties(fieldName)
    On Error GoTo 0
    
    'must clear errors from connection
    If rs.ActiveConnection.Errors > errorCount Then
        rs.ActiveConnection.Errors.Clear
        
    End If
End Function

Public Function GetNewDataFromRecordset(ByVal existingDataKvps As Scripting.Dictionary, ByVal sourceRecordset As ADODB.Recordset, ByVal fieldNameOrIndex As Variant) As Variant()
    DbErrors.GuardExpression Not (VarType(fieldNameOrIndex) = vbInteger Or VarType(fieldNameOrIndex) = vbLong Or VarType(fieldNameOrIndex) = vbString), _
                             ThisWorkbook.VBProject.name & "AdoUtilities.GetNewDataFromRecordset", "'fieldNameOrIndex' must be of Type 'Integer', 'Long', or 'String'"
                             
    'if not forward only cursor then move first
    If Not ((sourceRecordset.CursorType And adOpenForwardOnly) = adOpenForwardOnly) Then
        If Not sourceRecordset.BOF Then sourceRecordset.MoveFirst
        
    End If
    
    Dim tempArray() As Variant
    ReDim tempArray(sourceRecordset.Fields.Count - 1)
    
    Dim newDataKvps As Scripting.Dictionary
    Set newDataKvps = New Scripting.Dictionary
    
    Dim key As Variant
    Dim i As Long, j As Long
    Dim fld As ADODB.Field
    Do While Not sourceRecordset.EOF
        key = sourceRecordset.Fields(fieldNameOrIndex).value
    
        If Not existingDataKvps.Exists(key) Then
            For Each fld In sourceRecordset.Fields
                tempArray(j) = fld.value
            
                j = j + 1
            
            Next
            
            newDataKvps(key) = tempArray
            j = 0
        
        End If
        
        sourceRecordset.MoveNext
        i = i + 1
        
    Loop

    i = 0
    Dim result() As Variant
    If newDataKvps.Count > 0 Then
        ReDim result(0 To newDataKvps.Count - 1, 0 To sourceRecordset.Fields.Count - 1)
        
        If IsArray(tempArray) Then Erase tempArray
        
        For Each key In newDataKvps.Keys
            tempArray = newDataKvps(key)
            
            For j = LBound(tempArray) To UBound(tempArray)
                result(i, j) = tempArray(j)
                
            Next
            
            i = i + 1
            
        Next
        
        
    End If

    If Not ((sourceRecordset.CursorType And adOpenForwardOnly) = adOpenForwardOnly) Then
        If Not sourceRecordset.BOF Then sourceRecordset.MoveFirst
        
    End If
    
    GetNewDataFromRecordset = result
End Function

Public Function ToFieldIndex(ByVal sourceFields As ADODB.Fields, ByVal fieldName As String) As Long
    Dim result As Long
    result = Empty
    
    Dim fieldNameLocal As String
    fieldNameLocal = UCase$(Trim$(fieldName))
    
    Dim fieldNamesArray() As Variant
    ReDim fieldNamesArray(sourceFields.Count - 1)
    
    Dim i As Long

    For i = 0 To sourceFields.Count - 1
        If UCase$(Trim$(sourceFields.Item(i).name)) = fieldNameLocal Then
            result = i
            Exit For
            
        End If
        
    Next i
    
    ToFieldIndex = result
End Function

Public Function RecordsetToDictionary(ByVal rs As ADODB.Recordset, ByVal fieldNameOrIndex As Variant, _
Optional ByVal includeKeyFieldInValues As Boolean = False, Optional ByVal kekCompareMode As VbCompareMethod = vbBinaryCompare) As Scripting.Dictionary
    DbErrors.GuardExpression Not (VarType(fieldNameOrIndex) = vbInteger Or VarType(fieldNameOrIndex) = vbLong Or VarType(fieldNameOrIndex) = vbString), _
                             ThisWorkbook.VBProject.name & "AdoUtilities.RecordsetToDictionary", "'fieldNameOrIndex' must be of Type 'Integer', 'Long', or 'String'"
                             
    Dim result As Scripting.Dictionary
    Set result = New Scripting.Dictionary
    result.CompareMode = kekCompareMode
    
    Dim upperBound As Long
    upperBound = IIf((Not includeKeyFieldInValues) And (rs.Fields.Count > 2), rs.Fields.Count - 2, rs.Fields.Count - 1)
'    upperBound = IIf(Not (includeKeyFieldInValues) And (rs.Fields.Count > 2), rs.Fields.Count - 2, rs.Fields.Count - 1)
    
    Dim tempArray() As Variant
    ReDim tempArray(upperBound)
    
    Dim key As Variant
    Dim i As Long
    Dim j As Long
    Dim fld As ADODB.Field
    If includeKeyFieldInValues Then
        Do While Not rs.EOF
            key = rs.Fields(fieldNameOrIndex).value
        
            If Not result.Exists(key) Then
                For Each fld In rs.Fields
                    tempArray(i) = fld.value
                    i = i + 1
                Next
                
                result(key) = tempArray
                i = 0
            
            End If
            
            rs.MoveNext
        Loop
    
    Else
        If rs.Fields.Count > 2 Then
            If VarType(fieldNameOrIndex) = vbString Then
                Do While Not rs.EOF
                    key = rs.Fields(fieldNameOrIndex).value
                
                    If Not result.Exists(key) Then
                        For Each fld In rs.Fields
                            If UCase$(fld.name) <> fieldNameOrIndex Then
                                tempArray(i) = fld.value
                                i = i + 1
                            End If
                        Next
                        
                        result(key) = tempArray
                        i = 0
                    
                    End If
                    
                    rs.MoveNext
                Loop
        
           Else
                Do While Not rs.EOF
                    key = rs.Fields(fieldNameOrIndex).value
                
                    If Not result.Exists(key) Then
                        For Each fld In rs.Fields
                            If j <> fieldNameOrIndex Then
                                tempArray(i) = fld.value
                                i = i + 1
                            
                            End If
                            
                            j = j + 1
                        Next
                        
                        result(key) = tempArray
                        
                        i = 0
                        j = 0
                    
                    End If
                    
                    rs.MoveNext
                Loop
                
            End If
        Else
            Dim tempVal As Variant
            
            If VarType(fieldNameOrIndex) = vbString Then
                Do While Not rs.EOF
                    key = rs.Fields(fieldNameOrIndex).value
                
                    If Not result.Exists(key) Then
                        For Each fld In rs.Fields
                            If UCase$(fld.name) <> fieldNameOrIndex Then
                                tempVal = fld.value
                            
                            End If
                        
                        Next
                        
                        result(key) = tempVal
                    
                    End If
                    
                    rs.MoveNext
                Loop
                
            Else
                Do While Not rs.EOF
                    key = rs.Fields(fieldNameOrIndex).value
                
                    If Not result.Exists(key) Then
                        For Each fld In rs.Fields
                            If j <> fieldNameOrIndex Then
                                tempVal = fld.value
                            
                            End If
                            
                            j = j + 1
                        Next
                        
                        result(key) = tempVal
                    
                        j = 0
                    End If
                    
                    rs.MoveNext
                Loop
            
            End If
        
        End If
    
    End If

    If Not ((rs.CursorType And adOpenForwardOnly) = adOpenForwardOnly) Then
        rs.MoveFirst
    
    End If
    
    Set RecordsetToDictionary = result
End Function

Public Function RecordsetToArray(ByVal rs As ADODB.Recordset, Optional ByVal transpose As Boolean = False) As Variant()
    If transpose Then
        Dim tempArray() As Variant
        ReDim tempArray(rs.Recordcount - 1, rs.Fields.Count - 1)

        Dim recordsArray As Variant
        recordsArray = rs.GetRows()
        
        Dim j As Long, i As Long
        For j = 0 To UBound(recordsArray, 2)
            For i = 0 To UBound(recordsArray, 1)
                tempArray(j, i) = recordsArray(i, j)
            
            Next i
        
        Next j
        
        RecordsetToArray = tempArray
    
    Else
        RecordsetToArray = rs.GetRows()
    
    End If
End Function

Public Function FieldNamesToArray(ByVal rs As ADODB.Recordset, Optional ByVal makeProperCase As Boolean = False, _
Optional ByVal removeUnderscores As Boolean = False) As String()
    Dim result() As String
    ReDim result(rs.Fields.Count - 1)
    
    Dim i As Long
    Dim tempName As String
    For i = 0 To UBound(result)
        tempName = rs.Fields.Item(i).name
        
        If makeProperCase Then
            tempName = Application.Proper(tempName)
            
        End If
        
        If removeUnderscores Then
            tempName = Replace(tempName, "_", " ")
        
        End If
        
        result(i) = tempName
    
    Next i
    
    FieldNamesToArray = result
End Function

Public Sub FieldNamesToRange(ByRef fieldNamesArray As Variant, ByVal destinationRange As Range, _
Optional ByVal verticalOrientation As Boolean = False)
'fieldNamesArray must be a single dimensional array
'destinationRange should be like: Worksheet.Range(columnletter)
    DbErrors.GuardExpression destinationRange.Columns.Count > 1 Or destinationRange.Rows.Count > 1, _
                             ThisWorkbook.VBProject.name & "AdoUtilities.FieldNamesToRange", "'destinationRange' must be a single cell"
                             
    Dim lowerBound As Long
    lowerBound = LBound(fieldNamesArray)
    
    Dim upperBound As Long
    upperBound = UBound(fieldNamesArray)
    
    Dim result As Variant
    If verticalOrientation Then
        'write as rows
        ReDim result(lowerBound To upperBound, lowerBound To lowerBound)
        
        Dim i As Long
        For i = lowerBound To upperBound
            result(i, lowerBound) = fieldNamesArray(i)
        
        Next i
        
        destinationRange.Resize(upperBound).Value2 = result
    
    Else
        'default write as columns
        ReDim result(lowerBound To lowerBound, lowerBound To upperBound)
        
        Dim j As Long
        For j = lowerBound To upperBound
            result(lowerBound, j) = fieldNamesArray(j)
        
        Next j
        
        destinationRange.Resize(1, upperBound + 1).Value2 = result

    
    End If
End Sub

Public Function CreateAdoConnection(ByVal connString As String, ByVal cursorLocation As ADODB.CursorLocationEnum, _
Optional ByVal permissionsMode As ConnectModeEnum = ConnectModeEnum.adModeShareDenyNone) As ADODB.Connection
    Dim result As ADODB.Connection
    Set result = New ADODB.Connection
    
    'must set before opening
    result.cursorLocation = cursorLocation
    result.Mode = permissionsMode
    
    result.Open connString
    
    Set CreateAdoConnection = result
End Function

Public Function ExecuteWorkSheetQuery(ByVal db As ADODB.Connection, ByVal worksheetName As String, _
Optional FieldNames As String = "*", Optional ByVal joinClause As String = vbNullString, _
Optional ByVal predicateExpression As String = vbNullString, _
Optional ByVal orderByExpression As String = vbNullString, _
Optional ByVal rangeAddress As String = vbNullString) As ADODB.Recordset
    predicateExpression = IIf((predicateExpression <> vbNullString And InStr(UCase$(predicateExpression), "WHERE") = 0), "WHERE " & predicateExpression, predicateExpression)
    
    orderByExpression = IIf((orderByExpression <> vbNullString And InStr(UCase$(orderByExpression), "ORDER BY") = 0), "ORDER BY " & orderByExpression, orderByExpression)

    Dim queryString As String
    If rangeAddress <> vbNullString Then
        queryString = "select   " & SanitizeDelimitedFieldNames(FieldNames) & vbNewLine & _
                      "from     " & FormatExcelSourceForAdoQuery(worksheetName, rangeAddress) & vbNewLine & _
                      joinClause & vbNewLine & _
                      predicateExpression & vbNewLine & _
                      orderByExpression
                      
    Else
        queryString = "select   " & SanitizeDelimitedFieldNames(FieldNames) & vbNewLine & _
                      "from     " & FormatExcelWorksheetNameForAdoQuery(worksheetName) & vbNewLine & _
                      joinClause & vbNewLine & _
                      predicateExpression & vbNewLine & _
                      orderByExpression
    
    End If
                  
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
    Set .ActiveConnection = db
        .CommandText = queryString
        .commandType = adCmdText
    End With
        
    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    result.CursorType = adOpenKeyset
    result.LockType = adLockOptimistic
    
    Set result = cmd.Execute()
    
    Set ExecuteWorkSheetQuery = result
End Function

Public Function QuickExecuteWorkSheetQuery(ByVal workbookFileFullPath As String, ByVal worksheetName As String, _
Optional FieldNames As String = "*", Optional ByVal joinClause As String = vbNullString, _
Optional ByVal predicateExpression As String = vbNullString, _
Optional ByVal orderByExpression As String = vbNullString, _
Optional ByVal queryAsText As Boolean = True, _
Optional ByVal rangeAddress As String = vbNullString) As ADODB.Recordset
    Dim workBookConnectionString As String
    workBookConnectionString = GetWorkbookConnectionString(workbookFileFullPath, True, queryAsText)
                               
    predicateExpression = IIf((predicateExpression <> vbNullString And InStr(UCase$(predicateExpression), "WHERE") = 0), "WHERE " & predicateExpression, predicateExpression)
    
    orderByExpression = IIf((orderByExpression <> vbNullString And InStr(UCase$(orderByExpression), "ORDER BY") = 0), "ORDER BY " & orderByExpression, orderByExpression)
        
    Dim queryString As String
    If rangeAddress <> vbNullString Then
        queryString = "select   " & SanitizeDelimitedFieldNames(FieldNames) & vbNewLine & _
                      "from     " & FormatExcelSourceForAdoQuery(worksheetName, rangeAddress) & vbNewLine & _
                      joinClause & vbNewLine & _
                      predicateExpression & vbNewLine & _
                      orderByExpression
                      
    Else
        queryString = "select   " & SanitizeDelimitedFieldNames(FieldNames) & vbNewLine & _
                      "from     " & FormatExcelWorksheetNameForAdoQuery(worksheetName) & vbNewLine & _
                      joinClause & vbNewLine & _
                      predicateExpression & vbNewLine & _
                      orderByExpression
    
    End If
    
    Dim db As ADODB.Connection
    Set db = CreateAdoConnection(workBookConnectionString, adUseClient)
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
    Set .ActiveConnection = db
        .CommandText = queryString
        .commandType = adCmdText
    End With
        
    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    result.CursorType = adOpenKeyset
    result.LockType = adLockOptimistic
    
    Set result = cmd.Execute()
    
    Set result.ActiveConnection = Nothing
    
    Set QuickExecuteWorkSheetQuery = result
End Function

Public Function ExecuteCustomWorkSheetQuery(ByVal db As ADODB.Connection, ByVal sql As String) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
    Set .ActiveConnection = db
        .CommandText = sql
        .commandType = adCmdText
    End With
        
    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    result.CursorType = adOpenKeyset
    result.LockType = adLockOptimistic
    
    Set result = cmd.Execute()
    
    Set ExecuteCustomWorkSheetQuery = result
End Function

Public Function QuickExecuteCustomWorkSheetQuery(ByVal workbookFileFullPath As String, ByVal sql As String, _
Optional ByVal queryAsText As Boolean = True) As ADODB.Recordset
    Dim workBookConnectionString As String
    workBookConnectionString = GetWorkbookConnectionString(workbookFileFullPath, True, queryAsText)
    
    Dim db As ADODB.Connection
    Set db = CreateAdoConnection(workBookConnectionString, adUseClient)
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
    Set .ActiveConnection = db
        .CommandText = sql
        .commandType = adCmdText
    End With
        
    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    result.CursorType = adOpenKeyset
    result.LockType = adLockOptimistic
    
    Set result = cmd.Execute()
    
    Set result.ActiveConnection = Nothing
    
    Set QuickExecuteCustomWorkSheetQuery = result
End Function

Public Function GetWorkbookConnectionString(ByVal workbookFileFullPath As String, ByVal datsetHasHeaders As Boolean, ByVal queryAsText As Boolean)
    Dim result As String
    result = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & workbookFileFullPath & _
             ";Extended Properties='"
                
    Dim extendedProperty As String
    DbErrors.GuardExpression Not TryGetExcelFileTypeConnectionStringExtendedProperty(workbookFileFullPath, extendedProperty), _
        message:="Could not retrieve Excel file type extended property for connection string due to incorrect file type/extentsion." & _
        "Accepted file types are .xlsx, .xlsm, .xlsb, .xls."
    
    result = result & extendedProperty
    
    result = result & IIf(datsetHasHeaders, ";HDR=YES", ";HDR=NO")
    
    If queryAsText Then result = result & ";IMEX=1"
    
    result = result & "'"
    
    GetWorkbookConnectionString = result
End Function

Private Function TryGetExcelFileTypeConnectionStringExtendedProperty(ByVal workbookFileFullPath As String, ByRef returnExtendedProperty As String) As Boolean
    returnExtendedProperty = vbNullString
    
    Dim fileExtentsion As String
    fileExtentsion = LCase$(Trim$(GetFileExtenstion(workbookFileFullPath)))
    
    Select Case fileExtentsion
        Case "xlsb"
            returnExtendedProperty = "Excel 12.0"
            
        Case "xlsx"
            returnExtendedProperty = "Excel 12.0 Xml"
            
        Case "xlsm"
            returnExtendedProperty = "Excel 12.0 Macro"
        
        Case "xls"
            returnExtendedProperty = "Excel 8.0"
            
    End Select
        
    TryGetExcelFileTypeConnectionStringExtendedProperty = (LenB(Trim$(Replace(returnExtendedProperty, " ", vbNullString))) > 0)
End Function

Private Function GetFileExtenstion(ByVal fileFullNameOrPath As String) As String
    GetFileExtenstion = Right$(fileFullNameOrPath, (Len(fileFullNameOrPath) - InStrRev(fileFullNameOrPath, ".")))
End Function


Public Function FormatExcelWorksheetNameForAdoQuery(ByVal worksheetName As String) As String
    FormatExcelWorksheetNameForAdoQuery = "[" & worksheetName & "$]"
End Function

Public Function FormatExcelRangeForAdoQuery(ByVal sourceRange As Range) As String
    FormatExcelRangeForAdoQuery = "[" & sourceRange.Parent.name & "$" & sourceRange.Address(False, False) & "]"
End Function


Public Function FormatExcelSourceForAdoQuery(ByVal worksheetName As String, ByVal rangeAddress As String) As String
    FormatExcelSourceForAdoQuery = "[" & worksheetName & "$" & Replace(rangeAddress, "$", vbNullString) & "]"
End Function

Public Function SanitizeDelimitedFieldNames(ByVal delimitedFieldNames As String) As String
    SanitizeDelimitedFieldNames = Replace(Replace(Replace(delimitedFieldNames, ",[]", vbNullString), ", []", vbNullString), "[]", vbNullString)
End Function
