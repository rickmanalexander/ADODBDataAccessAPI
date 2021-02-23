Attribute VB_Name = "AdoUtilities"
'@Folder("ADODBDataAccess.Utils")
'@ModuleDescription("Common utility methods that extend the functionality of the API")
Option Explicit

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

Public Function IsNullOrEmpty(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If Len(Trim$(Replace(Replace(Replace(value, " ", vbNullString), VBA.Chr$(10), vbNullString), VBA.Chr$(13), vbNullString))) _
        Or VBA.IsNull(value) _
    Then
        IsNullOrEmpty = replacement
    
	Else
        IsNullOrEmpty = value
    
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

Public Function IsForwardOnlyCursor(ByVal cursorType As ADODB.CursorTypeEnum) As Boolean 
    IsForwardOnlyCursor = ((cursorType And adOpenForwardOnly) = adOpenForwardOnly)
End Function 

Public Function DbSupportsTransactions(ByVal db As ADODB.Connection) As Boolean
    Const TRANSACTION_PROPERTY_NAME As String = "Transaction DDL"
    DbSupportsTransactions = ConnectionHasProperty(db, TRANSACTION_PROPERTY_NAME)
End Function

Public Function ConnectionHasProperty(ByVal db As ADODB.Connection, ByVal fieldNameValue As String) As Boolean
    Dim errorCount As Long
    errorCount = db.Errors.Count
    
    On Error Resume Next
    ConnectionHasProperty = db.Properties(fieldNameValue)
    On Error GoTo 0
    
    'must clear errors from connection
    If db.Errors.Count > errorCount Then db.Errors.Clear
End Function

Public Function RecordsetHasProperty(ByRef rs As ADODB.Recordset, ByVal fieldNameValue As String) As Boolean
    Dim errorCount As Long
    errorCount = rs.ActiveConnection.Errors.Count
    
    On Error Resume Next
    RecordsetHasProperty = rs.Properties(fieldNameValue)
    On Error GoTo 0
    
    'must clear errors from connection
    If rs.ActiveConnection.Errors > errorCount Then rs.ActiveConnection.Errors.Clear
End Function

Public Function GetNewDataFromRecordset(ByVal existingDataKvps As Scripting.Dictionary, ByVal rs As ADODB.Recordset, ByVal keyFieldName As String) As Variant()
    If Not IsForwardOnlyCursor(rs.CursorType) Then
        If Not rs.BOF Then rs.MoveFirst
        
    End If
    
    Dim localKeyFieldName As String
    localKeyFieldName = UCase$(Trim$(keyFieldName))
    
    Dim tempArray() As Variant
    ReDim tempArray(rs.Fields.Count - 1)
    
    Dim newDataKvps As Scripting.Dictionary
    Set newDataKvps = New Scripting.Dictionary
    
    Dim key As Variant
    Dim i As Long, j As Long
    Dim fld As ADODB.Field
    Do While Not rs.EOF
        key = rs.Fields(localKeyFieldName).value
    
        If Not existingDataKvps.Exists(key) Then
            For Each fld In rs.Fields
                tempArray(j) = fld.value
            
                j = j + 1
            
            Next
            
            newDataKvps(key) = tempArray
            j = 0
        
        End If
        
        rs.MoveNext
        i = i + 1
        
    Loop

    i = 0
    Dim result() As Variant
    If newDataKvps.Count > 0 Then
        ReDim result(0 To newDataKvps.Count - 1, 0 To rs.Fields.Count - 1)
        
        If IsArray(tempArray) Then Erase tempArray
        
        For Each key In newDataKvps.Keys
            tempArray = newDataKvps(key)
            
            For j = LBound(tempArray) To UBound(tempArray)
                result(i, j) = tempArray(j)
                
            Next
            
            i = i + 1
            
        Next
        
        
    End If

    If Not IsForwardOnlyCursor(rs.CursorType) Then
        If Not rs.BOF Then rs.MoveFirst
        
    End If
    
    GetNewDataFromRecordset = result
End Function

Public Function GetFieldIndex(ByVal sourceFields As ADODB.Fields, ByVal fieldName As String) As Long
    Dim result As Long
    result = Empty
	
    Dim fieldNameLocal As String
    fieldNameLocal = UCase$(Trim$(fieldName))
    
    Dim fieldsNames() As Variant
    ReDim fieldsNames(sourceFields.Count - 1)
    
    Dim i As Long

    For i = 0 To sourceFields.Count - 1
        If UCase$(Trim$(sourceFields.Item(i).name)) = fieldNameLocal Then
            result = i
            Exit For
            
        End If
        
    Next i
    
    GetFieldIndex = result
End Function

Public Function RecordsetToDictionary(ByVal rs As ADODB.Recordset, ByVal keyFieldName As String, _
Optional ByVal includeKeyFieldInValues As Boolean = False) As Scripting.Dictionary
    Dim result As Scripting.Dictionary
    Set result = New Scripting.Dictionary
    result.CompareMode = vbTextCompare
	
    If Not IsForwardOnlyCursor(rs.CursorType) Then
        If Not rs.BOF Then rs.MoveFirst
        
    End If
    
    Dim localKeyFieldName As String
    localKeyFieldName = UCase$(Trim$(keyFieldName))
    
    Dim upperBound As Long
    upperBound = IIf((Not includeKeyFieldInValues) And (rs.Fields.Count > 2), rs.Fields.Count - 2, rs.Fields.Count - 1)
    
    Dim tempArray() As Variant
    ReDim tempArray(upperBound)
    
    Dim key As Variant
    Dim i As Long
    Dim fld As ADODB.Field
    If includeKeyFieldInValues Then
        Do While Not rs.EOF
            key = rs.Fields(localKeyFieldName).value
        
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
            Do While Not rs.EOF
                key = rs.Fields(localKeyFieldName).value
            
                If Not result.Exists(key) Then
                    For Each fld In rs.Fields
                        If UCase$(fld.name) <> localKeyFieldName Then
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
            Dim tempVal As Variant
            Do While Not rs.EOF
                key = rs.Fields(localKeyFieldName).value
            
                If Not result.Exists(key) Then
                    For Each fld In rs.Fields
                        If UCase$(fld.name) <> localKeyFieldName Then
                            tempVal = fld.value
                        
						End If
                    
					Next
                    
                    result(key) = tempVal
                
				End If
                
                rs.MoveNext
            Loop
        
		End If
    
	End If

    If Not IsForwardOnlyCursor(rs.CursorType) Then
        If Not rs.BOF Then rs.MoveFirst
        
    End If

    Set RecordsetToDictionary = result
End Function

Public Function RecordsetToArray(ByVal rs As ADODB.Recordset, ByVal transpose As Boolean) As Variant
	Dim result As Variant 
	
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
        
        result = tempArray
    
    Else
        result = rs.GetRows()
    
    End If
   
    If Not IsForwardOnlyCursor(rs.CursorType) Then
        If Not rs.BOF Then rs.MoveFirst
        
    End If

	RecordsetToArray = result
End Function

Public Function FieldNamesToArray(ByVal rs As ADODB.Recordset, Optional ByVal makeProperCase As Boolean = False, _
Optional ByVal removeUnderscores As Boolean = False) As Variant()
    Dim result() As Variant
    ReDim result(rs.Fields.Count - 1)
    
    Dim i As Long
    Dim tempName As String
    For i = 0 To rs.Fields.Count - 1
        tempName = rs.Fields.Item(i).name
        
        If makeProperCase Then tempName = Application.Proper(tempName)
        If removeUnderscores Then tempName = Replace(tempName, "_", " ")
        
        result(i) = tempName
    Next i
    
    FieldNamesToArray = result
End Function

Public Sub FieldNamesToRange(ByRef fieldsNames As Variant, ByVal destinationRange As Range, _
Optional ByVal transpose As Boolean = False)
	'fieldsNames must be a single dimensional array
	'destinationRange should be like: Worksheet.Range(columnletter)
    Dim lowerBound As Long
    lowerBound = LBound(fieldsNames)
    
    Dim upperBound As Long
    upperBound = UBound(fieldsNames)
    
    Dim result As Variant
    If transpose Then
        ReDim result(lowerBound To upperBound, lowerBound To lowerBound)
        
        Dim i As Long
        For i = lowerBound To upperBound
            result(i, lowerBound) = fieldsNames(i)
        
		Next i
        
        destinationRange.Resize(upperBound).Value2 = result
    
	Else
        ReDim result(lowerBound To lowerBound, lowerBound To upperBound)
        
        Dim j As Long
        For j = lowerBound To upperBound
            result(lowerBound, j) = fieldsNames(j)
        
		Next j
        
        destinationRange.Resize(1, upperBound + 1).Value2 = result
    
	End If
End Sub


Public Function CreateAdoConnection(ByVal connString As String, ByVal cursorLocation As ADODB.CursorLocationEnum) As ADODB.Connection
    Dim result As ADODB.Connection
    Set result = New ADODB.Connection
    
    result.CursorLocation = cursorLocation  'must set before opening
    result.Open connString
    
    Set CreateAdoConnection = result
End Function

Public Function ExecuteWorkSheetQuery(ByVal db As ADODB.Connection, ByVal workSheetName As String, _
Optional fieldNames As String = "*", Optional ByVal joinClause As String = vbNullString, _
Optional ByVal predicateExpression As String = vbNullString, _ 
Optional ByVal orderByExpression As String = vbNullString) As ADODB.Recordset
    Dim queryString As String
    queryString = "SELECT " & SanitizeDelimitedfieldNames(fieldNames) & vbNewLine & _
                  "FROM [" & Replace(workSheetName, "$", vbNullString) & "$] " & vbNewLine & _
                  joinClause & vbNewLine & _
                  IIf((predicateExpression <> vbNullString And InStr(UCase$(predicateExpression), "WHERE") = 0), "WHERE " & predicateExpression, predicateExpression) & vbNewLine & _
                  IIf((orderByExpression <> vbNullString And InStr(UCase$(orderByExpression), "ORDER BY") = 0), "ORDER BY " & orderByExpression, orderByExpression)
                  
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

Public Function QuickExecuteWorkSheetQuery(ByVal fileFullPath As String, ByVal workSheetName As String, _
Optional fieldNames As String = "*", Optional ByVal joinClause As String = vbNullString, _
Optional ByVal predicateExpression As String = vbNullString, Optional ByVal orderByExpression As String = vbNullString, _
Optional ByVal queryAsText As Boolean = False) As ADODB.Recordset
    Dim connString As String
    connString = GetWorkbookConnectionString(fileFullPath, True, queryAsText)    

    Dim db As ADODB.Connection
    Set db = CreateAdoConnection(connString, adUseClient)
	
    Dim queryString As String
    queryString = "SELECT " & SanitizeDelimitedfieldNames(fieldNames) & vbNewLine & _
                  "FROM [" & Replace(workSheetName, "$", vbNullString) & "$] " & vbNewLine & _
                  joinClause & vbNewLine & _
                  IIf((predicateExpression <> vbNullString And InStr(UCase$(predicateExpression), "WHERE") = 0), "WHERE " & predicateExpression, predicateExpression) & vbNewLine & _
                  IIf((orderByExpression <> vbNullString And InStr(UCase$(orderByExpression), "ORDER BY") = 0), "ORDER BY " & orderByExpression, orderByExpression)
    
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

Public Function QuickExecuteCustomWorkSheetQuery(ByVal fileFullPath As String, ByVal sql As String, _
Optional ByVal queryAsText As Boolean = False) As ADODB.Recordset
    Dim connString As String
    connString = GetWorkbookConnectionString(fileFullPath, True, queryAsText)
    
    Dim db As ADODB.Connection
    Set db = CreateAdoConnection(connString, adUseClient)
    
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

Public Function GetWorkbookConnectionString(ByVal fileFullPath As String, ByVal datsetHasHeaders As Boolean, ByVal queryAsText As Boolean)
    Dim result As String
    result = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileFullPath & ";Extended Properties='Excel 12.0"
                
    result = result & IIf(datsetHasHeaders, ";HDR=YES", ";HDR=NO")
    
    If queryAsText Then result = result & ";IMEX=1"
    
    result = result & "'"
    
    GetWorkbookConnectionString = result
End Function

Public Function SanitizeDelimitedfieldNames(ByVal fieldNames As String) As String
    SanitizeDelimitedfieldNames = Replace(Replace(Replace(fieldNames, ",[]", vbNullString), ", []", vbNullString), "[]", vbNullString)
End Function


