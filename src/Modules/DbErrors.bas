Attribute VB_Name = "DbErrors"
'@Folder("ADODBDataAccess")
'@ModuleDescription("Global procedures for throwing common DbErrors.")
Option Explicit

Public Const ADODBCustomError As Long = vbObjectError Or 32
Public Const ADODBCommandTimeoutError As Long = -2147217871

'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError(Optional ByVal adoError As ADODB.Error, Optional ByVal commandTimeout As Long)
    On Error GoTo CleanFail
    If Not adoError Is Nothing Then
        Dim errorDescr As String
        Dim timeMsg As String
        If adoError.Number = ADODBCommandTimeoutError Then
            timeMsg = IIf(commandTimeout <> 0, "longer than " & commandTimeout & " seconds. ", "too long. ")

            errorDescr = "The current command has timed out." & vbNewLine & vbNewLine & _
                         "This error occurs when server traffic is high, or when the " & _
                         "execution of a command takes " & timeMsg & _
                         "Please wait and try again later. " & _
                         "If the issue persists, then please contact the developer of this " & _
                         "project." & vbNewLine
        Else
            errorDescr = adoError.Description
        End If
        
        With VBA.Information.Err
            .Clear

            .Raise adoError.Number, adoError.Source, errorDescr, adoError.HelpFile, adoError.HelpContext
        End With
    Else
        With VBA.Information.Err
            If .Number <> 0 Then
                'Debug.Print "Error " & .Number, .Description
                .Raise .Number, .Source, .Description
            End If
        End With
    End If
    
CleanExit:
    Exit Sub
    
CleanFail:
    Resume CleanExit
End Sub

'@Description("Raises a run-time error if the specified Boolean expression is True.")
Public Sub GuardExpression(ByVal throw As Boolean, _
Optional ByVal errorSource As String = "AdoDbDataAccess.DbErrors", _
Optional ByVal message As String = "Invalid procedure call or argument.")
    If throw Then VBA.Information.Err.Raise ADODBCustomError, errorSource, message
End Sub

'@Description("Raises a run-time error if the specified instance isn't the default instance.")
Public Sub GuardNonDefaultInstance(ByVal instance As Object, ByVal defaultInstance As Object, _
Optional ByVal errorSource As String = "AdoDbDataAccess.DbErrors", _
Optional ByVal message As String = "Method should be invoked from the default/predeclared instance of this class.")
    Debug.Assert TypeName(instance) = TypeName(defaultInstance)
    GuardExpression Not instance Is defaultInstance, errorSource, message
End Sub

'@Description("Raises a run-time error if the specified object reference is already set.")
Public Sub GuardDoubleInitialization(ByVal instance As Object, _
Optional ByVal errorSource As String = "AdoDbDataAccess.DbErrors", _
Optional ByVal message As String = "Object is already initialized.")
    GuardExpression Not instance Is Nothing, errorSource, message
End Sub

'@Description("Raises a run-time error if the specified object reference is Nothing.")
Public Sub GuardNullReference(ByVal instance As Object, _
Optional ByVal errorSource As String = "AdoDbDataAccess.DbErrors", _
Optional ByVal message As String = "Object reference cannot be Nothing.")
    GuardExpression instance Is Nothing, errorSource, message
End Sub

'@Description("Raises a run-time error if the specified string is empty.")
Public Sub GuardEmptyString(ByVal value As String, _
Optional ByVal errorSource As String = "AdoDbDataAccess.DbErrors", _
Optional ByVal message As String = "String cannot be empty.")
    GuardExpression value = vbNullString, errorSource, message
End Sub
