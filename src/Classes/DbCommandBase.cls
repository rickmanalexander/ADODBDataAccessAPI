VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object intended to be used as a private instance field in a class implementing the IDbCommand interface (composition in place of inheritance)."
'@Exposed
'@Folder("ADODBDataAccess.DbCommand.Abstract")
'@ModuleDescription("An object intended to be used as a private instance field in a class implementing the IDbCommand interface (composition in place of inheritance).")
'@PredeclaredId
Option Explicit

Implements IDbCommandBase

Private Type TCommand
    ParameterProvider As IParameterProvider
End Type

Private this As TCommand

Public Function Create(ByVal provider As IParameterProvider) As IDbCommandBase
    DbErrors.GuardNonDefaultInstance Me, DbCommandBase
    DbErrors.GuardNullReference provider
    
    Dim result As DbCommandBase
    Set result = New DbCommandBase
    Set result.ParameterProvider = provider
    
    Set Create = result
End Function

'@Inject: this member should only be invoked by Me.Create, where Me is the class' default/predeclared instance.
'@Ignore ProcedureNotUsed: false positive with v2.5.0.5418
Friend Property Set ParameterProvider(ByVal value As IParameterProvider)
    DbErrors.GuardDoubleInitialization this.ParameterProvider
    Set this.ParameterProvider = value
End Property

Private Function CreateCommand(ByVal db As IDbConnection, ByVal commandType As ADODB.CommandTypeEnum, ByVal commandTimeout As Long, ByVal sql As String, ByRef args() As Variant) As ADODB.Command
    DbErrors.GuardNullReference db
    DbErrors.GuardEmptyString sql
    DbErrors.GuardExpression (Not ((db.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
    If commandType = adCmdText Then DbErrors.GuardExpression Not ValidateOrdinalArguments(sql, args), message:="Arguments supplied are inconsistent with the provided command string parameters."

    Dim cmd As ADODB.Command
    Set cmd = db.CreateCommand(commandType, commandTimeout, sql)
    
    On Error GoTo CleanFail
    Dim arg As ADODB.Parameter
    For Each arg In this.ParameterProvider.FromValues(args)
        cmd.parameters.Append arg
    Next
    
CleanExit:
    Set CreateCommand = cmd
    Exit Function
    
CleanFail:
    Resume CleanExit
End Function

Public Function ValidateOrdinalArguments(ByVal sql As String, ByRef args() As Variant) As Boolean
    On Error GoTo CleanFail
    Dim result As Boolean
    
    Dim expected As Long
    expected = Len(sql) - Len(Replace(sql, "?", vbNullString))
    
    Dim actual As Long
    On Error GoTo CleanFail 'if there are no args, LBound/UBound are both out of bounds
    actual = UBound(args) + (1 - LBound(args))
    
CleanExit:
    result = (expected = actual)
    ValidateOrdinalArguments = result
    Exit Function
    
CleanFail:
    actual = 0
    Resume CleanExit
End Function

Private Sub AsycAwait(ByRef cmd As ADODB.Command)
    Do While ((cmd.State And ADODB.ObjectStateEnum.adStateExecuting) = ADODB.ObjectStateEnum.adStateExecuting)
    Loop
End Sub

Private Function Execute(ByVal cmd As ADODB.Command) As ADODB.Recordset
    DbErrors.GuardNullReference cmd
    DbErrors.GuardNullReference cmd.ActiveConnection

    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    result.CursorLocation = adUseClient
    
    result.Open Source:=cmd, CursorType:=adOpenStatic
    
    Set Execute = result
End Function

Private Function ExecuteAsync(ByVal cmd As ADODB.Command, ByVal awaitResults As Boolean) As ADODB.Recordset
    DbErrors.GuardNullReference cmd
    DbErrors.GuardNullReference cmd.ActiveConnection

    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    result.CursorLocation = adUseClient
    
    result.Open Source:=cmd, CursorType:=adOpenStatic, Options:=(ExecuteOptionEnum.adAsyncExecute)
    
    If awaitResults Then AsycAwait cmd
    
    Set ExecuteAsync = result
End Function

Private Function ExecuteScalar(ByVal db As IDbConnection, ByVal sql As String, ByRef args() As Variant) As Variant
    DbErrors.GuardEmptyString sql
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommand(db, adCmdText, 0, sql, args)
    
    Dim rs As ADODB.Recordset
    Set rs = Execute(cmd)
    
    Dim result As Variant
    result = rs.Fields.Item(0).value
    
    Set rs.ActiveConnection = Nothing
    
    ExecuteScalar = result
End Function

Private Function ExecuteScalarAsync(ByVal db As IDbConnection, ByVal sql As String, ByRef args() As Variant) As Variant
    DbErrors.GuardEmptyString sql
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommand(db, adCmdText, 0, sql, args)
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteAsync(cmd, True)
    
    Dim result As Variant
    result = rs.Fields.Item(0).value
    
    Set rs.ActiveConnection = Nothing
    
    ExecuteScalarAsync = result
End Function

Private Function IDbCommandBase_CreateCommand(ByVal db As IDbConnection, ByVal commandType As ADODB.CommandTypeEnum, ByVal commandTimeout As Long, ByVal sql As String, ByRef args() As Variant) As ADODB.Command
    Set IDbCommandBase_CreateCommand = CreateCommand(db, commandType, commandTimeout, sql, args)
End Function

Private Function IDbCommandBase_Execute(ByVal cmd As ADODB.Command) As ADODB.Recordset
    Set IDbCommandBase_Execute = Execute(cmd)
End Function

Private Function IDbCommandBase_ExecuteAsync(ByVal cmd As ADODB.Command, ByVal awaitResults As Boolean) As ADODB.Recordset
    Set IDbCommandBase_ExecuteAsync = ExecuteAsync(cmd, awaitResults)
End Function

Private Function IDbCommandBase_ExecuteScalar(ByVal db As IDbConnection, ByVal sql As String, ByRef args() As Variant) As Variant
    IDbCommandBase_ExecuteScalar = ExecuteScalar(db, sql, args)
End Function

Private Function IDbCommandBase_ExecuteScalarAsync(ByVal db As IDbConnection, ByVal sql As String, ByRef args() As Variant) As Variant
    IDbCommandBase_ExecuteScalarAsync = ExecuteScalarAsync(db, sql, args)
End Function

Private Function IDbCommandBase_ValidateOrdinalArguments(ByVal sql As String, args() As Variant) As Boolean
    IDbCommandBase_ValidateOrdinalArguments = ValidateOrdinalArguments(sql, args)
End Function
