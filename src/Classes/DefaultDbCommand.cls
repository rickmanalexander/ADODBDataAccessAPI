VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close."
'@Folder("ADODBDataAccess.DbCommand")
'@ModuleDescription("An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close.")
'@PredeclaredId
Option Explicit

Implements IDbCommand

Private Type TCommand
    Base As IDbCommandBase
    Connection As IDbConnection
End Type

Private this As TCommand

Public Function Create(ByVal db As IDbConnection, ByVal baseCommand As IDbCommandBase) As IDbCommand
    DbErrors.GuardNonDefaultInstance Me, DefaultDbCommand
    DbErrors.GuardNullReference db
    DbErrors.GuardNullReference baseCommand
    
    Dim result As DefaultDbCommand
    Set result = New DefaultDbCommand
    Set result.Connection = db
    Set result.Base = baseCommand
    
    Set Create = result
End Function

Friend Property Set Connection(ByVal value As IDbConnection)
    DbErrors.GuardDoubleInitialization this.Connection
    Set this.Connection = value
End Property

Friend Property Set Base(ByVal value As IDbCommandBase)
    Set this.Base = value
End Property

Private Sub GuardInvalidOrdinalArguments(ByVal sql As String, ByRef args() As Variant)
    DbErrors.GuardExpression Not this.Base.ValidateOrdinalArguments(sql, args), message:="Arguments provided are inconsistent with expected ordinal parameters."
End Sub

Private Function IDbCommand_ExecuteQuery(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
        
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, commandType, 0, sql, localArgs)
    
    Dim result As ADODB.Recordset
    Set result = this.Base.Execute(cmd)
    
    Set IDbCommand_ExecuteQuery = result
End Function

Private Function IDbCommand_ExecuteQueryAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
        
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, commandType, 0, sql, localArgs)
    
    Dim result As ADODB.Recordset
    Set result = this.Base.ExecuteAsync(cmd, True)
    
    Set IDbCommand_ExecuteQueryAsync = result
End Function

Private Sub IDbCommand_ExecuteNonQuery(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant)
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
        
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, commandType, 0, sql, localArgs)
    
    '@Ignore FunctionReturnValueDiscarded
    cmd.Execute Options:=(ExecuteOptionEnum.adExecuteNoRecords)
End Sub

Private Sub IDbCommand_ExecuteNonQueryAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant)
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
        
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, commandType, 0, sql, localArgs)
    
    '@Ignore FunctionReturnValueDiscarded
    cmd.Execute Options:=(ExecuteOptionEnum.adExecuteNoRecords + ExecuteOptionEnum.adAsyncExecute)
    
End Sub

Private Function IDbCommand_ExecuteWithParameters(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
    
    Dim emptyArgs() As Variant
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, commandType, 0, sql, emptyArgs)
    
    Dim i As Long
    For i = LBound(params) To UBound(params)
        DbErrors.GuardExpression (Not (TypeOf params(i) Is ADODB.Parameter)), "Parameters passed to 'ExecuteWithParameters' must be of Type 'ADODB.Parameter"
        cmd.parameters.Append params(i)
    Next
    
    Dim result As ADODB.Recordset
    Set result = this.Base.Execute(cmd)
    
    Set IDbCommand_ExecuteWithParameters = result
End Function

Private Function IDbCommand_ExecuteWithParametersAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
        
    Dim emptyArgs() As Variant
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, commandType, 0, sql, emptyArgs)
    
    Dim i As Long
    For i = LBound(params) To UBound(params)
        DbErrors.GuardExpression (Not (TypeOf params(i) Is ADODB.Parameter)), "Parameters passed to 'ExecuteWithParametersAsync' must be of Type 'ADODB.Parameter"
        cmd.parameters.Append params(i)
    Next
    
    Dim result As ADODB.Recordset
    Set result = this.Base.ExecuteAsync(cmd, True)
    
    Set IDbCommand_ExecuteWithParametersAsync = result
End Function

Private Function IDbCommand_ExecuteScalar(ByVal sql As String, ParamArray args() As Variant) As Variant
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
    
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    GuardInvalidOrdinalArguments sql, localArgs
    
    IDbCommand_ExecuteScalar = this.Base.ExecuteScalar(this.Connection, sql, localArgs)
End Function

Private Function IDbCommand_ExecuteScalarAsync(ByVal sql As String, ParamArray args() As Variant) As Variant
    DbErrors.GuardNullReference this.Connection
    DbErrors.GuardExpression (Not ((this.Connection.State And adStateOpen) = adStateOpen)), message:="Connection is not open."
    
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    GuardInvalidOrdinalArguments sql, localArgs
    
    IDbCommand_ExecuteScalarAsync = this.Base.ExecuteScalarAsync(this.Connection, sql, localArgs)
End Function


