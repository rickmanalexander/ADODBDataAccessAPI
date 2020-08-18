VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that deals with the database connection and builds and executes a parameterized ADODB command from a command string and parameter values. Returns disconnected recordsets."
'@Folder("ADODBDataAccess.DbCommand")
'@ModuleDescription("An object that deals with the database Connection and builds and executes a parameterized ADODB command from a command string and parameter values. Returns disconnected recordsets.")
'@PredeclaredId
Option Explicit

Implements IDbCommand

Private Type TCommand
    ConnectionString As String
    ConnectionFactory As DbConnectionFactory
    Base As IDbCommandBase
End Type

Private this As TCommand

Public Function Create(ByVal connString As String, ByVal factory As DbConnectionFactory, ByVal baseCommand As IDbCommandBase) As IDbCommand
    DbErrors.GuardNonDefaultInstance Me, AutoDbCommand
    DbErrors.GuardEmptyString connString
    DbErrors.GuardNullReference factory
    DbErrors.GuardNullReference baseCommand
    
    Dim result As AutoDbCommand
    Set result = New AutoDbCommand
    
    result.ConnectionString = connString
    Set result.ConnectionFactory = factory
    Set result.Base = baseCommand
    
    Set Create = result
End Function

Friend Property Let ConnectionString(ByVal value As String)
    DbErrors.GuardEmptyString value
    this.ConnectionString = value
End Property

Friend Property Set Base(ByVal value As IDbCommandBase)
    DbErrors.GuardDoubleInitialization this.Base
    Set this.Base = value
End Property

Friend Property Set ConnectionFactory(ByVal value As DbConnectionFactory)
    DbErrors.GuardDoubleInitialization this.ConnectionFactory
    Set this.ConnectionFactory = value
End Property

Private Sub GuardInvalidOrdinalArguments(ByVal sql As String, ByRef args() As Variant)
    DbErrors.GuardExpression Not this.Base.ValidateOrdinalArguments(sql, args), message:="Arguments provided are inconsistent with expected ordinal parameters."
End Sub

Private Function IDbCommand_ExecuteQuery(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    Dim localArgs() As Variant
    localArgs = args 'cannot pass ParamArray array directly to ByRef array argument
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    With this.ConnectionFactory.Create(this.ConnectionString)
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, commandType, 0, sql, localArgs)
        
        Dim result As ADODB.Recordset
        Set result = this.Base.Execute(cmd)
        
        Set result.ActiveConnection = Nothing
        
        Set IDbCommand_ExecuteQuery = result
    End With
End Function

Private Function IDbCommand_ExecuteQueryAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    Dim localArgs() As Variant
    localArgs = args 'cannot pass ParamArray array directly to ByRef array argument
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    With this.ConnectionFactory.Create(this.ConnectionString)
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, commandType, 0, sql, localArgs)
        
        Dim result As ADODB.Recordset
        Set result = this.Base.ExecuteAsync(cmd, True)
        
        Set result.ActiveConnection = Nothing
        
        Set IDbCommand_ExecuteQueryAsync = result
    End With
End Function

Private Sub IDbCommand_ExecuteNonQuery(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant)
    Dim localArgs() As Variant
    localArgs = args 'cannot pass ParamArray array directly to ByRef array argument
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    With this.ConnectionFactory.Create(this.ConnectionString)
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, commandType, 0, sql, localArgs)
        
        '@Ignore FunctionReturnValueDiscarded
        cmd.Execute Options:=(ExecuteOptionEnum.adExecuteNoRecords)
    End With
End Sub

Private Sub IDbCommand_ExecuteNonQueryAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant)
    Dim localArgs() As Variant
    localArgs = args 'cannot pass ParamArray array directly to ByRef array argument
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localArgs
    
    With this.ConnectionFactory.Create(this.ConnectionString)
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, commandType, 0, sql, localArgs)
        
        '@Ignore FunctionReturnValueDiscarded
        cmd.Execute Options:=(ExecuteOptionEnum.adExecuteNoRecords + ExecuteOptionEnum.adAsyncExecute)
    End With
End Sub

Private Function IDbCommand_ExecuteWithParameters(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    Dim localParams() As Variant
    localParams = params
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localParams

    With this.ConnectionFactory.Create(this.ConnectionString)
        Dim emptyArgs() As Variant
        
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, commandType, 0, sql, emptyArgs)
        
        Dim i As Long
        For i = LBound(params) To UBound(params)
            DbErrors.GuardExpression (Not (TypeOf params(i) Is ADODB.Parameter)), "Parameters passed to 'ExecuteWithParameters' must be of Type 'ADODB.Parameter"
            cmd.parameters.Append params(i)
        Next
        
        Dim result As ADODB.Recordset
        Set result = this.Base.Execute(cmd)
        
        Set result.ActiveConnection = Nothing
        
        Set IDbCommand_ExecuteWithParameters = result
    End With
End Function

Private Function IDbCommand_ExecuteWithParametersAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    Dim localParams() As Variant
    localParams = params
    If commandType = adCmdText Then GuardInvalidOrdinalArguments sql, localParams

    With this.ConnectionFactory.Create(this.ConnectionString)
        Dim emptyArgs() As Variant
        
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, commandType, 0, sql, emptyArgs)
        
        Dim i As Long
        For i = LBound(params) To UBound(params)
            DbErrors.GuardExpression (Not (TypeOf params(i) Is ADODB.Parameter)), "Parameters passed to 'ExecuteWithParametersAsync' must be of Type 'ADODB.Parameter"
            cmd.parameters.Append params(i)
        Next
        
        Dim result As ADODB.Recordset
        Set result = this.Base.ExecuteAsync(cmd, True)
        
        Set result.ActiveConnection = Nothing
        
        Set IDbCommand_ExecuteWithParametersAsync = result
    End With
End Function

Private Function IDbCommand_ExecuteScalar(ByVal sql As String, ParamArray args() As Variant) As Variant
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    GuardInvalidOrdinalArguments sql, localArgs
    
    With this.ConnectionFactory.Create(this.ConnectionString)
        IDbCommand_ExecuteScalar = this.Base.ExecuteScalar(.Self, sql, localArgs)
    End With
End Function

Private Function IDbCommand_ExecuteScalarAsync(ByVal sql As String, ParamArray args() As Variant) As Variant
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    GuardInvalidOrdinalArguments sql, localArgs
    
    With this.ConnectionFactory.Create(this.ConnectionString)
        IDbCommand_ExecuteScalarAsync = this.Base.ExecuteScalarAsync(.Self, sql, localArgs)
    End With
End Function