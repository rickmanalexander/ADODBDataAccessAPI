VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates a database transaction."
'@Folder("ADODBDataAccess.UnitOfWork")
'@ModuleDescription("An object that encapsulates a database transaction.")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IUnitOfWork

Private Type TUnitOfWork
    Committed As Boolean
    RolledBack As Boolean
    Connection As IDbConnection
    CommandFactory As IDbCommandFactory
End Type

Private this As TUnitOfWork

'@Description("Creates a new unit of work using default configurations.")
Public Function FromConnectionString(ByVal connString As String) As IUnitOfWork
Attribute FromConnectionString.VB_Description = "Creates a new unit of work using default configurations."
    Dim db As IDbConnection
    Set db = DbConnection.Create(connString)
    
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(AdoTypeMappings.Default)
    
    Dim baseCommand As IDbCommandBase
    Set baseCommand = DbCommandBase.Create(provider)
    
    Dim factory As IDbCommandFactory
    Set factory = DefaultDbCommandFactory.Create(baseCommand)
    
    Set FromConnectionString = UnitOfWork.Create(db, factory)
End Function

Public Function Create(ByVal db As IDbConnection, ByVal factory As IDbCommandFactory) As IUnitOfWork
    DbErrors.GuardNonDefaultInstance Me, UnitOfWork
    DbErrors.GuardNullReference factory
    DbErrors.GuardNullReference db
    DbErrors.GuardExpression (Not ((db.State And adStateOpen) = adStateOpen)), message:="Connection should be open."
    
    Dim result As UnitOfWork
    Set result = New UnitOfWork
    Set result.CommandFactory = factory
    Set result.Connection = db
    
    Set Create = result
End Function

Friend Property Set Connection(ByVal value As IDbConnection)
    DbErrors.GuardDoubleInitialization this.Connection
    DbErrors.GuardNullReference value
    Set this.Connection = value
    this.Connection.BeginTransaction
End Property

Friend Property Set CommandFactory(ByVal value As IDbCommandFactory)
    DbErrors.GuardDoubleInitialization this.CommandFactory
    DbErrors.GuardNullReference value
    Set this.CommandFactory = value
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    If Not this.Committed Then this.Connection.RollbackTransaction
    On Error GoTo 0
End Sub

Private Sub IUnitOfWork_Commit()
    DbErrors.GuardExpression this.Committed, message:="Transaction is already committed."
    DbErrors.GuardExpression this.RolledBack, message:="Transaction was rolled back."
    On Error Resume Next ' not all providers support transactions
    this.Connection.CommitTransaction
    this.Committed = True
    On Error GoTo 0
End Sub

Private Function IUnitOfWork_Command() As IDbCommand
    Set IUnitOfWork_Command = this.CommandFactory.Create(this.Connection)
End Function

Private Sub IUnitOfWork_Rollback()
    DbErrors.GuardExpression this.Committed, message:="Transaction is already committed."
    On Error Resume Next 'not all providers support transactions
    this.Connection.RollbackTransaction
    this.RolledBack = True
    On Error GoTo 0
End Sub
