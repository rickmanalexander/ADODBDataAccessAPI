VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Adapter object for client to specify actions to execute based on events raised by an IDbConnection."

Option Explicit

Public Event Connecting(connString As String, UserID As String, Password As String, configuration As Long, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
Public Event ConnectionComplete(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
Public Event Disconnecting(eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
Public Event InitiatingTransaction(ByVal level As Long, ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
Public Event CommittedTransaction(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
Public Event RolledbackTransaction(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
Public Event Executing(errorSource As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, configuration As Long, eventStatus As ADODB.EventStatusEnum, ByVal cmd As ADODB.Command, ByVal results As ADODB.Recordset, ByVal db As ADODB.Connection)
Public Event ExecutionComplete(ByVal RecordsAffected As Long, ByVal adoError As ADODB.Error, ByVal eventStatus As ADODB.EventStatusEnum, ByVal cmd As ADODB.Command, ByVal results As ADODB.Recordset, ByVal db As ADODB.Connection)
Public Event InfoMessageRecieved(ByVal adoError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)

Implements IDbConnection
Implements IDbConnectionObserver

Private Type TDbConnectionAdapter
    Connection As IDbConnection
End Type

Private this As TDbConnectionAdapter

Public Function Create(ByVal db As IDbConnection) As DbConnectionAdapter
    Dim result As DbConnectionAdapter
    Set result = New DbConnectionAdapter
    Set result.Connection = db
    
    Set Create = result
End Function


Friend Property Set Connection(ByVal value As IDbConnection)
    DbErrors.GuardDoubleInitialization this.Connection
    Set this.Connection = value
End Property


'IDbConnection
Private Property Get IDbConnection_AdoConnection() As ADODB.Connection
    Set IDbConnection_AdoConnection = this.Connection.AdoConnection
End Property

Private Sub IDbConnection_BeginTransaction()
        this.Connection.BeginTransaction
End Sub

Private Sub IDbConnection_CommitTransaction()
        this.Connection.CommitTransaction
End Sub

Private Function IDbConnection_CreateCommand(ByVal commandType As ADODB.CommandTypeEnum, ByVal commandTimeout As Long, ByVal sql As String) As ADODB.Command
    Set IDbConnection_CreateCommand = this.Connection.CreateCommand(commandType, commandTimeout, sql)
End Function

Private Sub IDbConnection_RollbackTransaction()
    this.Connection.RollbackTransaction
End Sub

Private Property Get IDbConnection_Self() As IDbConnection
    Set IDbConnection_Self = this.Connection
End Property

Private Property Get IDbConnection_State() As ADODB.ObjectStateEnum
    IDbConnection_State = this.Connection.State
End Property



'IDbConnectionObserver
Private Sub IDbConnectionObserver_OnConnecting(connString As String, UserID As String, Password As String, configuration As Long, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
        RaiseEvent Connecting(connString, UserID, Password, configuration, eventStatus, db)
End Sub

Public Sub IDbConnectionObserver_OnConnectionComplete(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
        RaiseEvent ConnectionComplete(dbError, eventStatus, db)
End Sub

Private Sub IDbConnectionObserver_OnDisconnecting(eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
        RaiseEvent Disconnecting(eventStatus, db)
End Sub

Public Sub IDbConnectionObserver_OnInitiatingTransaction(ByVal level As Long, ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
        RaiseEvent InitiatingTransaction(level, dbError, eventStatus, db)
End Sub

Public Sub IDbConnectionObserver_OnCommittedTransaction(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
        RaiseEvent CommittedTransaction(dbError, eventStatus, db)
End Sub

Private Sub IDbConnectionObserver_OnRolledbackTransaction(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
        RaiseEvent RolledbackTransaction(dbError, eventStatus, db)
End Sub

Private Sub IDbConnectionObserver_OnExecuting(errorSource As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, configuration As Long, eventStatus As ADODB.EventStatusEnum, ByVal cmd As ADODB.Command, ByVal results As ADODB.Recordset, ByVal db As ADODB.Connection)
        RaiseEvent Executing(errorSource, CursorType, LockType, configuration, eventStatus, cmd, results, db)
End Sub

Private Sub IDbConnectionObserver_OnExecutionComplete(ByVal RecordsAffected As Long, ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal cmd As ADODB.Command, ByVal results As ADODB.Recordset, ByVal db As ADODB.Connection)
        RaiseEvent ExecutionComplete(RecordsAffected, dbError, eventStatus, cmd, results, db)
End Sub

Private Sub IDbConnectionObserver_OnInfoMessageRecieved(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
        RaiseEvent InfoMessageRecieved(dbError, eventStatus, db)
End Sub
