VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbConnectionObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Sub OnConnecting(connString As String, UserID As String, Password As String, configuration As Long, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
End Sub

Public Sub OnConnectionComplete(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
End Sub

Private Sub OnDisconnecting(eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
End Sub


Public Sub OnInitiatingTransaction(ByVal level As Long, ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
End Sub

Public Sub OnCommittedTransaction(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
End Sub

Private Sub OnRolledbackTransaction(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
End Sub


Private Sub OnExecuting(errorSource As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, configuration As Long, eventStatus As ADODB.EventStatusEnum, ByVal cmd As ADODB.Command, ByVal results As ADODB.Recordset, ByVal db As ADODB.Connection)
End Sub

Private Sub OnExecutionComplete(ByVal RecordsAffected As Long, ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal cmd As ADODB.Command, ByVal results As ADODB.Recordset, ByVal db As ADODB.Connection)
End Sub

Private Sub OnInfoMessageRecieved(ByVal dbError As ADODB.Error, eventStatus As ADODB.EventStatusEnum, ByVal db As ADODB.Connection)
End Sub
