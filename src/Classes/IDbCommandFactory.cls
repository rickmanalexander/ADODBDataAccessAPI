VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An abstract factory that creates IDbCommand objects."
'@Exposed
'@Folder("ADODBDataAccess.DbCommand.Abstract")
'@ModuleDescription("An abstract factory that creates IDbCommand objects.")
'@Interface
Option Explicit

'@Description("Creates a new IDbCommand object.")
Public Function Create(ByVal db As IDbConnection) As IDbCommand
Attribute Create.VB_Description = "Creates a new IDbCommand object."
End Function

Private Sub Class_Initialize()
    Err.Raise 5, VBA.TypeName(Me), "An Interface must not be instantiated."
End Sub