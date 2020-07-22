VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An implementation of an abstract factory that creates DbConnection objects."
'@Exposed
'@Folder("ADODBDataAccess.DbConnection")
'@ModuleDescription("An implementation of an abstract factory that creates DbConnection objects.")
Option Explicit

Public Function Create(ByVal connString As String) As IDbConnection
    Set Create = DbConnection.Create(connString)
End Function
