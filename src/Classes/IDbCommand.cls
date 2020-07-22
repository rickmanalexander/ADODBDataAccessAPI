VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object abstracting a database command."
'@Folder("ADODBDataAccess.DbCommand.Abstract")
'@ModuleDescription("Represents an object abstracting a database command.")
'@Interface
'@Exposed
Option Explicit

'@Description("Returns an ADODB.Recordset that contains the results of the specified parameterized query.")
Public Function ExecuteQuery(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
End Function

'@Description("Returns an ADODB.Recordset that contains the results of the specified parameterized query Asynchronously.")
Public Function ExecuteQueryAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
End Function


'@Description("Executes the specified parameterized SQL statement/command.")
Public Sub ExecuteNonQuery(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant)
End Sub

'@Description("Executes the specified parameterized SQL statement/command Asynchronously.")
Public Sub ExecuteNonQueryAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray args() As Variant)
End Sub


'@Description("Returns an ADODB.Recordset that contains the results of the specified parameterized query.")
Public Function ExecuteWithParameters(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
End Function

'@Description("Returns an ADODB.Recordset that contains the results of the specified parameterized query Asynchronously.")
Public Function ExecuteWithParametersAsync(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
End Function


'@Description("Returns the value of the first field of the first record returned by the specified parameterized query.")
Public Function ExecuteScalar(ByVal sql As String, ParamArray args() As Variant) As Variant
End Function

'@Description("Returns the value of the first field of the first record returned by the specified parameterized query Asynchronously.")
Public Function ExecuteScalarAsync(ByVal sql As String, ParamArray args() As Variant) As Variant
End Function
