Attribute VB_Name = "ApplicationConfig"
'@Folder("ADODBDataAccess")
'@ModuleDescription("Internal methods used to configure this addin on startup.")
Option Explicit
Option Private Module

Private Const APPLICATION_NAME As String = "ADODBDataAccessAPI"
Private Const WORKBOOK_NAME As String = APPLICATION_NAME & ".xlam"

Private Const DEVELOPER_NAME As String = "Alexander Rickman"
Private Const DEVELOPER_USERNAME As String = "ARICKMAN"
Private Const DEVELOPER_EMAIL As String = "xxxxxxxxx.xxxxxxx@xxxxxxxx.com"

'*******************************************************************************
'allows developer to save if the shift is pressed
#If VBA7 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

Private Const KEY_MASK As Integer = &HFF80
Private Const SHIFT_KEY = &H10
Private Const CTRL_KEY = &H11
Private Const ALT_KEY = &H12
'*******************************************************************************

Public Sub OnWorkBookOpen()
    On Error GoTo CleanFail
    If Not IsApplicationDeveloper(DEVELOPER_USERNAME) Then
        If Not IsValidApplicationFileName(ThisWorkbook.name, WORKBOOK_NAME) Then ExitApp
                                                
        If Not ReferenceResolver.TryAddDllReferences(dllReference:=CommonDllVbProjectReference.AdoDbRef + CommonDllVbProjectReference.ScriptingRuntimeRef) _
        Then
            ReferenceResolver.DisplayReferenceError DEVELOPER_NAME, DEVELOPER_EMAIL
            
            ExitApp
        End If
    End If
    
CleanExit:
    Exit Sub

CleanFail:
    ManageApplicationStartupError
    
    Resume CleanExit
End Sub
                
Public Function IsValidApplicationFileName(ByVal currentApplicationFileName As String, ByVal expectedApplicationFileName As String) As Boolean
    Dim result As Boolean
    result = False
                            
    If (currentApplicationFileName = expectedApplicationFileName) Then
            result = True
    Else
                    MsgBox "Looks like this application's original name has been changed. " & _
                                             vbNewLine & vbNewLine & _
                                             "To use this application, its name must remain as the following:" & _
                                             vbNewLine & vbNewLine & _
                                    expectedApplicationFileName & vbNewLine & vbNewLine & _
                                             "Clicking 'Okay' will automatically exit this application. " & _
                                             "Once closed, you MUST restore the name " & _
                                             "to the one mentioned above.", _
                                             vbCritical, "Error: Unauthorized Name Change"
    End If
                            
    IsValidApplicationFileName = result
End Function

Public Function AllowWorkbookSave(Optional ByVal warningMessage As String = "In order to preserve the structure of this application, " & _
                                                                                "saving has been disabled.") As Boolean
    Dim result As Boolean
    result = False
    
    'only a user to save if they are holding down shift AND they are the admin
    If CBool(GetKeyState(SHIFT_KEY) And KEY_MASK) Then
        If UCase$(Trim$(Environ$("USERNAME"))) = DEVELOPER_USERNAME Then
            result = True
        Else
            MsgBox warningMessage, vbExclamation + vbOKOnly, vbNullString
        End If
    Else
        MsgBox warningMessage, vbExclamation + vbOKOnly, vbNullString
    End If
    
    AllowWorkbookSave = result
End Function

Public Sub ExitApp()
    If Application.Workbooks.Count = 1 Then
        Application.EnableEvents = False    'this will reset itself when Excel is opened again
        Application.Quit
    Else
        ThisWorkbook.Saved = True
        ThisWorkbook.Close
    End If
End Sub

Public Function IsApplicationDeveloper(ByVal expectedDeveloperUserName As String) As Boolean
    IsApplicationDeveloper = (UCase$(Trim$(Environ$("USERNAME"))) = UCase$(Trim$(expectedDeveloperUserName)))
End Function

Private Sub ManageApplicationStartupError()
    MsgBox "Application StartUp Error" & vbNewLine & vbNewLine & _
           "An error occured while this application was attempting to load. " & _
           "If this issue persists, please contact the developer of this project. " & vbNewLine & vbNewLine & _
           "This application will now exit.", vbCritical, vbNullString
    
    ExitApp
End Sub
