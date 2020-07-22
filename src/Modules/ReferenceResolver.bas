Attribute VB_Name = "ReferenceResolver"
'@Folder("ADODBDataAccess")
'@ModuleDescription("Internal methods used to resolve commonly used references in a VBproject.")
Option Explicit
Option Private Module

Public Enum CommonDllVbProjectReference
    AdoDbRef = 2 ^ 1
    AdoDDlExtRef = 2 ^ 2
    ScriptingRuntimeRef = 2 ^ 3
    VbScriptRegExpRef = 2 ^ 4
    VbaExtensibilityRef = 2 ^ 5
    MscoreeRef = 2 ^ 6
    MsCorLibRef = 2 ^ 7
    MsCommonControlsRef = 2 ^ 8
End Enum

Public Const REFERENCE_EXISTS_ERROR_NUMBER As Long = 32813
    
Private Const DEBUG_REFERENCE_SEPARATOR As String = "--------------------------------------------------------------------------" & vbNewLine

Public Function TryAddDllReferences(ByVal dllReference As CommonDllVbProjectReference) As Boolean
    Dim result As Boolean
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.AdoDbRef) _
    Then
        result = TryAddDllReference("{B691E011-1797-432E-907A-4D8C69339129}", 6, 1)
    End If
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.AdoDDlExtRef) _
    Then
        result = TryAddDllReference("{00000600-0000-0010-8000-00AA006D2EA4}", 6, 0)
    End If
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.ScriptingRuntimeRef) _
    Then
        result = TryAddDllReference("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
    End If
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.VbScriptRegExpRef) _
    Then
        result = TryAddDllReference("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5)
    End If
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.VbaExtensibilityRef) _
    Then
        result = TryAddDllReference("{0002E157-0000-0000-C000-000000000046}", 5, 3)
    End If
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.MscoreeRef) _
    Then
        result = TryAddDllReference("{5477469E-83B1-11D2-8B49-00A0C9B7C9C4}", 2, 4)
    End If
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.MsCorLibRef) _
    Then
        result = TryAddDllReference("{BED7F4EA-1A96-11D2-8F08-00A0C9A6186D}", 2, 4)
    End If
    
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.MsCommonControlsRef) _
    Then
        result = TryAddDllReference("{A0518CD7-EE94-40FE-BA50-AAF72E4F4410}", 2, 2)
    End If
        
    TryAddDllReferences = result
End Function

Public Function TryAddDllReference(ByVal guid As String, ByVal majorVersion As Long, ByVal minorVersion As Long) As Boolean
    On Error Resume Next
    ThisWorkbook.VBProject.References.AddFromGuid guid, majorVersion, minorVersion
    TryAddDllReference = ((Err.Number = REFERENCE_EXISTS_ERROR_NUMBER) Or (Err.Number = 0))
    On Error GoTo 0
End Function

Public Sub DisplayReferenceError(ByVal developerName As String, ByVal developerEmailAddress As String)
    MsgBox "An error occured while attempting to add External reference(s) required " & _
           "for this Application." & vbNewLine & vbNewLine & _
           "Please contact " & developerName & " at " & developerEmailAddress, _
            vbCritical, vbNullString
End Sub

Private Function EnumHasFlag(ByVal flagsOrDefault As Long, ByVal searchFlag As Long) As Boolean
    EnumHasFlag = ((flagsOrDefault And searchFlag) = searchFlag)
End Function

Public Sub PrintAllReferences(ByVal project As Object)
    Dim ref As Object
    For Each ref In project.References
        If Not ref.BuiltIn Then
            PrintToImmediateWindow ref

        End If
    Next
End Sub

Public Sub PrintToImmediateWindow(ByVal reference As Object)
    Debug.Print "'" + DEBUG_REFERENCE_SEPARATOR
    Debug.Print "' GUID:                  " + reference.guid
    Debug.Print "' Major Version Number:  " + CStr(reference.Major)
    Debug.Print "' Minor Version Number:  " + CStr(reference.Minor)
    Debug.Print "' FullPath:              " + reference.fullPath
    Debug.Print "' Name:                  " + reference.name
    Debug.Print "' Description:           " + reference.Description
    Debug.Print "' BuiltIn:               " + CStr(reference.BuiltIn)
    Debug.Print "'" + DEBUG_REFERENCE_SEPARATOR
End Sub
