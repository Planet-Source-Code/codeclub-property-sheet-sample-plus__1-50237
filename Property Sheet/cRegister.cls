VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private strTemp As String
Private strTemp2 As String



Public Function UnReg(ByVal inFileSpec As String, Optional Register As Boolean = True) As Boolean
    
    Dim lLib As Long
    Dim lpDLLEntryPoint As Long
    Dim lpThreadID As Long
    Dim lpExitCode As Long
    Dim mThread
    Dim mResult As Long
    
    On Error Resume Next
    
        lLib = LoadLibrary(inFileSpec)
        If lLib = 0 Then
            UnReg = False
            Exit Function
        End If
        
        If Register = True Then
            lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
        Else
            lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
        End If
        If lpDLLEntryPoint = vbNull Then
            GoTo handleError1
        End If
        
        mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
        If mThread = 0 Then
            GoTo handleError1
        End If
        
        mResult = WaitForSingleObject(mThread, 10000)
        If mResult <> 0 Then
            GoTo handleError2
        End If
    
        CloseHandle mThread
        FreeLibrary lLib
        
        UnReg = True
        
    Exit Function
    
handleError1:
    FreeLibrary lLib
    UnReg = False
    Exit Function
    
handleError2:
    FreeLibrary lLib
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode
    UnReg = False
End Function

Public Function RegisterPPAll() As Boolean
'********Register property sheet handler********
'Call this function if you want the property pages to be
'available for all files and folders
'************************************************
        'strTemp = GetSettingString(HKEY_CLASSES_ROOT, "PropertyPages.PSHandler\Clsid", "")
        'for using another different page add procedures in
        'modDeclares, add another form, and another class
        'use the classname to get is CLSID. for eg
        'strTemp = GetSettingString(HKEY_CLASSES_ROOT, "PropertyPages.PSHandler2\Clsid", "")
            
            'SaveSettingString HKEY_CLASSES_ROOT, "*\shellex\PropertySheetHandlers\SDRE", "", strTemp
            'SaveSettingString HKEY_CLASSES_ROOT, "Directory\shellex\PropertySheetHandlers\SDRE", "", strTemp
    
    Dim blnRet As Boolean
    
    With New cRegistry
        .ClassKey = HKEY_CLASSES_ROOT
        .SectionKey = "PropertyPages.PSHandler\Clsid"
        .Value = ""
        strTemp = .Value

        If strTemp <> "" Then
            .SectionKey = "*\shellex\PropertySheetHandlers\SDRE"
            .ValueKey = ""
            .Value = strTemp
            
            .SectionKey = "Directory\shellex\PropertySheetHandlers\SDRE"
            .ValueKey = ""
            .Value = strTemp
            
            blnRet = True
        Else
            blnRet = False
        End If
    End With
    
    RegisterPPAll = blnRet
    
End Function

Public Function RegisterPPDirectoriesOnly() As Boolean
'********Register property sheet handler********
'call this if property pages are to be available for
'folders alone
'************************************************
    
    'strTemp = GetSettingString(HKEY_CLASSES_ROOT, "PropertyPages.PSHandler\Clsid", "")
    'If strTemp <> "" Then
    '    SaveSettingString HKEY_CLASSES_ROOT, "Directory\shellex\PropertySheetHandlers\SDRE", "", strTemp
    'Else
    'End If
    
    Dim blnRet As Boolean
    
    With New cRegistry
        .ClassKey = HKEY_CLASSES_ROOT
        .SectionKey = "PropertyPages.PSHandler\Clsid"
        .ValueKey = ""
        strTemp = .Value
        
        If strTemp <> "" Then
            .SectionKey = "Directory\shellex\PropertySheetHandlers\SDRE"
            .ValueKey = ""
            .Value = strTemp
            
            blnRet = True
        Else
            blnRet = False
        End If
                
    End With
    
    RegisterPPDirectoriesOnly = blnRet
    
End Function

Public Function RegisterPPAllFileTypes() As Boolean
'********Register property sheet handler********
'call this if property pages are to be available for
'files alone
'************************************************
    
'    strTemp = GetSettingString(HKEY_CLASSES_ROOT, "PropertyPages.PSHandler\Clsid", "")
'    If strTemp <> "" Then
'        SaveSettingString HKEY_CLASSES_ROOT, "*\shellex\PropertySheetHandlers\SDRE", "", strTemp
'    Else
'        MsgBox "not done"
'    End If
    
    Dim blnRet As Boolean
    
    With New cRegistry
        .ClassKey = HKEY_CLASSES_ROOT
        .SectionKey = "PropertyPages.PSHandler\Clsid"
        .ValueKey = ""
        strTemp = .Value
        
        If strTemp <> "" Then
            .SectionKey = "*\shellex\PropertySheetHandlers\SDRE"
            .ValueKey = ""
            .Value = strTemp
            
            blnRet = True
        Else
            blnRet = False
        End If
                
    End With
    
    RegisterPPAllFileTypes = blnRet
        
End Function

Public Function RegisterPPFileType(ByVal Extension As String) As Boolean
'********Register property sheet handler********
'call this if property pages are to be available for
'some specific file types only
'************************************************

'    strTemp = GetSettingString(HKEY_CLASSES_ROOT, "PropertyPages.PSHandler\Clsid", "")
'    strTemp2 = GetSettingString(HKEY_CLASSES_ROOT, "." & Extension, "")
'    If strTemp <> "" Then
'        SaveSettingString HKEY_CLASSES_ROOT, strTemp2 & "\shellex\PropertySheetHandlers\SDRE", "", strTemp
'    Else
'        MsgBox "not done"
'    End If
'
    Dim blnRet As Boolean
    
    With New cRegistry
        .ClassKey = HKEY_CLASSES_ROOT
        .SectionKey = "PropertyPages.PSHandler\Clsid"
        .ValueKey = ""
        strTemp = .Value
        
        .SectionKey = "." & Extension
        .ValueKey = ""
        strTemp2 = .Value
        
        If strTemp <> "" Then
            .SectionKey = strTemp2 & "\shellex\PropertySheetHandlers\SDRE"
            .ValueKey = ""
            .Value = strTemp
            
            blnRet = True
        Else
            blnRet = False
        End If
                
    End With
    
    RegisterPPFileType = blnRet
        
End Function
