Attribute VB_Name = "mdlGlobalData"
'*********************************************************************************************
'
' Edanmo's Shell Extensions - Property Sheet Handler
'
' Global variables
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Use at your own risk.
'
' Tested with:
'              * Windows Me / Windows XP
'              * VB6 SP5
'
' History:
'           08/24/1999 - This code was released
'
'*********************************************************************************************'*********************************************************************************************
Option Explicit

' See readme.txt
Declare Function AddPage Lib "psadd.dll" (ByVal lpfnAddPage As Long, ByVal hPage As Long, ByVal lParam As Long) As Long

' Control reference
Public g_oHandlerRef As PSHandler

' Dialog template
Public g_tDlgTemplate As DLGTEMPLATE

' Array of selected filenames
Public g_asSelectedFiles() As String



Sub Main()
    
    With New cRegister
        .RegisterPPFileType "frm"
    End With
    
    'RegisterPPAll
    'to unregister use Unreg App.Path & "\" & App.EXEName & ".dll",False
    'Unreg App.Path & "\" & App.EXEName & ".dll", False
    
End Sub


