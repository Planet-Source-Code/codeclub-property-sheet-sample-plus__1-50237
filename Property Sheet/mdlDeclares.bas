Attribute VB_Name = "mdlDeclares"
'*********************************************************************************************
'
' Edanmo's Shell Extensions - Property Sheet Handler
'
' Support functions
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

Type DLGTEMPLATE
    Style As Long
    dwExtendedStyle As Long
    cdit As Integer
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    Menu As Integer
    Class As String * 7
    Caption As Integer
End Type

' Window Styles
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000

' Messages
Public Const WM_SETFOCUS = &H7
Public Const WM_NOTIFY = &H4E
Public Const WM_INITDIALOG = &H110

Declare Function SetParentAPI Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Boolean

'  SetWindowPos Flags

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_CONTROLPARENT = &H10000

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const LB_SETTABSTOPS = &H192

Public Function AddrOf(ByVal Addr As Long) As Long

    AddrOf = Addr

End Function



Public Function DlgProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error Resume Next
    
    Select Case uMsg
    
        Case WM_INITDIALOG
   
            Load frmInfoPage
            
            ' Move the form to the
            ' property page
            SetParentAPI frmInfoPage.hwnd, hwnd

            ' Pass the page handle to
            ' the form
            frmInfoPage.m_lSheet = hwnd

            ' Pass the dialog handle to
            ' the from
            frmInfoPage.m_lPropSheet = GetParent(hwnd)

            ' Show the form
            SetWindowPos frmInfoPage.hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_SHOWWINDOW
            
            DlgProc = True

        Case WM_NOTIFY ' Handle property sheet notifications
            Dim PSPN As PSHNOTIFY

            MoveMemory PSPN, ByVal lParam, Len(PSPN)

            Select Case PSPN.Hdr.code

                Case PSN_SETACTIVE
                    DlgProc = frmInfoPage.SetActive

                Case PSN_APPLY
                    DlgProc = frmInfoPage.Apply

                Case PSN_RESET
                    DlgProc = frmInfoPage.Cancel

            End Select


    End Select

End Function

Public Function PSCallback(ByVal hwnd As Long, ByVal uMsg As Long, PSP As PROPSHEETPAGE) As Long
    
    Select Case uMsg
        Case PSPCB_CREATE
        
            ' return True to create the page
            ' or false to destroy it
            PSCallback = True
            
        Case PSPCB_RELEASE
        
            ' Release the control reference
            ' so the DLL can unload
            Set g_oHandlerRef = Nothing
            
    End Select

End Function






