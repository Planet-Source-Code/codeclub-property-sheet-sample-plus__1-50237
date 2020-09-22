VERSION 5.00
Begin VB.Form frmInfoPage 
   BorderStyle     =   0  'None
   ClientHeight    =   4695
   ClientLeft      =   3795
   ClientTop       =   3045
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstHeader 
      Height          =   1140
      Left            =   1035
      TabIndex        =   0
      Top             =   3420
      Width           =   4005
   End
   Begin VB.ListBox lstCtrls 
      Height          =   1140
      Left            =   1035
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2130
      Width           =   4005
   End
   Begin VB.ListBox lstOCXs 
      Columns         =   2
      Height          =   1140
      Left            =   1035
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   4005
   End
   Begin VB.TextBox txtCaption 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   585
      Width           =   4005
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   315
      Width           =   4005
   End
   Begin VB.TextBox txtVer 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   45
      Width           =   4005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Header:"
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   3420
      Width           =   570
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controls:"
      Height          =   195
      Left            =   375
      TabIndex        =   3
      Top             =   2130
      Width           =   615
   End
   Begin VB.Label lblVer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&VB Version:"
      Height          =   195
      Left            =   165
      TabIndex        =   11
      Top             =   45
      Width           =   825
   End
   Begin VB.Label lblCapt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Form caption:"
      Height          =   195
      Left            =   30
      TabIndex        =   7
      Top             =   585
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Name:"
      Height          =   195
      Left            =   525
      TabIndex        =   9
      Top             =   315
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OCXs:"
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   840
      Width           =   450
   End
End
Attribute VB_Name = "frmInfoPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Edanmo's Shell Extensions - Property Sheet Handler
'
' Property Page form
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
'           02/14/2000 - The form can be navigated
'                        with the keyboard.
'           08/24/1999 - This code was released
'
'*********************************************************************************************'*********************************************************************************************
'
' Notes: The TabIndex property of control does
'        not work since the property sheet is
'        changing the focus instead of VB. The
'        tab order is given by the z-order.
'        The control at the bottom of the z-order
'        is the first control in the tab order.
'
'*********************************************************************************************
Option Explicit

Public m_lPropSheet As Long
Public m_lSheet As Long
'*********************************************************************************************
' This function is called from DlgProc when
' the user clicks on the Apply or OK button
'*********************************************************************************************
Public Function Apply() As PSNOTIFYRESULTS

   ' This page doesn't allow changes.
   ' If your page allows the user to make
   ' changes you should save them here.
   
End Function

'*********************************************************************************************
' This function called is from DlgProc when
' the user clicks on Cancel button
'*********************************************************************************************
Public Function Cancel() As PSNOTIFYRESULTS

End Function

'*********************************************************************************************
' Notifies the dialog about a
' change in the data. The properties
' dialog will enable the Apply button.
'
' This Sub is not used in this sample.
'
'*********************************************************************************************
Private Sub Changed(Optional ByVal EnableApply As Boolean = True)
    
    If EnableApply Then
        SendMessage m_lPropSheet, PSM_CHANGED, m_lSheet, ByVal 0&
    Else
        SendMessage m_lPropSheet, PSM_UNCHANGED, m_lSheet, ByVal 0&
    End If
    
End Sub

Private Sub LoadInfo()

    Dim sLine As String
    Dim lPos As Long
    
    On Error Resume Next
    
    ' Read the FRM file and
    ' extract the info
    
    Open g_asSelectedFiles(0) For Input As #1
                 
    Line Input #1, sLine
    txtVer.Text = Mid$(sLine, 9)
    lstHeader.AddItem sLine
    
    ' Fill OCXs list
    lstOCXs.Clear
    Do While InStr(sLine, "Begin") = 0
                
        If Left$(sLine, 6) = "Object" Then
        
            sLine = Trim$(Mid$(sLine, InStr(sLine, ";") + 1))
            
            lstOCXs.AddItem Mid$(sLine, 2, Len(sLine) - 2)
            
        End If
        
        Line Input #1, sLine
        
        lstHeader.AddItem sLine
        
    Loop
    
    ' Get the form name
    sLine = Trim$(Mid$(sLine, InStr(sLine, "Form") + 4))
    txtName.Text = sLine
    
    ' Search form caption
    Do While Not EOF(1) And InStr(sLine, "Begin ") = 0
        
        Line Input #1, sLine
        
        lstHeader.AddItem sLine
        
        If InStr(sLine, "Caption") Then
            
            sLine = Trim$(Mid$(sLine, InStr(sLine, "=") + 1))
            sLine = Mid$(sLine, 2, Len(sLine) - 2)
            
            txtCaption.Text = sLine
            Exit Do
            
        End If
        
    Loop
    
    ' Fill control list
    lstCtrls.Clear
    Do While Not EOF(1) And Left$(sLine, 3) <> "End"

        If InStr(sLine, "Begin ") Then

            sLine = Trim$(Mid$(sLine, InStr(sLine, "Begin ") + 6))
            lPos = InStr(sLine, " ")

            lstCtrls.AddItem Mid$(sLine, lPos + 1) & vbTab & Left$(sLine, lPos - 1)

        End If

        Line Input #1, sLine
        
        lstHeader.AddItem sLine

    Loop
        
    Close

End Sub

'*********************************************************************************************
' This function is called from DlgProc when
' the page is being activated
'*********************************************************************************************
Public Function SetActive() As PSNOTIFYRESULTS
    
    SetActive = PSNRET_NOERROR
    
End Function

Private Sub Form_Load()
    Dim S As Long

    ' Change window style
    S = GetWindowLong(hwnd, GWL_STYLE)
    S = (S And Not WS_POPUP) Or WS_CHILD
    SetWindowLong hwnd, GWL_STYLE, S
    
    ' Change the extended style
    S = GetWindowLong(hwnd, GWL_EXSTYLE)
    S = S Or WS_EX_CONTROLPARENT
    SetWindowLong hwnd, GWL_EXSTYLE, S
    
    S = 80
    ' Set tab stop for lstCtrls
    SendMessage lstCtrls.hwnd, LB_SETTABSTOPS, 1, S
    
    LoadInfo
    
    txtVer.ZOrder 1
    txtName.ZOrder 1
    txtCaption.ZOrder 1
    lstOCXs.ZOrder 1
    lstCtrls.ZOrder 1
    lstHeader.ZOrder 1
    
End Sub

