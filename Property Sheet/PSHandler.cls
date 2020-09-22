VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PSHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Edanmo's Shell Extensions - Property Sheet Handler
'
' Handler Object
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

' Property sheet has to implement
' IShellPropSheetExt and IShellExtInit
Implements olelib.IShellPropSheetExt
Implements olelib.IShellExtInit

'*********************************************************************************************
' Called by the shell to initialize then
' extension.
'
' pidlFolder: pointer to IDL of the parent folder
' lpIDataObject: DataObject containing the selected file names
' hkeyProgID: registry key of the ProgID
'
'*********************************************************************************************
Private Sub IShellExtInit_Initialize(ByVal pidlFolder As Long, ByVal lpIDataObject As olelib.IDataObject, ByVal hkeyProgID As Long)

    GetSelectedFiles g_asSelectedFiles, lpIDataObject

End Sub
'*********************************************************************************************
' Called by the shell to add pages to the
' properties dialog
'
' lpfnAddPage: pointer to a function that will add the pages
' lParam: data to pass to lpfnAddPage
'
'*********************************************************************************************
Private Sub IShellPropSheetExt_AddPages(ByVal lpfnAddPage As Long, ByVal lParam As Long)
Dim PSP As PROPSHEETPAGE, hPage As Long, Caption As String

    On Error Resume Next
    
    ' Store the page caption
    Caption = "&Form Info"
    
    ' Fill PROPSHEETPAGE struct
    With PSP
        .dwSize = Len(PSP)
        
        ' PSP_USETITLE: will use pszTitle member
        ' PSP_USECALLBACK: the pfnCallback member is a valid pointer
        ' PSP_DLGINDIRECT: pResource is a pointer to a DLGTEMPLATE struct
        .dwFlags = PSP_USETITLE Or PSP_USECALLBACK Or PSP_DLGINDIRECT Or PSP_USEHICON
         
        ' Fill the DLGTEMPLATE struct
        ' g_tDlgTemplate must be public because the shell
        ' will use it later.
        With g_tDlgTemplate
            .Style = WS_VISIBLE Or WS_TABSTOP
        End With
        
        .pResource = VarPtr(g_tDlgTemplate)
        
        ' Set the dialog proc function pointer
        .pfnDlgProc = AddrOf(AddressOf mdlDeclares.DlgProc)
        
        ' Set the callback function pointer
        .pfnCallback = AddrOf(AddressOf mdlDeclares.PSCallback)
        
        ' Set the tab caption
        .pszTitle = StrPtr(Caption)
        
        ' Set the tab icon
        .hIcon = 0
        
    End With
    
    ' Create the property page
    hPage = CreatePropertySheetPageW(PSP)
    
    If hPage Then
                
        If AddPage(lpfnAddPage, hPage, lParam) = 0 Then
        
            ' Destroy the page if it cannot
            ' be added.
            DestroyPropertySheetPage hPage
            
        Else
        
            ' Increment object reference
            ' count to prevent the dll from
            ' unloading. Also that reference
            ' will be used in the DlgProc
            ' function to set the parent
            ' of this control to the property
            ' page.
            
            Set g_oHandlerRef = Me
            
        End If
        
    End If
    
End Sub

Private Sub IShellPropSheetExt_ReplacePage(ByVal uPageID As Long, ByVal lpfnReplaceWith As Long, ByVal lParam As Long)

    ' This member is only called for
    ' control panel pages
    
End Sub


