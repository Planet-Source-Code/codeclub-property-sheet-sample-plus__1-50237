Attribute VB_Name = "mdlFilesFromIDO"
'*********************************************************************************************
'
' Shell Extensions
'
' GetSelectedFiles function
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
'           03/23/2000 - Added error checking on the call
'                        to IDataObject.GetData
'           12/01/1999 - This code was released
'
'*********************************************************************************************
Option Explicit


'
' GetSelectedFiles
'
' Get the filenames from a IDataObject object
'
' Files(): array to fill
' IDO: IDataObject interface with selected files
'
Public Sub GetSelectedFiles(Files() As String, IDO As IDataObject)
Dim FMT As FORMATETC, STM As STGMEDIUM
Dim Idx As Long     ' Current array index
Dim Max As Long     ' Filename count

    ' Catch all error an ignore
    ' them. Why? Because if VB
    ' raises an error and
    ' Windows Explorer doesn't
    ' expects one it will produce
    ' a GPF
    On Error Resume Next
    
    ' Erase the array
    Erase Files()
    
    ' Fill the FORMATETC struct
    ' to retrieve the filename
    ' data in CF_HDROP format
    With FMT
        .cfFormat = CF_HDROP
        .TYMED = TYMED_HGLOBAL
        .dwAspect = DVASPECT_CONTENT
    End With
    
    ' Get the data from IDataObject
    ' This call will fill the STM
    ' struct with a pointer to
    ' the DROPFILES struct
    If IDO.GetData(FMT, STM) = S_OK Then ' Get files only if GetData returns S_OK
    
      ' Get file name count
      Max = DragQueryFile(STM.Data, -1, vbNullString, 0)
                  
      ReDim Preserve Files(0 To Max - 1)
          
      ' Get filenames
      For Idx = 0 To Max - 1
      
          Files(Idx) = String$(260, 0)
          
          DragQueryFile STM.Data, Idx, Files(Idx), Len(Files(Idx))
          
          If InStr(Files(Idx), vbNullChar) > 0 Then Files(Idx) = Left$(Files(Idx), InStr(Files(Idx), vbNullChar) - 1)
          
      Next
      
      ' Release memory used by
      ' STM.Data
      ReleaseStgMedium STM
   
   End If
   
End Sub


