VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iSuperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Name.......... iSuperClass
'File.......... iSuperClass.cls
'Dependencies.. Meaningless without cSuperClass
'Description... cSuperClass implemention interface model as used by cSuperClass to report window messages
'               to the owner of the cSuperClass instance. ??? look up implements in VB help ;>)
'Author........ Paul_Caton@hotmail.com
'Date.......... June, 13th 2002
'Copyright..... None.

Option Explicit

Public Sub After(ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

  '
  'lReturn  - Will contain the return value from the previous WndProc, populate it yourself as the message/situation/Windows dictates
  '
  'hWnd     - Handle of the window
  '
  'uMsg     - Message number
  '
  'wParam   - Message related data
  '
  'lParam   - Message related data
  '
  
End Sub



