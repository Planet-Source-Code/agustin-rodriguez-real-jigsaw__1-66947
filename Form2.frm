VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   675
      Left            =   -30
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   570
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long

Private Capture As Integer
Private XX As Single
Private YY As Single
Private Type PointAPI
    x As Long
    y As Long
End Type
Private Pt As PointAPI


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
XX = x * Screen.TwipsPerPixelX: YY = y * Screen.TwipsPerPixelY
        Capture = True
        ReleaseCapture
        SetCapture Me.hWnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Capture Then
        GetCursorPos Pt
        Move Pt.x * Screen.TwipsPerPixelX - XX, Pt.y * Screen.TwipsPerPixelY - YY
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Capture = False
End Sub

Private Sub Form_Resize()
Image1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
