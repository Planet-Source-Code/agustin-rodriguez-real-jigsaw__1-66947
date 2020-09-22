Attribute VB_Name = "Module1"
Public Const LWA_COLORKEY As Long = &H1
Public Const LWA_ALPHA As Long = &H2
Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Background_choised As Integer


Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal i As Long, ByVal i As Long, ByVal W As Long, ByVal i As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const FLOODFILLSURFACE As Long = 1
Public Objects(0 To 50) As New Img_Object
Public Qt_objects As Integer
Public Ordem_Obj As String

Public Type Txt_Data
    Text As String
    x As Integer
    y As Integer
    BorderColor As Long  ' Cor da Borda
    FillColor As Long    'Cor do Preenchimento
    DrawWidth As Integer 'Grossura da borda
    DrawStyle As Boolean 'Solido ou Transparente
    Ctrl_Color As Long
    Angle As Integer
    FontName As String
    TW As Single
    TH As Single
    
End Type

Public Move_Text As Integer
Public Ordem(1 To 1000) As Integer
Public T(1 To 1000) As Txt_Data
Public Qt_Text As Integer



Public Sub Put_in_order()
Dim i As Integer
For i = 0 To Qt_objects - 1
Objects(i).Show , Form2
Next
Form1.Show , Form2



If Asc(Mid(Ordem_Obj, 1, 1)) = 255 Then
    Form1.Show , Background
    Else
    Objects(Asc(Mid(Ordem_Obj, 1, 1))).Show , Background
End If

For i = 2 To Len(Ordem_Obj)

If Asc(Mid(Ordem_Obj, i, 1)) = 255 Then
    Form1.Show , Objects(Asc(Mid(Ordem_Obj, i - 1, 1)))
Else
    If Asc(Mid(Ordem_Obj, i - 1, 1)) = 255 Then
        Objects(Asc(Mid(Ordem_Obj, i, 1))).Show , Form1
    Else
        Objects(Asc(Mid(Ordem_Obj, i, 1))).Show , Objects(Asc(Mid(Ordem_Obj, i - 1, 1)))
    End If
End If

Next

End Sub
