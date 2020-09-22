VERSION 5.00
Begin VB.Form Img_Object 
   Appearance      =   0  'Flat
   BackColor       =   &H00BC614E&
   BorderStyle     =   0  'None
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFE&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Object_form.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   164
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Original 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1470
      Left            =   1440
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   11
      Top             =   3045
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CheckBox FillShape 
      Caption         =   "Check1"
      Height          =   240
      Left            =   780
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicCol 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   465
      Left            =   300
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   3165
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Link 
      Height          =   285
      Left            =   705
      TabIndex        =   10
      Top             =   3105
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Transparent 
      Caption         =   "255"
      Height          =   285
      Left            =   270
      TabIndex        =   8
      Top             =   2790
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Translucent 
      Caption         =   "0"
      Height          =   285
      Left            =   255
      TabIndex        =   7
      Top             =   2310
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Arquivo 
      Height          =   285
      Left            =   195
      TabIndex        =   6
      Top             =   1965
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label EscalaY 
      Caption         =   "1"
      Height          =   285
      Left            =   225
      TabIndex        =   5
      Top             =   1530
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label EscalaX 
      Caption         =   "1"
      Height          =   285
      Left            =   210
      TabIndex        =   4
      Top             =   1185
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label EscalaZY 
      Caption         =   "0"
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   810
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label EscalaZX 
      Caption         =   "0"
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Angle 
      Caption         =   "0"
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "Img_Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long

Private Type PointAPI
    x As Long
    y As Long
End Type

Private Pt As PointAPI
Private Radiano As Single
Private Capture As Integer
Private Button_press As Integer
Private Saturação As Integer
Private Gamma As Integer
Private Control_press As Integer
Private Intensidade_de_Saturação As Integer
Private Correção_de_Gamma As Integer
Private Não_processe As Integer

Private BaseWidth As Long
Private BaseHeight As Long

Private XX As Single
Private YY As Single

Private Const WM_MOUSEWHEEL       As Long = &H20A

Private sc          As cSuperClass

Implements iSuperClass



Private Sub Form_Activate()

    SetLayeredWindowAttributes Me.hWnd, Transparent, Translucent, LWA_COLORKEY Or LWA_ALPHA
    Set_image

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
  Dim r As Integer, x As Long, y As Long, i As Integer, s As Integer
         
    Select Case KeyCode
              
      Case 65, 83, 87, 90, 107, 109, 255
        
        If KeyCode = 107 Then
xxx:
            EscalaZX = EscalaZX + 0.1
            EscalaZY = EscalaZY + 0.1
        End If
    
        If KeyCode = 109 Then
            EscalaZX = EscalaZX - 0.1
            EscalaZY = EscalaZY - 0.1
        End If
    
        AutoRedraw = True
        x = Left: y = Top
        If KeyCode = 255 Then
            GoTo pula
        End If
        
        If EscalaZX <= 0.1 Then
            EscalaZX = 0.1
        End If
        If EscalaZY <= 0.1 Then
            EscalaZY = 0.1
        End If
pula:
        
        Set_image
        Exit Sub
        
      Case 38, 40
       
        Translucent = (Translucent - 5 * (KeyCode = 38) + 5 * (KeyCode = 40))
        If Translucent < 10 Then
            Translucent = 10
        End If
        If Translucent > 255 Then
            Translucent = 255
        End If
        SetLayeredWindowAttributes Me.hWnd, Transparent, Translucent, LWA_COLORKEY Or LWA_ALPHA
        Exit Sub
    
      Case 39
       
        Angle = Angle + 15
        If Angle = 360 Then
            Angle = 0
        End If
        Pintar
        
      Case 37
     
        Angle = Angle - 15
        If Angle = 0 Then
            Angle = 360
        End If
        Pintar
  
    End Select

End Sub

Private Sub Form_Load()

  Dim Normalwindowstyle As Long
  Dim Ret As Long
  Dim n As String
  Dim col As Long
  Dim i As Integer
    
    EscalaZX = 1
    EscalaZY = 1
   
    Set sc = New cSuperClass
  
    With sc
        Call .AddMsg(WM_MOUSEWHEEL)
        Call .Subclass(hWnd, Me)
    End With
  
    Translucent = 255
    Transparent = 0
    
    Normalwindowstyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    col = 12345678
    SetLayeredWindowAttributes Me.hWnd, Transparent, Translucent, LWA_COLORKEY Or LWA_ALPHA
    
    Correção_de_Gamma = 100
    Set_image
    Move Screen.Width / 2, Screen.Height / 3
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim returnval As Long
  Dim i As Integer
    
    Button_press = Button
    Control_press = Shift
        
    Form1.Trabalho.Width = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
    Form1.Trabalho.Height = Form1.Trabalho.Width
      
    
    For i = 0 To Qt_objects
        Lf(i) = f(i).Left
        Tf(i) = f(i).Top
    Next i
    xxx = Left: yyy = Top
        
    MouseIcon = LoadPicture(App.Path & "\hmove.cur")
    XX = x * Screen.TwipsPerPixelX: YY = y * Screen.TwipsPerPixelY
    Capture = True
    ReleaseCapture
    SetCapture Me.hWnd
    
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim i As Integer

    If Capture Then
        GetCursorPos Pt
        Move Pt.x * Screen.TwipsPerPixelX - XX, Pt.y * Screen.TwipsPerPixelY - YY
        Tag = "x"
        If Button = 2 Then
            For i = 0 To Qt_objects
                If f(i).Tag <> "x" Then
                    f(i).Move Lf(i) + (Left - xxx), Tf(i) + (Top - yyy)
                End If
            Next i
            Tag = ""
        End If
        
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    MouseIcon = LoadPicture(App.Path & "\harrow.cur")
    Capture = False
    Tag = ""

End Sub

Private Sub Form_Paint()

  Static vez As Integer

    If vez = 0 Then
        vez = 1
        Cls
        Rotate Me.hdc, ScaleWidth / 2, ScaleHeight / 2, Angle, PicCol.hdc, 0, 0, PicCol.Width, PicCol.Height
        AutoRedraw = True
    
        Rotate Me.hdc, ScaleWidth / 2, ScaleHeight / 2, Angle, PicCol.hdc, 0, 0, PicCol.Width, PicCol.Height
        AutoRedraw = False
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set sc = Nothing

End Sub

Private Sub PicCol_Resize()

    If Não_processe Then
        Exit Sub
    End If

    Form1.Trabalho.Width = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
    Form1.Trabalho.Height = Form1.Trabalho.Width
    Width = Form1.Trabalho.ScaleWidth * Screen.TwipsPerPixelX
    Height = Form1.Trabalho.ScaleHeight * Screen.TwipsPerPixelY
 
    BaseWidth = Width
    BaseHeight = Height

End Sub

Private Sub Pintar()

    Form1.Trabalho.BackColor = Transparent
    Form1.Trabalho.Cls
    Rotate Form1.Trabalho.hdc, Form1.Trabalho.ScaleWidth / 2, Form1.Trabalho.ScaleHeight / 2, Angle, PicCol.hdc, 0, 0, PicCol.Width, PicCol.Height
    Me.PaintPicture Form1.Trabalho.Image, 0, 0, Form1.Trabalho.Width * EscalaX, Form1.Trabalho.Height * EscalaY, 0, 0, Form1.Trabalho.Width, Form1.Trabalho.Height, vbSrcCopy
        
End Sub

Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  
  'Case WM_MOUSEWHEEL
  
    Select Case wParam
      Case -7864308 'rotate
        'Debug.Print "UP+Shift+Control";
        'Form_KeyDown 39, 0
        
      Case 7864332 'rotate
        'Debug.Print "DOWN+Shift+Control";
        'Form_KeyDown 37, 0
        
      Case 7864324
        'Debug.Print "UP+Shift";
        Form_KeyDown Asc("A"), 0
        
        ''''''''''''''''Form_KeyDown 40, 0
        
      Case -7864316
        'Debug.Print "DOWN+Shift";
        Form_KeyDown Asc("S"), 0
        
        'Form_KeyDown 38, 0
        
      Case -7864320
        Form_KeyDown 39, 0
        ''''''''''''Form_KeyDown 107, 0
        'Debug.Print "DOWN";
        
      Case 7864320
        Form_KeyDown 37, 0
        ''''''''''''Form_KeyDown 109, 0
        'Debug.Print "UP";
        
      Case -7864312
        Form_KeyDown Asc("W"), 0
        ''''''''''Form_KeyDown 40, 1
        'Debug.Print "DOWN + Control";
      
      Case 7864328
        Form_KeyDown Asc("Z"), 0
        ''''''''''''''Form_KeyDown 38, 1
        
        'Debug.Print "UP + Control";
    End Select
        
    'Debug.Print wParam;

End Sub

Public Sub Set_image()

  Dim Quadrado_Original As Long
  Dim Quadrado_Atual As Long
  Dim i As Integer
            
    Form1.Trabalho.Width = Original.Width * EscalaZX
    Form1.Trabalho.Height = Original.Height * EscalaZY
    Form1.Trabalho.Cls
    
    Form1.Trabalho.PaintPicture Original.Image, 0, 0, Form1.Trabalho.Width, Form1.Trabalho.Height, 0, 0, Original.Width, Original.Height, vbSrcCopy
                
    Não_processe = True
    Quadrado_Original = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
    PicCol.Width = Form1.Trabalho.Width
    PicCol.Height = Form1.Trabalho.Height
    Quadrado_Atual = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
    Não_processe = False
        
    PicCol_Resize
    PicCol.Cls
    AutoRedraw = True
    
    PicCol.PaintPicture Form1.Trabalho.Image, 0, 0, Form1.Trabalho.Width, Form1.Trabalho.Height, 0, 0, Form1.Trabalho.Width, Form1.Trabalho.Height, vbSrcCopy
    
    Pintar
    AutoRedraw = False
    Refresh
    Move Left - ((Quadrado_Atual - Quadrado_Original) / 2) * Screen.TwipsPerPixelX, Top - ((Quadrado_Atual - Quadrado_Original) / 2) * Screen.TwipsPerPixelY
    
End Sub


