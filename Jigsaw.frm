VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Real Jigsaw"
   ClientHeight    =   4410
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12180
   Icon            =   "Jigsaw.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   812
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Height          =   390
      Left            =   -15
      ScaleHeight     =   330
      ScaleWidth      =   6810
      TabIndex        =   12
      Top             =   0
      Width           =   6870
      Begin VB.VScrollBar VScroll1 
         Height          =   315
         Left            =   1215
         Max             =   -2
         Min             =   -10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
         Value           =   -5
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   315
         Left            =   3060
         Max             =   -2
         Min             =   -10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   15
         Value           =   -5
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2220
         TabIndex        =   17
         Top             =   105
         Width           =   120
      End
      Begin VB.Label N_row 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   16
         Top             =   15
         Width           =   615
      End
      Begin VB.Label N_col 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2445
         TabIndex        =   15
         Top             =   15
         Width           =   615
      End
   End
   Begin VB.PictureBox Trabalho 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8295
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   10
      Top             =   750
      Visible         =   0   'False
      Width           =   630
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   10560
      Top             =   5835
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   885
      Left            =   6855
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   9
      Top             =   540
      Visible         =   0   'False
      Width           =   1305
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   8280
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   5790
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   8
      Left            =   11715
      Picture         =   "Jigsaw.frx":164A
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   8
      Top             =   6390
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   7
      Left            =   9285
      Picture         =   "Jigsaw.frx":14FD0
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   7
      Top             =   6390
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   6
      Left            =   6840
      Picture         =   "Jigsaw.frx":28956
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   6
      Top             =   6390
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   5
      Left            =   11715
      Picture         =   "Jigsaw.frx":3C2DC
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   5
      Top             =   3945
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   4
      Left            =   9270
      Picture         =   "Jigsaw.frx":4FC62
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   4
      Top             =   3945
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   3
      Left            =   6840
      Picture         =   "Jigsaw.frx":635E8
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   3
      Top             =   3945
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   2
      Left            =   11730
      Picture         =   "Jigsaw.frx":76F6E
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   2
      Top             =   1500
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   1
      Left            =   9300
      Picture         =   "Jigsaw.frx":8A8F4
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Index           =   0
      Left            =   6855
      Picture         =   "Jigsaw.frx":9E27A
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Menu Files 
      Caption         =   "Files"
      Begin VB.Menu Open_picture 
         Caption         =   "Open Picture"
      End
      Begin VB.Menu Open_Jigsaw 
         Caption         =   "Open Jigsaw"
      End
      Begin VB.Menu Save_Jigsaw 
         Caption         =   "Save Jigsaw"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Rotation 
         Caption         =   "Rotation"
         Checked         =   -1  'True
      End
      Begin VB.Menu Show_hide_picture 
         Caption         =   "Show/Hide Picture"
      End
      Begin VB.Menu Show_Pieces_number 
         Caption         =   "Show Pieces Numbers"
      End
      Begin VB.Menu Zoom__in_out 
         Caption         =   "Zoom"
         Begin VB.Menu zoom 
            Caption         =   "In 10 %"
            Index           =   0
         End
         Begin VB.Menu zoom 
            Caption         =   "Out 10 %"
            Index           =   1
         End
         Begin VB.Menu zoom 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu zoom 
            Caption         =   "50 %"
            Index           =   3
         End
         Begin VB.Menu zoom 
            Caption         =   "75 %"
            Index           =   4
         End
         Begin VB.Menu zoom 
            Caption         =   "100 %"
            Index           =   5
         End
         Begin VB.Menu zoom 
            Caption         =   "150 %"
            Index           =   6
         End
         Begin VB.Menu zoom 
            Caption         =   "200 %"
            Index           =   7
         End
      End
   End
   Begin VB.Menu Create_Jigsaw 
      Caption         =   "           Create Jigsaw         "
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu About_index 
         Caption         =   "Autor: Agustin Rodriguez"
         Index           =   0
      End
      Begin VB.Menu About_index 
         Caption         =   "E-Mail: virtual_guitar_1@hotmail.com"
         Index           =   1
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Help_index 
         Caption         =   "- Use the Mouse Wheel to Rotate the piece"
         Index           =   0
      End
      Begin VB.Menu Help_index 
         Caption         =   "- Drag one Piece using the Mouse Button 1"
         Index           =   1
      End
      Begin VB.Menu Help_index 
         Caption         =   "- Drag all Jigsaw using the Mouse  Button 2"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Actual_picture_file As String * 256
Private Actual_jigsaw_file As String

Private G As New clsGradient
Private W As Single
Private h As Single
Private DW As Single
Private DH As Single
Private Type Data_files
    Left As Long
    Top As Long
    Angle  As Long
    EscalaX As Single
    EscalaY As Single
    EscalaZX As Single
    EscalaZY As Single
End Type

Private Sub Create_Jigsaw_Click()

  Dim i As Single
  Dim k As Single

    Screen.MousePointer = 11
    Picture2.Visible = False

    For i = 0 To Qt_objects - 1
        Unload f(i)
    Next i
    
    ReDim f(N_row * N_col)
    Qt_objects = 0
    
    Picture2.Cls
    W = Picture2.Width / N_row
    h = Picture2.Height / N_col
    DW = W * 30 / Picture1(0).Width
    DH = h * 30 / Picture1(0).Width

    For i = 0 To Picture2.Width - 1 Step W
        For k = 0 To Picture2.Height - 1 Step h
            If i = 0 Then
                If k = 0 Then
                    Create_Piece 0, i, k
                    GoTo siga
                End If
    
                If Int(k) = Int((N_col - 1) * h) Then
                    Create_Piece 6, i, k
                    GoTo siga
                End If
    
                Create_Piece 3, i, k
                GoTo siga
              Else

                If Int(i) = Int((N_row - 1) * W) Then
                    If k = 0 Then
                        Create_Piece 2, i, k
                        GoTo siga
                    End If
    
                    If Int(k) = Int((N_col - 1) * h) Then
                        Create_Piece 8, i, k
                        GoTo siga
                    End If
    
                    Create_Piece 5, i, k
                    GoTo siga
                End If
            End If

            If k = 0 Then
                Create_Piece 1, i, k
                GoTo siga
              Else
                If Int(k) = Int((N_col - 1) * h) Then
                    Create_Piece 7, i, k
                    GoTo siga
                  Else
                    Create_Piece 4, i, k
                    GoTo siga
                End If
            End If
siga:
        Next k
    Next i
    
    Put_or_remove_numbers
    Screen.MousePointer = 0

End Sub

Private Sub Exit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    G.Color1 = &HFF
    G.Color2 = 0
        
End Sub

Private Sub Form_Resize()

  Dim i As Integer
    Picture4.Width = ScaleWidth
    G.Angle = -DeriveTheta(ScaleWidth / 2, ScaleHeight / 2, 10, 1000)
    G.Draw Me
    Refresh
    
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim i As Integer

    For i = 0 To Qt_objects - 1
        Unload f(i)
    Next i
    Unload Form2
    End

End Sub

Private Sub Open_Jigsaw_Click()
    
    CD.FileName = ""
    CD.DefaultExt = ".Jig"
    CD.Filter = "Jigsaw|*.Jig"
    CD.Filter = "Jigsaw|*.Jig"
    CD.Action = 1
    
    If CD.FileName = "" Then
        Exit Sub
    End If
    
    Load_Jigsaw (CD.FileName)
    
    
End Sub

Private Sub Open_picture_Click()
    
    CD.FileName = ""
    CD.Filter = "All Files|*.bmp;*gif;*.jpg|BitMap|*.bmp|Compuserve|*.gif|Jasc|*.jpg"
    CD.Action = 1
    If CD.FileName = "" Then
        Exit Sub
    End If
    
    Load_PIcture (CD.FileName)
    Actual_picture_file = CD.FileName
    
End Sub

Private Sub Create_Piece(Index As Integer, ii As Single, kk As Single)

    If Picture2.Picture = 0 Then
        Exit Sub
    End If
    
    If Rotation.Checked Then
        f(Qt_objects).Angle = Int(Rnd * 23) * 15
    End If
    f(Qt_objects).Show , Me
    f(Qt_objects).Original.Width = W + DW
    f(Qt_objects).Original.Height = h + DH
    f(Qt_objects).Original.PaintPicture Picture2.Picture, 0, 0, W + DW, h + DH, ii, kk - DH, W + DW, h + DH, vbSrcCopy
    f(Qt_objects).Original.PaintPicture Picture1(Index).Picture, 0, 0, W + DW, h + DH, , , , , vbSrcAnd
    f(Qt_objects).Original.Picture = f(Qt_objects).Original.Image
    f(Qt_objects).Move Screen.Width / 3 + Int(Rnd * Screen.Width / 3), Screen.Height / 3 + Int(Rnd * Screen.Height / 3)
    Qt_objects = Qt_objects + 1

End Sub

Private Sub Rotation_Click()

    Rotation.Checked = Rotation.Checked Xor -1

End Sub

Private Sub Save_Jigsaw_Click()

  Dim Free As Integer
  Dim x As String
  Dim s As Integer
  Dim i As Integer
  Dim Arq As Data_files
    CD.Flags = &H2
    CD.FileName = Actual_jigsaw_file
    CD.DefaultExt = ".Jig"
    CD.Filter = "Jigsaw|*.Jig"
    CD.Action = 2
    If CD.FileName = "" Then
        Exit Sub
    End If
    Free = FreeFile

    Open CD.FileName For Binary As Free
    Put #Free, 1, Actual_picture_file
    s = N_row.Caption
    Put #Free, , s
    s = N_col.Caption
    Put #Free, , s
    s = Rotation.Checked
    Put #Free, , s
    s = Show_Pieces_number.Checked
    Put #Free, , s
    
    For i = 0 To Qt_objects - 1
        Arq.Left = f(i).Left
        Arq.Top = f(i).Top
        Arq.Angle = f(i).Angle.Caption
        Arq.EscalaX = f(i).EscalaX.Caption
        Arq.EscalaY = f(i).EscalaY.Caption
        Arq.EscalaZX = f(i).EscalaZX.Caption
        Arq.EscalaZY = f(i).EscalaZY.Caption
        Put #Free, , Arq
    Next i
    
    Close Free
    
End Sub

Private Sub Show_hide_picture_Click()

    Form2.Visible = Not Form2.Visible

End Sub

Private Sub Show_Piecs_number_Click()

End Sub

Private Sub Show_Pieces_number_Click()

  Dim i As Integer

    Show_Pieces_number.Checked = Show_Pieces_number.Checked Xor -1

    Put_or_remove_numbers

End Sub

Private Sub VScroll1_Change()

    N_row = Abs(VScroll1.Value)

End Sub

Private Sub VScroll2_Change()

    N_col = Abs(VScroll2.Value)

End Sub

Private Sub Zoom_Click(Index As Integer)

  Dim i As Integer
    Dim W As Long
    Dim h As Long
    Dim l As Long
    Dim T As Long
    
    Screen.MousePointer = 11
    
    For i = 0 To Qt_objects
    W = f(i).Width
    h = f(i).Height
    l = f(i).Left
    T = f(i).Top
    
        Select Case Index
          Case 0
            f(i).Form_KeyDown 109, 0
          Case 1
            f(i).Form_KeyDown 107, 0
          Case 3
            f(i).EscalaZX = 0.5
            f(i).EscalaZY = 0.5
            f(i).Set_image
          Case 4
            f(i).EscalaZX = 0.75
            f(i).EscalaZY = 0.75
            f(i).Set_image
          Case 5
            f(i).EscalaZX = 1
            f(i).EscalaZY = 1
            f(i).Set_image
          Case 6
            f(i).EscalaZX = 1.5
            f(i).EscalaZY = 1.5
            f(i).Set_image
          Case 7
            f(i).EscalaZX = 2
            f(i).EscalaZY = 2
            f(i).Set_image
        End Select
'f(i).Move l * (f(i).Width / W), T * (f(i).Height / h)
    f(i).Move l * (f(i).Width / W) - (f(i).Width - W), T * (f(i).Height / h) - (f(i).Height - h)
    
    Next i
    Screen.MousePointer = 0

End Sub

Private Function DeriveTheta(x As Single, y As Single, DestX As Single, DestY As Single) As Single

  Dim TempX As Single
  Dim TempY As Single
    
    On Error GoTo erro
    TempX = DestX - x
    TempY = DestY - y
    
    If TempX >= 0 Then
        DeriveTheta = Atn(TempY / TempX) * 57.2957795130824 + 90
    End If
    If TempX <= 0 Then
        DeriveTheta = Atn(TempY / TempX) * 57.2957795130824 + 270
    End If

Exit Function

erro:
    Resume Next

End Function

Private Sub Put_or_remove_numbers()

  Dim i As Integer

    For i = 0 To Qt_objects - 1
        If Show_Pieces_number.Checked Then
            f(i).Original.ForeColor = 1
            f(i).Original.CurrentX = 0
            f(i).Original.CurrentY = DH
            f(i).Original.Print i + 1
            f(i).Original.ForeColor = &HFFFFFF
            f(i).Original.CurrentX = 2
            f(i).Original.CurrentY = DH + 2
            f(i).Original.Print i + 1
          Else
            f(i).Original.Cls
        End If
        f(i).Set_image
    Next i

End Sub

Private Sub Load_PIcture(Name As String)
  
  Dim x As Single
  Dim y As Single
  Static first As Integer
  
    Picture3.Picture = LoadPicture(Name)
    x = Picture3.Width
    y = Picture3.Height
    Do While x > 640 Or y > 480
        x = x - x / 100
        y = y - y / 100
    Loop
    
    Do While x < 160 Or y < 120
        x = x + x / 100
        y = y + y / 100
    Loop
    
    Picture2.Move 10, 35, x, y
    Picture2.PaintPicture Picture3.Picture, 0, 0, x, y
    Picture2.Picture = Picture2.Image
    
    Picture3.Picture = LoadPicture("")
    Form2.Move 0, 1200, 3000, (3000 / Picture2.Width) * Picture2.Height
    Form2.Image1.Picture = Picture2.Picture
    If first Then
        Form2.Show , Me
    Else
        first = True
    End If
End Sub

Private Sub Load_Jigsaw(JigSaw_Name As String)

  Dim Free As Integer
  Dim i As Integer
  Dim x As String
  Dim s As Integer
  Dim Arq As Data_files
  Dim Bkp_actual_picture_file As String

    Bkp_actual_picture_file = Actual_picture_file
    Actual_jigsaw_file = JigSaw_Name
    
    Free = FreeFile
    Open JigSaw_Name For Binary As Free
    Get #Free, 1, Actual_picture_file
    If File_can_to_be_load(Actual_picture_file) = False Then
        For i = Len(Actual_picture_file) To 1 Step -1
            If Mid$(Actual_picture_file, i, 1) = "\" Then
                Actual_picture_file = App.Path & Mid$(Actual_picture_file, i)
                If File_can_to_be_load(Actual_picture_file) = False Then
                    s = MsgBox("Picture not found" & vbCrLf & Trim(Actual_picture_file), vbCritical)
                    Actual_picture_file = Bkp_actual_picture_file
                    Close Free
                    Exit Sub
                  Else
                    Exit For
                End If
            End If
        Next i
    End If
    
    Load_PIcture (Actual_picture_file)
    DoEvents
    
    Get #Free, , s
    N_row.Caption = s
    Get #Free, , s
    N_col.Caption = s
    Get #Free, , s
    Rotation.Checked = s
    Get #Free, , s
    Show_Pieces_number.Checked = s
    
    Create_Jigsaw_Click
    
    For i = 0 To Qt_objects - 1
        Get #Free, , Arq
        f(i).Left = Arq.Left
        f(i).Top = Arq.Top
        f(i).Angle.Caption = Arq.Angle
        f(i).EscalaX.Caption = Arq.EscalaX
        f(i).EscalaY.Caption = Arq.EscalaX
        f(i).EscalaZX.Caption = Arq.EscalaZX
        f(i).EscalaZY.Caption = Arq.EscalaZY
        f(i).Set_image
    Next i
    Close #Free

End Sub

Public Sub First_time()
  Load_Jigsaw (App.Path & "\Jigsaw Splash.jig")
End Sub

      
Private Function File_can_to_be_load(x As String)
On Error GoTo erro

File_can_to_be_load = False
If Dir(x) <> "" Then
    File_can_to_be_load = True
End If

sair:
Exit Function

erro:
Resume sair

End Function
