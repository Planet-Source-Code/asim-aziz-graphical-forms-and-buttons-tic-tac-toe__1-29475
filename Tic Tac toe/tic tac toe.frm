VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4500
   ClientLeft      =   2100
   ClientTop       =   3555
   ClientWidth     =   6000
   Icon            =   "tic tac toe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "tic tac toe.frx":0442
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2730
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   35
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":44F1
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":4BF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":52F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":59E9
            Key             =   "exitp"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":60E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":6929
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":716D
            Key             =   "cross"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tic tac toe.frx":7D19
            Key             =   "ring"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2550
      Picture         =   "tic tac toe.frx":89A1
      Top             =   3750
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000040&
      BorderWidth     =   2
      Height          =   2985
      Left            =   1290
      Top             =   690
      Width           =   3435
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   1
      Left            =   1365
      Top             =   735
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   2
      Left            =   2490
      Top             =   735
      Width           =   1035
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   3
      Left            =   3615
      Top             =   735
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   4
      Left            =   1365
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   5
      Left            =   2490
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   6
      Left            =   3615
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   7
      Left            =   1365
      Top             =   2685
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   8
      Left            =   2490
      Top             =   2685
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   9
      Left            =   3615
      Top             =   2685
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000040&
      BorderWidth     =   2
      X1              =   162
      X2              =   162
      Y1              =   243.667
      Y2              =   47
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000040&
      BorderWidth     =   2
      X1              =   238
      X2              =   238
      Y1              =   243.667
      Y2              =   47
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000040&
      BorderWidth     =   2
      X1              =   87
      X2              =   313.667
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000040&
      BorderWidth     =   2
      X1              =   87
      X2              =   313.667
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   435
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   210
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   435
      Index           =   1
      Left            =   4575
      TabIndex        =   2
      Top             =   210
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   0
      Left            =   5100
      Picture         =   "tic tac toe.frx":91D3
      Top             =   3660
      Width           =   525
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   1
      Left            =   390
      Picture         =   "tic tac toe.frx":98B9
      Top             =   3660
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   555
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   555
      Index           =   1
      Left            =   4785
      TabIndex        =   0
      Top             =   660
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main form of this application buttons and X O s are all standard image
'controls (images are in imagelist1).
'the grid in the center (the play area is an array of 9 image controls ie. image1.
'The logic i've used for the actual game might be confusing at first look.
'my primary goal wasnt the perfect unbeatable tic tac toe logic anyway (You can
'always have a go at it in hard mode:) )
'any sujjestions are welcome  "chirisoft@flashmail.com"
'sorry for the lack or rather abscence of comments within code

Option Explicit
Dim Rgn1 As Long
Dim P1 As Integer, P2 As Integer, A As Integer, B As Integer, C As Integer
Dim Allowed As Boolean
Public Difficulty As Integer

Private Sub Form_Load()
'to create form with rounded edges
Rgn1 = CreateRoundRectRgn(0, 0, 400, 300, 20, 20)
SetWindowRgn hWnd, Rgn1, True

Difficulty = 2
Allowed = True
Randomize
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteObject Rgn1
End Sub


Private Sub Image1_Click(Index As Integer)
On Error Resume Next
If Allowed = True Then
If Image1(Index).Tag = "" Then
Image1(Index).Picture = ImageList1.ListImages("cross").Picture
Image1(Index).Tag = "p"


If CheckUserWin Then
Allowed = False
P1 = P1 + 1
Label2(0).Caption = P1
blink
ClearAll
ElseIf CheckWin Then
Allowed = False
P2 = P2 + 1
Label2(1).Caption = P2
blink
ClearAll
Else
TakeTurn
DoEvents
Sleep 200
If CheckDraw Then
ClearAll
End If
End If

Else: Beep
End If
End If
End Sub


Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
If Index = 0 Then
Image2(0).Picture = ImageList1.ListImages(4).Picture
ElseIf Index = 1 Then
Image2(1).Picture = ImageList1.ListImages(2).Picture
End If
End If
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then

If Index = 0 Then
Image2(0).Picture = ImageList1.ListImages(3).Picture
If X > 0 And X < 525 And Y > 0 And Y < 450 Then
Unload Me
End
End If


ElseIf Index = 1 Then
Image2(1).Picture = ImageList1.ListImages(1).Picture
If X > 0 And X < 525 And Y > 0 And Y < 450 Then
ClearAll
P1 = 0
P2 = 0
Label2(0).Caption = 0
Label2(1).Caption = 0
End If
End If

End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&
End If
End Sub
Private Function CheckWin() As Boolean
On Error Resume Next
If Image1(1).Tag = "c" And Image1(2).Tag = "c" And Image1(3).Tag <> "p" Then
Image1(3).Picture = ImageList1.ListImages("ring").Picture
Image1(3).Tag = "c"
CheckWin = True
A = 1
B = 2
C = 3

ElseIf Image1(1).Tag = "c" And Image1(3).Tag = "c" And Image1(2).Tag <> "p" Then
Image1(2).Picture = ImageList1.ListImages("ring").Picture
Image1(2).Tag = "c"
CheckWin = True
A = 1
B = 2
C = 3

ElseIf Image1(2).Tag = "c" And Image1(3).Tag = "c" And Image1(1).Tag <> "p" Then
Image1(1).Picture = ImageList1.ListImages("ring").Picture
Image1(1).Tag = "c"
CheckWin = True
A = 1
B = 2
C = 3

ElseIf Image1(1).Tag = "c" And Image1(4).Tag = "c" And Image1(7).Tag <> "p" Then
Image1(7).Picture = ImageList1.ListImages("ring").Picture
Image1(7).Tag = "c"
CheckWin = True
A = 1
B = 4
C = 7

ElseIf Image1(1).Tag = "c" And Image1(7).Tag = "c" And Image1(4).Tag <> "p" Then
Image1(4).Picture = ImageList1.ListImages("ring").Picture
Image1(4).Tag = "c"
CheckWin = True
A = 1
B = 4
C = 7

ElseIf Image1(4).Tag = "c" And Image1(7).Tag = "c" And Image1(1).Tag <> "p" Then
Image1(1).Picture = ImageList1.ListImages("ring").Picture
Image1(1).Tag = "c"
CheckWin = True
A = 1
B = 4
C = 7

ElseIf Image1(1).Tag = "c" And Image1(5).Tag = "c" And Image1(9).Tag <> "p" Then
Image1(9).Picture = ImageList1.ListImages("ring").Picture
Image1(9).Tag = "c"
CheckWin = True
A = 1
B = 5
C = 9

ElseIf Image1(1).Tag = "c" And Image1(9).Tag = "c" And Image1(5).Tag <> "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
CheckWin = True
A = 1
B = 5
C = 9

ElseIf Image1(5).Tag = "c" And Image1(9).Tag = "c" And Image1(1).Tag <> "p" Then
Image1(1).Picture = ImageList1.ListImages("ring").Picture
Image1(1).Tag = "c"
CheckWin = True
A = 1
B = 5
C = 9

ElseIf Image1(2).Tag = "c" And Image1(5).Tag = "c" And Image1(8).Tag <> "p" Then
Image1(8).Picture = ImageList1.ListImages("ring").Picture
Image1(8).Tag = "c"
CheckWin = True
A = 2
B = 5
C = 8

ElseIf Image1(2).Tag = "c" And Image1(8).Tag = "c" And Image1(5).Tag <> "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
CheckWin = True
A = 2
B = 5
C = 8

ElseIf Image1(5).Tag = "c" And Image1(8).Tag = "c" And Image1(2).Tag <> "p" Then
Image1(2).Picture = ImageList1.ListImages("ring").Picture
Image1(2).Tag = "c"
CheckWin = True
A = 2
B = 5
C = 8

ElseIf Image1(3).Tag = "c" And Image1(6).Tag = "c" And Image1(9).Tag <> "p" Then
Image1(9).Picture = ImageList1.ListImages("ring").Picture
Image1(9).Tag = "c"
CheckWin = True
A = 3
B = 6
C = 9

ElseIf Image1(3).Tag = "c" And Image1(9).Tag = "c" And Image1(6).Tag <> "p" Then
Image1(6).Picture = ImageList1.ListImages("ring").Picture
Image1(6).Tag = "c"
CheckWin = True
A = 3
B = 6
C = 9

ElseIf Image1(6).Tag = "c" And Image1(9).Tag = "c" And Image1(3).Tag <> "p" Then
Image1(3).Picture = ImageList1.ListImages("ring").Picture
Image1(3).Tag = "c"
CheckWin = True
A = 3
B = 6
C = 9

ElseIf Image1(3).Tag = "c" And Image1(5).Tag = "c" And Image1(7).Tag <> "p" Then
Image1(7).Picture = ImageList1.ListImages("ring").Picture
Image1(7).Tag = "c"
CheckWin = True
A = 3
B = 5
C = 7

ElseIf Image1(3).Tag = "c" And Image1(7).Tag = "c" And Image1(5).Tag <> "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
CheckWin = True
A = 3
B = 5
C = 7

ElseIf Image1(5).Tag = "c" And Image1(7).Tag = "c" And Image1(3).Tag <> "p" Then
Image1(3).Picture = ImageList1.ListImages("ring").Picture
Image1(3).Tag = "c"
CheckWin = True
A = 3
B = 5
C = 7

ElseIf Image1(4).Tag = "c" And Image1(5).Tag = "c" And Image1(6).Tag <> "p" Then
Image1(6).Picture = ImageList1.ListImages("ring").Picture
Image1(6).Tag = "c"
CheckWin = True
A = 4
B = 5
C = 6

ElseIf Image1(4).Tag = "c" And Image1(6).Tag = "c" And Image1(5).Tag <> "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
CheckWin = True
A = 4
B = 5
C = 6

ElseIf Image1(5).Tag = "c" And Image1(6).Tag = "c" And Image1(4).Tag <> "p" Then
Image1(4).Picture = ImageList1.ListImages("ring").Picture
Image1(4).Tag = "c"
CheckWin = True
A = 4
B = 5
C = 6

ElseIf Image1(7).Tag = "c" And Image1(8).Tag = "c" And Image1(9).Tag <> "p" Then
Image1(9).Picture = ImageList1.ListImages("ring").Picture
Image1(9).Tag = "c"
CheckWin = True
A = 7
B = 8
C = 9

ElseIf Image1(7).Tag = "c" And Image1(9).Tag = "c" And Image1(8).Tag <> "p" Then
Image1(8).Picture = ImageList1.ListImages("ring").Picture
Image1(8).Tag = "c"
CheckWin = True
A = 7
B = 8
C = 9

ElseIf Image1(8).Tag = "c" And Image1(9).Tag = "c" And Image1(7).Tag <> "p" Then
Image1(7).Picture = ImageList1.ListImages("ring").Picture
Image1(7).Tag = "c"
CheckWin = True
A = 7
B = 8
C = 9

Else: CheckWin = False

End If

End Function

Private Sub TakeTurn()
On Error Resume Next
Randomize
If Difficulty = 1 Then
Random

ElseIf Difficulty = 2 Then

If Rnd > 0.2 Then
If Not Defend Then
If Not Smart Then
Random
End If
End If
Else: Random
End If

ElseIf Difficulty = 3 Then

If Not Defend Then
If Not Smart Then
Random
End If
End If
End If
End Sub



Private Sub ClearAll()
On Error Resume Next
Dim i As Integer

For i = 1 To 9
Image1(i).Picture = LoadPicture("")
Image1(i).Tag = ""
Image1(i).Visible = True
Next i
Allowed = True
End Sub
Private Function CheckUserWin() As Boolean
On Error Resume Next
If Image1(1).Tag = "p" And Image1(2).Tag = "p" And Image1(3).Tag = "p" Then
CheckUserWin = True
A = 1
B = 2
C = 3

ElseIf Image1(1).Tag = "p" And Image1(4).Tag = "p" And Image1(7).Tag = "p" Then
CheckUserWin = True
A = 1
B = 4
C = 7

ElseIf Image1(1).Tag = "p" And Image1(5).Tag = "p" And Image1(9).Tag = "p" Then
CheckUserWin = True
A = 1
B = 5
C = 9

ElseIf Image1(2).Tag = "p" And Image1(5).Tag = "p" And Image1(8).Tag = "p" Then
CheckUserWin = True
A = 2
B = 5
C = 8

ElseIf Image1(3).Tag = "p" And Image1(6).Tag = "p" And Image1(9).Tag = "p" Then
CheckUserWin = True
A = 3
B = 6
C = 9

ElseIf Image1(3).Tag = "p" And Image1(5).Tag = "p" And Image1(7).Tag = "p" Then
CheckUserWin = True
A = 3
B = 5
C = 7

ElseIf Image1(4).Tag = "p" And Image1(5).Tag = "p" And Image1(6).Tag = "p" Then
CheckUserWin = True
A = 4
B = 5
C = 6

ElseIf Image1(7).Tag = "p" And Image1(8).Tag = "p" And Image1(9).Tag = "p" Then
CheckUserWin = True
A = 7
B = 8
C = 9
Else: CheckUserWin = False

End If

End Function

Private Function CheckDraw() As Boolean
On Error Resume Next
Dim i As Integer, count As Integer

For i = 1 To 9
DoEvents
If Image1(i).Tag = "" Then
CheckDraw = False
Exit Function
End If
Next i
CheckDraw = True
End Function
Private Sub blink()
On Error Resume Next
Dim i As Integer

For i = 1 To 16
If Image1(A).Visible = True And Image1(B).Visible = True And Image1(C).Visible = True Then
Image1(A).Visible = False
Image1(B).Visible = False
Image1(C).Visible = False
DoEvents
Else
Image1(A).Visible = True
Image1(B).Visible = True
Image1(C).Visible = True
End If
DoEvents
Sleep (40)
Next i
End Sub
Private Function Defend() As Boolean
On Error Resume Next

If Image1(1).Tag = "p" And Image1(2).Tag = "p" And Image1(3).Tag = "" Then
Image1(3).Picture = ImageList1.ListImages("ring").Picture
Image1(3).Tag = "c"
Defend = True

ElseIf Image1(1).Tag = "p" And Image1(3).Tag = "p" And Image1(2).Tag = "" Then
Image1(2).Picture = ImageList1.ListImages("ring").Picture
Image1(2).Tag = "c"
Defend = True

ElseIf Image1(2).Tag = "p" And Image1(3).Tag = "p" And Image1(1).Tag = "" Then
Image1(1).Picture = ImageList1.ListImages("ring").Picture
Image1(1).Tag = "c"
Defend = True

ElseIf Image1(1).Tag = "p" And Image1(4).Tag = "p" And Image1(7).Tag = "" Then
Image1(7).Picture = ImageList1.ListImages("ring").Picture
Image1(7).Tag = "c"
Defend = True

ElseIf Image1(1).Tag = "p" And Image1(7).Tag = "p" And Image1(4).Tag = "" Then
Image1(4).Picture = ImageList1.ListImages("ring").Picture
Image1(4).Tag = "c"
Defend = True

ElseIf Image1(4).Tag = "p" And Image1(7).Tag = "p" And Image1(1).Tag = "" Then
Image1(1).Picture = ImageList1.ListImages("ring").Picture
Image1(1).Tag = "c"
Defend = True

ElseIf Image1(1).Tag = "p" And Image1(5).Tag = "p" And Image1(9).Tag = "" Then
Image1(9).Picture = ImageList1.ListImages("ring").Picture
Image1(9).Tag = "c"
Defend = True

ElseIf Image1(1).Tag = "p" And Image1(9).Tag = "p" And Image1(5).Tag = "" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Defend = True

ElseIf Image1(5).Tag = "p" And Image1(9).Tag = "p" And Image1(1).Tag = "" Then
Image1(1).Picture = ImageList1.ListImages("ring").Picture
Image1(1).Tag = "c"
Defend = True

ElseIf Image1(2).Tag = "p" And Image1(5).Tag = "p" And Image1(8).Tag = "" Then
Image1(8).Picture = ImageList1.ListImages("ring").Picture
Image1(8).Tag = "c"
Defend = True

ElseIf Image1(2).Tag = "p" And Image1(8).Tag = "p" And Image1(5).Tag = "" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Defend = True

ElseIf Image1(5).Tag = "p" And Image1(8).Tag = "p" And Image1(2).Tag = "" Then
Image1(2).Picture = ImageList1.ListImages("ring").Picture
Image1(2).Tag = "c"
Defend = True

ElseIf Image1(3).Tag = "p" And Image1(6).Tag = "p" And Image1(9).Tag = "" Then
Image1(9).Picture = ImageList1.ListImages("ring").Picture
Image1(9).Tag = "c"
Defend = True

ElseIf Image1(3).Tag = "p" And Image1(9).Tag = "p" And Image1(6).Tag = "" Then
Image1(6).Picture = ImageList1.ListImages("ring").Picture
Image1(6).Tag = "c"
Defend = True

ElseIf Image1(6).Tag = "p" And Image1(9).Tag = "p" And Image1(3).Tag = "" Then
Image1(3).Picture = ImageList1.ListImages("ring").Picture
Image1(3).Tag = "c"
Defend = True

ElseIf Image1(3).Tag = "p" And Image1(5).Tag = "p" And Image1(7).Tag = "" Then
Image1(7).Picture = ImageList1.ListImages("ring").Picture
Image1(7).Tag = "c"
Defend = True

ElseIf Image1(3).Tag = "p" And Image1(7).Tag = "p" And Image1(5).Tag = "" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Defend = True

ElseIf Image1(5).Tag = "p" And Image1(7).Tag = "p" And Image1(3).Tag = "" Then
Image1(3).Picture = ImageList1.ListImages("ring").Picture
Image1(3).Tag = "c"
Defend = True

ElseIf Image1(4).Tag = "p" And Image1(5).Tag = "p" And Image1(6).Tag = "" Then
Image1(6).Picture = ImageList1.ListImages("ring").Picture
Image1(6).Tag = "c"
Defend = True

ElseIf Image1(4).Tag = "p" And Image1(6).Tag = "p" And Image1(5).Tag = "" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Defend = True

ElseIf Image1(5).Tag = "p" And Image1(6).Tag = "p" And Image1(4).Tag = "" Then
Image1(4).Picture = ImageList1.ListImages("ring").Picture
Image1(4).Tag = "c"
Defend = True

ElseIf Image1(7).Tag = "p" And Image1(8).Tag = "p" And Image1(9).Tag = "" Then
Image1(9).Picture = ImageList1.ListImages("ring").Picture
Image1(9).Tag = "c"
Defend = True

ElseIf Image1(7).Tag = "p" And Image1(9).Tag = "p" And Image1(8).Tag = "" Then
Image1(8).Picture = ImageList1.ListImages("ring").Picture
Image1(8).Tag = "c"
Defend = True

ElseIf Image1(8).Tag = "p" And Image1(9).Tag = "p" And Image1(7).Tag = "" Then
Image1(7).Picture = ImageList1.ListImages("ring").Picture
Image1(7).Tag = "c"
Defend = True

Else: Defend = False
End If
End Function

Private Function Smart() As Boolean
On Error Resume Next

Dim Index As Integer

If (Image1(1).Tag = "p" And Image1(5).Tag = "p") And Image1(7).Tag = "" Then
Image1(7).Picture = ImageList1.ListImages("ring").Picture
Image1(7).Tag = "c"
Smart = True

ElseIf (Image1(7).Tag = "p" And Image1(5).Tag = "p") And Image1(1).Tag = "" Then
Image1(1).Picture = ImageList1.ListImages("ring").Picture
Image1(1).Tag = "c"
Smart = True

ElseIf (Image1(3).Tag = "p" And Image1(5).Tag = "p") And Image1(9).Tag = "" Then
Image1(9).Picture = ImageList1.ListImages("ring").Picture
Image1(9).Tag = "c"
Smart = True

ElseIf (Image1(9).Tag = "p" And Image1(5).Tag = "p") And Image1(3).Tag = "" Then
Image1(3).Picture = ImageList1.ListImages("ring").Picture
Image1(3).Tag = "c"
Smart = True

ElseIf Image1(2).Tag = "" And Image1(3).Tag = "" And Image1(4).Tag = "" And Image1(5).Tag = "" And Image1(6).Tag = "" And Image1(7).Tag = "" And Image1(8).Tag = "" And Image1(9).Tag = "" And Image1(1).Tag = "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Smart = True

ElseIf Image1(1).Tag = "" And Image1(2).Tag = "" And Image1(4).Tag = "" And Image1(5).Tag = "" And Image1(6).Tag = "" And Image1(7).Tag = "" And Image1(8).Tag = "" And Image1(9).Tag = "" And Image1(3).Tag = "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Smart = True

ElseIf Image1(2).Tag = "" And Image1(3).Tag = "" And Image1(4).Tag = "" And Image1(5).Tag = "" And Image1(6).Tag = "" And Image1(1).Tag = "" And Image1(8).Tag = "" And Image1(9).Tag = "" And Image1(7).Tag = "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Smart = True

ElseIf Image1(2).Tag = "" And Image1(3).Tag = "" And Image1(4).Tag = "" And Image1(5).Tag = "" And Image1(6).Tag = "" And Image1(7).Tag = "" And Image1(8).Tag = "" And Image1(1).Tag = "" And Image1(9).Tag = "p" Then
Image1(5).Picture = ImageList1.ListImages("ring").Picture
Image1(5).Tag = "c"
Smart = True


ElseIf Image1(2).Tag = "p" And Image1(1).Tag = "" And Image1(9).Tag = "" And Image1(3).Tag = "" And Image1(4).Tag = "" And Image1(5).Tag = "" And Image1(6).Tag = "" And Image1(7).Tag = "" And Image1(8).Tag = "" Then
Again4:
Index = 1 + 2 * Rnd
If Index = 1 Or Index = 3 Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Smart = True
Else: GoTo Again4
End If

ElseIf Image1(4).Tag = "p" And Image1(1).Tag = "" And Image1(2).Tag = "" And Image1(3).Tag = "" And Image1(5).Tag = "" And Image1(6).Tag = "" And Image1(7).Tag = "" And Image1(8).Tag = "" And Image1(9).Tag = "" Then
Again5:
Index = 1 + 6 * Rnd
If Index = 1 Or Index = 7 Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Smart = True
Else: GoTo Again5
End If

ElseIf Image1(6).Tag = "p" And Image1(1).Tag = "" And Image1(2).Tag = "" And Image1(3).Tag = "" And Image1(5).Tag = "" And Image1(4).Tag = "" And Image1(7).Tag = "" And Image1(8).Tag = "" And Image1(9).Tag = "" Then
Again6:
Index = 3 + 6 * Rnd
If Index = 3 Or Index = 9 Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Smart = True
Else: GoTo Again6
End If

ElseIf Image1(8).Tag = "p" And Image1(1).Tag = "" And Image1(2).Tag = "" And Image1(3).Tag = "" And Image1(5).Tag = "" And Image1(6).Tag = "" And Image1(7).Tag = "" And Image1(4).Tag = "" And Image1(9).Tag = "" Then
Again7:
Index = 7 + 2 * Rnd
If Index = 7 Or Index = 9 Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Smart = True
Else: GoTo Again7
End If


ElseIf (Image1(1).Tag = "p" And Image1(9).Tag = "p") And (Image1(2).Tag <> "p" And Image1(3).Tag <> "p" And Image1(4).Tag <> "p" And Image1(5).Tag <> "p" And Image1(6).Tag <> "p" And Image1(7).Tag <> "p" And Image1(8).Tag <> "p") Then
Again:
Index = 2 + 6 * Rnd
If Index = 2 Or Index = 4 Or Index = 6 Or Index = 8 Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Smart = True
Else: GoTo Again
End If

ElseIf (Image1(7).Tag = "p" And Image1(3).Tag = "p") And (Image1(1).Tag <> "p" And Image1(2).Tag <> "p" And Image1(4).Tag <> "p" And Image1(5).Tag <> "p" And Image1(6).Tag <> "p" And Image1(8).Tag <> "p" And Image1(9).Tag <> "p") Then
Again2:
Index = 2 + 6 * Rnd
If Index = 2 Or Index = 4 Or Index = 6 Or Index = 8 Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Smart = True
Else: GoTo Again2
End If

ElseIf Image1(2).Tag = "" And Image1(3).Tag = "" And Image1(4).Tag = "" And Image1(1).Tag = "" And Image1(6).Tag = "" And Image1(7).Tag = "" And Image1(8).Tag = "" And Image1(9).Tag = "" And Image1(5).Tag = "p" Then
Again3:
Index = 1 + 8 * Rnd
If Index = 1 Or Index = 3 Or Index = 7 Or Index = 9 Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Smart = True
Else: GoTo Again3
End If



Else: Smart = False

End If
End Function
Private Sub Random()
On Error Resume Next
Dim Index As Integer, i As Integer
Randomize

For i = 1 To 20
Index = 8 * Rnd + 1
If Image1(Index).Tag = "" Then
Image1(Index).Picture = ImageList1.ListImages("ring").Picture
Image1(Index).Tag = "c"
Exit For
End If
Next i

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image3.Picture = ImageList1.ListImages(6).Picture
End If
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image3.Picture = ImageList1.ListImages(5).Picture
If X > 0 And X < 915 And Y > 0 And Y < 315 Then
Form3.Text1 = Label1(0)
Form3.Text2 = Label1(1)
Form3.Option1(Difficulty).Value = True
Form3.Show vbModal
End If
End If
End Sub

