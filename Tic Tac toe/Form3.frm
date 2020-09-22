VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   4680
   ClientLeft      =   4455
   ClientTop       =   3690
   ClientWidth     =   4680
   FillColor       =   &H80000012&
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2070
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   35
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":3822
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":3F26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000040&
      Height          =   195
      Index           =   2
      Left            =   2161
      TabIndex        =   8
      Top             =   1260
      UseMaskColor    =   -1  'True
      Width           =   225
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000040&
      Height          =   195
      Index           =   3
      Left            =   2997
      TabIndex        =   7
      Top             =   1260
      UseMaskColor    =   -1  'True
      Width           =   225
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Left            =   1530
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   2955
      UseMaskColor    =   -1  'True
      Width           =   195
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2941
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "Comp"
      Top             =   2040
      Width           =   945
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   781
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "Player1"
      Top             =   2040
      Width           =   945
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000040&
      Height          =   195
      Index           =   1
      Left            =   1325
      TabIndex        =   0
      Top             =   1260
      UseMaskColor    =   -1  'True
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Always On Top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   1815
      TabIndex        =   9
      Top             =   2925
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P2 Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Index           =   1
      Left            =   2933
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P1 Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Index           =   0
      Left            =   773
      TabIndex        =   4
      Top             =   1680
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   2078
      Picture         =   "Form3.frx":462A
      Stretch         =   -1  'True
      Top             =   3690
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Easy     Normal    Hard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   1170
      TabIndex        =   1
      Top             =   840
      Width           =   2235
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'settigs form
'basically an elliptic region is applied to this form
'ok button is again an image control as in the main form
Option Explicit
Dim Rgn As Long

Private Sub Check1_Click()
If Check1.Value = Checked Then
SetWindowPos Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
ElseIf Check1.Value = Unchecked Then
SetWindowPos Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End If
End Sub

Private Sub Form_Load()
Rgn = CreateEllipticRgn(0, 0, 312, 312)
SetWindowRgn hWnd, Rgn, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage hWnd, WM_NCLBUTTONDOWN, 2, 0&
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteObject Rgn
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image1.Picture = ImageList1.ListImages(2).Picture
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image1.Picture = ImageList1.ListImages(1).Picture
If (X > 0 And X < 525) And (Y > 0 And Y < 450) Then
Me.Hide
End If
End If



End Sub

Private Sub Option1_Click(Index As Integer)
Form1.Difficulty = Index
End Sub

Private Sub Text1_Change()
Form1.Label1(0) = Text1
End Sub

Private Sub Text2_Change()
Form1.Label1(1) = Text2
End Sub
