VERSION 5.00
Begin VB.Form CaoDatAce 
   Caption         =   "A-CE"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdso0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   23
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdphepam 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   22
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdcham 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   21
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdngoac2 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   20
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdngoac1 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   19
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13800
      TabIndex        =   18
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdchia 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdnhan 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdcong 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   15
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdtru 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   14
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdac 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   13
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdenter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdso5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdso9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdso8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdso7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdso6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdso4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdso3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   5
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdso2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdso1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdkhung 
      Height          =   6615
      Left            =   4320
      TabIndex        =   1
      Top             =   2520
      Width           =   6135
   End
   Begin VB.TextBox txtHienThi 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   0
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Label lbl1 
      Caption         =   "Phan Mem May Tinh Bo Tui Cao Dat A-CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   9015
   End
End
Attribute VB_Name = "CaoDatAce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Double, B As Double, KQ As Double
Dim PhepTinh As Integer

Private Sub cmdac_Click()
txtHienThi.Text = ""
End Sub

Private Sub cmdcham_Click()
txtHienThi.Text = txtHienThi.Text & "."
End Sub
' Phep Tinh /
Private Sub cmdchia_Click()
A = txtHienThi.Text
PhepTinh = 4
txtHienThi.Text = ""
End Sub
' Phep Tinh +
Private Sub cmdcong_Click()
A = txtHienThi.Text
PhepTinh = 1
txtHienThi.Text = ""
End Sub

Private Sub cmdenter_Click()
B = txtHienThi.Text
Select Case PhepTinh
Case 1
KQ = A + B
Case 2
KQ = A - B
Case 3
KQ = A * B
Case 4
KQ = A / B
End Select

' KQ = A + B
' KQ = A - B
' KQ = A * B
' Kq = A / B
txtHienThi.Text = KQ
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdngoac1_Click()
txtHienThi.Text = txtHienThi.Text & "("
End Sub

Private Sub cmdngoac2_Click()
txtHienThi.Text = txtHienThi.Text & ")"
End Sub
' Phep Tinh *
Private Sub cmdnhan_Click()
A = txtHienThi.Text
PhepTinh = 3
txtHienThi.Text = ""
End Sub
' Dau am
Private Sub cmdphepam_Click()
txtHienThi.Text = txtHienThi.Text & "-"
End Sub

' Hien thi len text

Private Sub cmdso0_Click()
txtHienThi.Text = txtHienThi.Text & 0
End Sub

Private Sub cmdso1_Click()
txtHienThi.Text = txtHienThi.Text & 1

End Sub

Private Sub cmdso2_Click()
txtHienThi.Text = txtHienThi.Text & 2
End Sub

Private Sub cmdso3_Click()
txtHienThi.Text = txtHienThi.Text & 3
End Sub

Private Sub cmdso4_Click()
txtHienThi.Text = txtHienThi.Text & 4
End Sub

Private Sub cmdso5_Click()
txtHienThi.Text = txtHienThi.Text & 5
End Sub

Private Sub cmdso6_Click()
txtHienThi.Text = txtHienThi.Text & 6
End Sub

Private Sub cmdso7_Click()
txtHienThi.Text = txtHienThi.Text & 7
End Sub

Private Sub cmdso8_Click()
txtHienThi = txtHienThi.Text & 8
End Sub

Private Sub cmdso9_Click()
txtHienThi = txtHienThi.Text & 9
End Sub
' Phep Tinh -
Private Sub cmdtru_Click()
A = txtHienThi.Text
PhepTinh = 2
txtHienThi.Text = ""
End Sub

