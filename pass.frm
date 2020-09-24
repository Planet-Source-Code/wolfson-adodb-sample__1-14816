VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8760
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   300
      Left            =   3720
      TabIndex        =   4
      Top             =   2265
      Width           =   1260
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3030
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1485
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   3030
      TabIndex        =   2
      Top             =   1140
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   5435
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   4815
         Left            =   1680
         Picture         =   "pass.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Veritabaný Oluþturuluyor....."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   5400
         Width           =   3375
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   765
      Left            =   55435
      TabIndex        =   11
      Top             =   1650
      Width           =   840
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2685
      Left            =   0
      Picture         =   "pass.frx":3C024
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2100
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   2280
      TabIndex        =   10
      Top             =   1890
      Width           =   1380
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   3720
      TabIndex        =   9
      Top             =   1890
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2160
      TabIndex        =   8
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   2160
      TabIndex        =   7
      Top             =   1185
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "ADODB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   870
      Left            =   2250
      TabIndex        =   6
      Top             =   120
      Width           =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub Command1_Click()
Dim say As Integer
Dim aaa As Integer
If Text2.Text = "ADODB" Then
    MDIForm1.ProgressBar1.Visible = True
    MDIForm1.ProgressBar1.Min = 0
    MDIForm1.ProgressBar1.Max = 15000
    For say = 1 To 15000
        MDIForm1.ProgressBar1.Value = say
        say = say + 1
        If say = 9800 Then Unload Me
    Next
    MDIForm1.mnc(0).Enabled = True
    MDIForm1.mnas(1).Enabled = True
    MDIForm1.mne(2).Enabled = True
    MDIForm1.ProgressBar1.Visible = False
Else
    If aaa = 3 Then
        MsgBox "Deleting Password !.. " & Chr(13) & "You are not an authorized user..", vbOKOnly + vbCritical, "Warning  !!!!"
        MDIForm1.ProgressBar1.Visible = True
        MDIForm1.ProgressBar1.Min = 0
        MDIForm1.ProgressBar1.Max = 15000
        For say = 1 To 15000
            MDIForm1.ProgressBar1.Value = say
            say = say + 1
        Next
        MDIForm1.ProgressBar1.Visible = False
        Unload Me
        Exit Sub
    End If
    If MsgBox("Wrong Password!.. " & Chr(13) & "Check the CapsLock Button..", vbOKOnly + vbCritical, "Wrong Password !!!") Then
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
    End If
End If
End Sub
Private Sub Command2_Click()
    End
End Sub
Private Sub Form_Load()
    Dim strComputerName As String * 255
    GetComputerName strComputerName, 255
    Text1 = strComputerName
    Label8.Caption = " " & Date
    Timer1.Enabled = True
    Timer1.Interval = 100
End Sub
Private Sub Timer1_Timer()
    Label7 = " " & Time
End Sub

