VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   3030
      TabIndex        =   0
      Top             =   2220
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "sample"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "       Designed by WolFSON (The TURK)                  Have fun in learning VB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2100
      TabIndex        =   2
      Top             =   1500
      Width           =   3045
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
      Left            =   2310
      TabIndex        =   1
      Top             =   150
      Width           =   2640
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2685
      Left            =   0
      Picture         =   "about.frx":0000
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2100
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
