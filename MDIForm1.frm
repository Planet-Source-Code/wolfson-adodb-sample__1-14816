VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   "Simple ADODB example     by WolFSON...."
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8415
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   1  'Align Top
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Max             =   15000
   End
   Begin VB.Menu mnc 
      Caption         =   "&Customer Info"
      Index           =   0
   End
   Begin VB.Menu mnas 
      Caption         =   "&About"
      Index           =   1
   End
   Begin VB.Menu mne 
      Caption         =   "&Exit"
      Index           =   2
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Form1.Top = (MDIForm1.ScaleHeight - Form1.Height) / 2
    Form1.Left = (MDIForm1.ScaleWidth - Form1.Width) / 2
    MDIForm1.mnc(0).Enabled = False
    MDIForm1.mnas(1).Enabled = False
    MDIForm1.mne(2).Enabled = False
End Sub
Private Sub mnas_Click(Index As Integer)
    Form3.Top = (MDIForm1.ScaleHeight - Form3.Height) / 2
    Form3.Left = (MDIForm1.ScaleWidth - Form3.Width) / 2
End Sub
Private Sub mnc_Click(Index As Integer)
    Form2.Show
End Sub
Private Sub mne_Click(Index As Integer)
    End
End Sub
