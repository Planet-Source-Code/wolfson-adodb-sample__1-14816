VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11685
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "main.frx":0442
      Height          =   5235
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9234
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "NAME"
         Caption         =   "NAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SURNAME"
         Caption         =   "SURNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "TEL1"
         Caption         =   "TELEPHONE 1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "TEL2"
         Caption         =   "TELEPHONE 2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "FAX"
         Caption         =   "FAX"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ADRESS"
         Caption         =   "ADRESS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "KNOWLEDGE"
         Caption         =   "KNOWLEDGE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
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
      Height          =   315
      Left            =   1785
      TabIndex        =   0
      Top             =   630
      Width           =   4020
   End
   Begin VB.TextBox Text3 
      DataField       =   "TEL1"
      DataSource      =   "Adodc1"
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
      Height          =   315
      Left            =   1785
      TabIndex        =   2
      Top             =   1275
      Width           =   4020
   End
   Begin VB.TextBox Text7 
      DataField       =   "KNOWLEDGE"
      DataSource      =   "Adodc1"
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
      Height          =   645
      Left            =   7605
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1590
      Width           =   4020
   End
   Begin VB.TextBox Text6 
      DataField       =   "ADRESS"
      DataSource      =   "Adodc1"
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
      Height          =   975
      Left            =   7590
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   615
      Width           =   4020
   End
   Begin VB.TextBox Text5 
      DataField       =   "FAX"
      DataSource      =   "Adodc1"
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
      Height          =   315
      Left            =   1785
      TabIndex        =   4
      Top             =   1920
      Width           =   4020
   End
   Begin VB.TextBox Text4 
      DataField       =   "TEL2"
      DataSource      =   "Adodc1"
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
      Height          =   315
      Left            =   1785
      TabIndex        =   3
      Top             =   1590
      Width           =   4020
   End
   Begin VB.TextBox Text2 
      DataField       =   "SURNAME"
      DataSource      =   "Adodc1"
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
      Height          =   315
      Left            =   1785
      TabIndex        =   1
      Top             =   945
      Width           =   4020
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10320
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   6
      Top             =   7545
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Picture         =   "main.frx":0457
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Picture         =   "main.frx":05F1
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Picture         =   "main.frx":078B
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Picture         =   "main.frx":0925
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Picture         =   "main.frx":0ABF
            TextSave        =   "30.01.2001"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Picture         =   "main.frx":0C59
            TextSave        =   "09:08"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   11520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0DF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1335
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1787
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1BD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":202B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":247D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2AD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2BE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2CF5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1111
      ButtonWidth     =   1270
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "yeni"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "ekle"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "sil"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "yaz"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "First"
            Key             =   "ilk"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Previous"
            Key             =   "once"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Key             =   "sonra"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Last"
            Key             =   "son"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "sorgu"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "cik"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SURNAME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   0
      TabIndex        =   15
      Top             =   945
      Width           =   1770
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TELEPHONE 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   0
      TabIndex        =   14
      Top             =   1275
      Width           =   1770
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " IMPORTANT            KNOWLEDGE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   4
      Left            =   5820
      TabIndex        =   13
      Top             =   1590
      Width           =   1770
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ADRESS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   5820
      TabIndex        =   12
      Top             =   615
      Width           =   1770
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FAX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   1770
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TELEPHONE 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   1590
      Width           =   1770
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NAME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   630
      Width           =   1770
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SorSor As Boolean
Private Sub Form_Load()
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\adodb1.mdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Table1"
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
    Me.Caption = "ADODB SAMPLE[ Record Count : " & Adodc1.Recordset.RecordCount & "]"
    Set DataGrid1.DataSource = Adodc1
    SorSor = False
End Sub
Private Sub Form_Resize()
On Error GoTo aa:
DataGrid1.Left = 0
DataGrid1.Top = 2280
DataGrid1.Width = Me.ScaleWidth
DataGrid1.Height = Me.ScaleHeight - (Text1.Height + Text2.Height + Text3.Height + Text4.Height + Text5.Height + StatusBar2.Height + Toolbar1.Height + 25)
Exit Sub
aa:
Exit Sub
End Sub
Private Sub Text1_Change()
If SorSor = True Then
    Adodc1.RecordSource = "Select * From Table1 Where NAME LIKE '" & Text1 & "%';"
    Adodc1.Refresh
    If Text1 = "" Then
        Text1 = "": Text2 = "": Text3 = ""
        Text4 = "": Text5 = "": Text6 = ""
        Text7 = ""
    End If
End If
End Sub

Private Sub Text2_Change()
If SorSor = True Then
    Adodc1.RecordSource = "Select * From Table1 Where SURNAME LIKE '" & Text2 & "%';"
    Adodc1.Refresh
    If Text2 = "" Then
        Text1 = "": Text2 = "": Text3 = ""
        Text4 = "": Text5 = "": Text6 = ""
        Text7 = ""
    End If
End If

End Sub

Private Sub Text3_Change()
If SorSor = True Then
    Adodc1.RecordSource = "Select * From Table1 Where TEL1 LIKE '" & Text3 & "%';"
    Adodc1.Refresh
    If Text3 = "" Then
        Text1 = "": Text2 = "": Text3 = ""
        Text4 = "": Text5 = "": Text6 = ""
        Text7 = ""
    End If
End If

End Sub

Private Sub Text4_Change()
If SorSor = True Then
    Adodc1.RecordSource = "Select * From Table1 Where TEL2 LIKE '" & Text4 & "%';"
    Adodc1.Refresh
    If Text4 = "" Then
        Text1 = "": Text2 = "": Text3 = ""
        Text4 = "": Text5 = "": Text6 = ""
        Text7 = ""
    End If
End If
End Sub

Private Sub Text5_Change()
If SorSor = True Then
    Adodc1.RecordSource = "Select * From Table1 Where FAX LIKE '" & Text5 & "%';"
    Adodc1.Refresh
    If Text5 = "" Then
        Text1 = "": Text2 = "": Text3 = ""
        Text4 = "": Text5 = "": Text6 = ""
        Text7 = ""
    End If
End If

End Sub

Private Sub Text6_Change()
If SorSor = True Then
    Adodc1.RecordSource = "Select * From Table1 Where ADRESS LIKE '" & Text6 & "%';"
    Adodc1.Refresh
    If Text6 = "" Then
        Text1 = "": Text2 = "": Text3 = ""
        Text4 = "": Text5 = "": Text6 = ""
        Text7 = ""
    End If
End If

End Sub

Private Sub Text7_Change()
If SorSor = True Then
    Adodc1.RecordSource = "Select * From Table1 Where KNOWLEDGE LIKE '" & Text7 & "%';"
    Adodc1.Refresh
    If Text7 = "" Then
        Text1 = "": Text2 = "": Text3 = ""
        Text4 = "": Text5 = "": Text6 = ""
        Text7 = ""
    End If
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
            Case "yeni"
                If Toolbar1.Buttons.Item(1).Caption = "New" Then
                    Toolbar1.Buttons.Item(4).Enabled = False
                    Text1.Enabled = 1: Text2.Enabled = 1: Text3.Enabled = 1
                    Text4.Enabled = 1: Text5.Enabled = 1: Text6.Enabled = 1
                    Text7.Enabled = 1
                    Adodc1.Recordset.AddNew
                    Toolbar1.Buttons.Item(1).Caption = "Cancel"
                    Text1.SetFocus
                Else
                    Adodc1.Refresh
                    Toolbar1.Buttons.Item(4).Enabled = True
                    Text1.Enabled = 0: Text2.Enabled = 0: Text3.Enabled = 0
                    Text4.Enabled = 0: Text5.Enabled = 0: Text6.Enabled = 0
                    Text7.Enabled = 0
                    Toolbar1.Buttons.Item(1).Caption = "New"
                End If
            Case "ekle"
                If Not Len(Text1.Text) < 1 Then
                    Adodc1.Recordset.Update
                    Text1.Enabled = 0: Text2.Enabled = 0: Text3.Enabled = 0
                    Text4.Enabled = 0: Text5.Enabled = 0: Text6.Enabled = 0
                    Text7.Enabled = 0
                    Toolbar1.Buttons.Item(1).Caption = "New"
                    Toolbar1.Buttons.Item(4).Enabled = True
                    Me.Caption = "ADODB SAMPLE[ Record Count : " & Adodc1.Recordset.RecordCount & "]"
                Else
                    MsgBox "Mising Data Entry", vbApplicationModal + vbOKOnly + vbCritical, "ERROR !"
                    On Local Error Resume Next
                    Text1.SetFocus
                End If
            Case "sil"
                    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
                    On Local Error Resume Next
                    Adodc1.Recordset.Delete
                    Me.Caption = "ADODB SAMPLE[ Record Count : " & Adodc1.Recordset.RecordCount & "]"
            Case "sorgu"
                If Toolbar1.Buttons.Item(14).Caption = "Search" Then
                    Toolbar1.Buttons.Item(1).Enabled = False
                    Toolbar1.Buttons.Item(2).Enabled = False
                    Toolbar1.Buttons.Item(4).Enabled = False
                    Toolbar1.Buttons.Item(6).Enabled = False
                    Toolbar1.Buttons.Item(8).Enabled = False
                    Toolbar1.Buttons.Item(9).Enabled = False
                    Toolbar1.Buttons.Item(10).Enabled = False
                    Toolbar1.Buttons.Item(11).Enabled = False
                    Toolbar1.Buttons.Item(14).Caption = "Cancel"
                    Text1.Enabled = 1: Text2.Enabled = 1: Text3.Enabled = 1
                    Text4.Enabled = 1: Text5.Enabled = 1: Text6.Enabled = 1
                    Text7.Enabled = 1
                    DataGrid1.AllowAddNew = False
                    Text1.SetFocus
                    Text1.DataField = "": Text1 = ""
                    Text2.DataField = "": Text2 = ""
                    Text3.DataField = "": Text3 = ""
                    Text4.DataField = "": Text4 = ""
                    Text5.DataField = "": Text5 = ""
                    Text6.DataField = "": Text6 = ""
                    Text7.DataField = "": Text7 = ""
                    SorSor = True
                Else
                    SorSor = False
                    Toolbar1.Buttons.Item(1).Enabled = True
                    Toolbar1.Buttons.Item(2).Enabled = True
                    Toolbar1.Buttons.Item(4).Enabled = True
                    Toolbar1.Buttons.Item(6).Enabled = True
                    Toolbar1.Buttons.Item(8).Enabled = True
                    Toolbar1.Buttons.Item(9).Enabled = True
                    Toolbar1.Buttons.Item(10).Enabled = True
                    Toolbar1.Buttons.Item(11).Enabled = True
                    Toolbar1.Buttons.Item(14).Caption = "Search"
                    Text1.Enabled = 0: Text2.Enabled = 0: Text3.Enabled = 0
                    Text4.Enabled = 0: Text5.Enabled = 0: Text6.Enabled = 0
                    Text7.Enabled = 0
                    DataGrid1.AllowAddNew = True
                    Text1.DataField = "NAME"
                    Text2.DataField = "SURNAME"
                    Text3.DataField = "TEL1"
                    Text4.DataField = "TEL2"
                    Text5.DataField = "FAX"
                    Text6.DataField = "ADRESS"
                    Text7.DataField = "KNOWLEDGE"
                    Adodc1.RecordSource = "Select * From Table1"
                    Adodc1.Refresh
                End If
            Case "cik"
                    Unload Me
            Case "ilk"
                    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
                    If Not Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
            Case "once"
                    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
                    If Not Adodc1.Recordset.BOF Then Adodc1.Recordset.MovePrevious
            Case "sonra"
                    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
                    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveNext
            Case "son"
                    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
                    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
            Case "yaz"
                   ' CrystalReport1.ReportFileName = App.Path & "\report.rpt"
                   ' CrystalReport1.RetrieveDataFiles
                   ' CrystalReport1.WindowState = crptMaximized
                   ' CrystalReport1.Action = 1
    End Select
End Sub

