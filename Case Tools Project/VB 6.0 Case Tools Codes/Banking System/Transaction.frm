VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Transaction 
   Caption         =   "Transaction"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15630
   LinkTopic       =   "Form2"
   Picture         =   "Transaction.frx":0000
   ScaleHeight     =   9465
   ScaleWidth      =   15630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc transactionado 
      Height          =   735
      Left            =   12480
      Top             =   6480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Transaction.frx":0342
      OLEDBString     =   $"Transaction.frx":03CB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "traacc"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtaccno 
      DataField       =   "Account No"
      DataSource      =   "transactionado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txttransamt 
      DataField       =   "Transaction Amount"
      DataSource      =   "transactionado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   5
      Top             =   7800
      Width           =   2775
   End
   Begin VB.TextBox txtminbal 
      DataField       =   "Minimum Balance"
      DataSource      =   "transactionado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   4
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox txtcurbal 
      DataField       =   "Current Balance"
      DataSource      =   "transactionado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   3
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox txtrecaccno 
      DataField       =   "Reciver Account No"
      DataSource      =   "transactionado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   2
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Account no"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Reciver Account no"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Transaction Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Minimum Balance"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   6600
      Width           =   2415
   End
End
Attribute VB_Name = "Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
transactionado.Recordset.Fields("Account no") = txtaccno.Text
transactionado.Recordset.Fields("Reciver Account no") = txtrecaccno.Text
transactionado.Recordset.Fields("Current Balance") = txtcurbal.Text
transactionado.Recordset.Fields("Minimum Balance") = txtminbal.Text
transactionado.Recordset.Fields("transaction Amount") = txttransamt.Text
transactionado.Recordset.Update
MsgBox "User transaction Successful"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
txtaccno.Text = ""
txtrecaccno.Text = ""
txtcurbal.Text = ""
txtminbal.Text = ""
txttransamt.Text = ""

End Sub

Private Sub Form_Load()
transactionado.Recordset.AddNew
End Sub
