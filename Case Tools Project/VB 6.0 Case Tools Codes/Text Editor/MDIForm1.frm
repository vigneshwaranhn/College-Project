VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Form"
   ClientHeight    =   8340
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14805
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   3000
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Begin VB.Menu mnuTxtColor 
         Caption         =   "Text Color"
      End
      Begin VB.Menu mnuBackColor 
         Caption         =   "Background Color"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuTxtColor_Click()
cd.ShowColor
frmchild.TextBox.ForeColor = cd.Color
End Sub


Private Sub mnuBackColor_Click()
cd.ShowColor
frmchild.TextBox.BackColor = cd.Color
End Sub


Private Sub mnuFont_Click()
cd.Flags = cd1CFBoth
cd.ShowFont
frmchild.TextBox.Font = cd.FontName
frmchild.TextBox.FontItalic = cd.FontItalic
frmchild.TextBox.FontBold = cd.FontBold
frmchild.TextBox.FontSize = cd.FontSize
End Sub


Private Sub mnuExit_Click()
End
End Sub

