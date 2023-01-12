VERSION 5.00
Begin VB.Form ImportArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Arquivo"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   8415
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   4080
      Pattern         =   "*.xls*;*.xlsx*"
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblFiles 
      Caption         =   "Files:"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblDirectories 
      Caption         =   "Directories:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblDriver 
      Caption         =   "Drivers:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "ImportArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    txtPath.Text = File1.Path + "\" + File1.FileName
End Sub

Private Sub File1_DblClick()
    Principal.txtPath.Text = txtPath.Text
    Unload Me
End Sub
