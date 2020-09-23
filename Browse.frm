VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4005
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   4425
   LinkTopic       =   "Form2"
   ScaleHeight     =   4005
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtFilename 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2280
      Pattern         =   "*.EXE"
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EXE Filename"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim str As String
str = txtFilename.Text

Dim length As Integer
length = Len(str)

Dim str2 As String
str2 = Mid(str, 1, length - 4)

'The idea here is to extract the filename minus
'the .exe part
Form1.txtExeFilename = str2
Me.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
txtFilename.Text = File1.FileName
End Sub
