VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   Caption         =   "Using the Registry"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Open ControlPanel -> Sound "
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CheckBox chkDefault 
      BackColor       =   &H80000012&
      Caption         =   "Default"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkProgError 
      BackColor       =   &H80000012&
      Caption         =   "Program Error"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CheckBox chkMPop 
      BackColor       =   &H80000012&
      Caption         =   "Menu Popup"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkMCmd 
      BackColor       =   &H80000012&
      Caption         =   "Menu Command"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox chkExc 
      BackColor       =   &H80000012&
      Caption         =   "Exclamation"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox chkAst 
      BackColor       =   &H80000012&
      Caption         =   "Asterics"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox chkQue 
      BackColor       =   &H80000012&
      Caption         =   "Qustion"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkResDwn 
      BackColor       =   &H80000012&
      Caption         =   "Restore Dn."
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkResUp 
      BackColor       =   &H80000012&
      Caption         =   "Restore Up"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox chkMin 
      BackColor       =   &H80000012&
      Caption         =   "Minimize"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkMax 
      BackColor       =   &H80000012&
      Caption         =   "Maximize"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox chkClose 
      BackColor       =   &H80000012&
      Caption         =   "Close"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox chkOpen 
      BackColor       =   &H80000012&
      Caption         =   "Open"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Browse"
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtExeFilename 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtAppname 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Add"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Help"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   255
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Events To Play:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000E&
      Height          =   2415
      Left            =   2280
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ".EXE  filename"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sound Association"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long




Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_DYN_DATA = &H80000006
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_USERS = &H80000003

Const REG_SZ = 1
Private Sub cmdOK_Click()
Dim retValue As Long
Dim result As Long
Dim keyID As Long
Dim keyValue As String
Dim subKey As String
Dim bufSize As Long
Dim str As String
Dim str2 As String
str = "\.Default\AppEvents\Schemes\Apps\"
str2 = "\.Default\AppEvents\Schemes\Apps\"
''Create key: Here is the text of application name
str = str & txtExeFilename.Text
   retValue = RegCreateKey(HKEY_USERS, str, keyID)
   retValue = RegSetValueEx(keyID, subKey, 0&, REG_SZ, ByVal txtAppname.Text, Len(keyValue) + 1)
   If retValue = 0 Then
        'Create subkeys. Here we work with only one
        'Event:Open program
        If Form1.chkOpen.Value = 1 Then
            subKey = "Open"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            'restore value of str2 so that the next key can be created
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:Close
        If Form1.chkClose.Value = 1 Then
            subKey = "Close"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If

        
        'Event:Maximize
        If Form1.chkMax.Value = 1 Then
            subKey = "Maximize"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:Minimize
        If Form1.chkMin.Value = 1 Then
            subKey = "Minimize"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
    'Event:RestoreUp
        If Form1.chkResUp.Value = 1 Then
            subKey = "RestoreUp"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
    
    'Event:RestoreDown
        If Form1.chkResDwn.Value = 1 Then
            subKey = "RestoreDown"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
    
    'Event:Astericks
        If Form1.chkAst.Value = 1 Then
            subKey = "SystemAsterisk"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:Question
        If Form1.chkQue.Value = 1 Then
            subKey = "SystemQuestion"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:Exclamation
        If Form1.chkExc.Value = 1 Then
            subKey = "SystemExclamation"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:MenuCommand
        If Form1.chkMCmd.Value = 1 Then
            subKey = "MenuCommand"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:MenuPopup
        If Form1.chkMPop.Value = 1 Then
            subKey = "MenuPopup"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:ProgramError
        If Form1.chkProgError.Value = 1 Then
            subKey = "AppGPFault"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
        
        'Event:Default
        If Form1.chkDefault.Value = 1 Then
            subKey = ".Default"
            str2 = str2 & txtExeFilename.Text & "\" & subKey
            retValue = RegCreateKey(HKEY_USERS, str2, keyID)
            str2 = "\.Default\AppEvents\Schemes\Apps\"
        End If
    End If
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1"

End Sub

Private Sub Command4_Click()
Dim retValue As Long
Dim str As String
str = "\.Default\AppEvents\Schemes\Apps\"
str = str & txtExeFilename.Text
RegDeleteKey HKEY_USERS, str
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Form2.Visible = True
End Sub

Private Sub Command7_Click()
Me.chkAst.Value = 0
Me.chkClose.Value = 0
Me.chkDefault.Value = 0
Me.chkExc.Value = 0
Me.chkMax.Value = 0
Me.chkMCmd.Value = 0
Me.chkMin.Value = 0
Me.chkMPop.Value = 0
Me.chkOpen.Value = 0
Me.chkProgError.Value = 0
Me.chkQue.Value = 0
Me.chkResDwn.Value = 0
Me.chkResUp.Value = 0
Me.txtAppname.Text = ""
Me.txtExeFilename.Text = ""
End Sub
