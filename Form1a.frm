VERSION 5.00
Object = "{1F787982-E6F9-11D2-9945-0040056CD8C0}#1.0#0"; "QUICKREG.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "\\"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "Form1a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Use default directory or custom directory"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox RecPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Not found - Impossible delete files"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About..."
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      ToolTipText     =   "About Deleter"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Recent files :"
      Height          =   4935
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   2655
      Begin VB.CommandButton Command10 
         Caption         =   "Open Folder"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Open Recent files folder"
         Top             =   4440
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Refresh List"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         ToolTipText     =   "Refresh Recent files list"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Properties"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Show Recent file properties"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Execute"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "Execute selected Recent file"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Delete selected Recent file"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete All Recent files"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Delete All Recent files"
         Top             =   3360
         Width           =   2415
      End
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   120
         Pattern         =   "*.lnk"
         System          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      ToolTipText     =   "Minimize at System Tray icon"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      ToolTipText     =   "Remove System Tray icon and exit"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options :"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      Begin VB.CheckBox Check2 
         Caption         =   "Show SplashScreen at program startup"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Enable or Disable SplashScreen at program startup"
         Top             =   480
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Delete Recent files at program startup"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Enable or Disable delete Recent files at program startup"
         Top             =   240
         Width           =   3735
      End
   End
   Begin QUICKREGLib.QuickReg QuickReg1 
      Left            =   3120
      Top             =   4320
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   0
      SubKeyPath      =   "SoftwareMicrosoft\Windows\CurrentVersion"
   End
   Begin VB.Label Label1 
      Caption         =   "Windows Recent files path : "
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu fileme 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu executeme 
         Caption         =   "Execute"
      End
      Begin VB.Menu deleteme 
         Caption         =   "Delete selected"
         Enabled         =   0   'False
      End
      Begin VB.Menu deletemeall 
         Caption         =   "Delete All"
      End
      Begin VB.Menu refreshme 
         Caption         =   "Refresh"
      End
      Begin VB.Menu vuoto1 
         Caption         =   "-"
      End
      Begin VB.Menu propertiesme 
         Caption         =   "Properties"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim b As NOTIFYICONDATA
Dim Sel As String


Private Sub Check3_Click()
If Check3.Value = 0 Then
    RecPath.Enabled = False
    'Preleva dal registro la directory corrente dei file recenti
    On Error GoTo 10
    Dim Reg1 As String
    Dim Reg2 As Long
    Dim Reg3 As Long
    Dim Reg4 As String
    Reg1 = "Recent"
    Reg2 = 1
    Reg3 = 0
    Reg4 = "Not found - Impossible delete files"
    QuickReg1.RootKey = "HKEY_CURRENT_USER"
    QuickReg1.SubKeyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    QuickReg1.GetValueData Reg1, Reg2, Reg3, Reg4
    File1.Path = Reg4
    Form1.RecPath = Reg4
10
Else
    RecPath.Enabled = True
End If

End Sub

Private Sub Command1_Click()
Shell_NotifyIcon NIM_DELETE, b
If Check1.Value = 1 Then SaveSetting "Deleter", "Prev", "Check1", "1"
If Check1.Value = 0 Then SaveSetting "Deleter", "Prev", "Check1", "0"
If Check2.Value = 1 Then SaveSetting "Deleter", "Prev", "Check2", "1"
If Check2.Value = 0 Then SaveSetting "Deleter", "Prev", "Check2", "0"
If Check3.Value = 1 Then SaveSetting "Deleter", "Prev", "Check3", "1"
If Check3.Value = 0 Then SaveSetting "Deleter", "Prev", "Check3", "0"
SaveSetting "Deleter", "Prev", "Dir", Form1.RecPath.Text

End
End Sub

Private Sub Command10_Click()
Dim res&
res& = ShellExecute(hwnd, "open", vbNullString, vbNullString, Form1.File1.Path, SW_SHOW)
'Me.Caption = res
If res < 32 Then
    MsgBox "Unable to open folder"
End If

End Sub

Private Sub Command2_Click()
Me.Hide: Unload frmView: Unload About
End Sub

Private Sub Command3_Click()
Speedy.Show
executeme.Enabled = False
propertiesme.Enabled = False
propertiesme.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End Sub


Private Sub Command4_Click()
About.Show 1
End Sub

Private Sub Command5_Click()
On Error GoTo 10
Kill Sel
10
File1.Refresh: executeme.Enabled = False: deleteme.Enabled = False: propertiesme.Enabled = False: Command5.Enabled = False: Command7.Enabled = False: Command8.Enabled = False
End Sub

Private Sub Command6_Click()
File1.Refresh
deleteme.Enabled = False
propertiesme.Enabled = False
executeme.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End Sub

Private Sub Command7_Click()
On Error GoTo 10
s$ = Form1.File1.Path + "\" + Form1.File1.fileName
NomeSelezionato = s
frmView.Initialize s$
frmView.Show 1
10

End Sub

Private Sub Command8_Click()
Dim aaaa
aaaa = File1.Path + "\" + File1.fileName
Dim res&
res& = ShellExecute(hwnd, "open", aaaa, vbNullString, vbNullString, SW_SHOW)
If res < 32 Then
    MsgBox "Unable to open selected image"
End If
End Sub


Private Sub executeme_Click()
Command8_Click
End Sub

Private Sub deleteme_Click()
Command5_Click
End Sub

Private Sub deletemeall_Click()
Command3_Click
End Sub




Private Sub File1_Click()
Sel = File1.Path + "\" + File1.fileName
deleteme.Enabled = True
propertiesme.Enabled = True
executeme.Enabled = True
Command5.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
End Sub

Private Sub File1_DblClick()
On Error GoTo 10
Sel = File1.Path + "\" + File1.fileName
Kill Sel
10
File1.Refresh: executeme.Enabled = False: deleteme.Enabled = False: propertiesme.Enabled = False: Command5.Enabled = False: Command7.Enabled = False: Command8.Enabled = False
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Sel = File1.Path + "\" + File1.fileName
Me.PopupMenu fileme
End If
End Sub

Private Sub Form_Initialize()
b.cbSize = Len(b)
b.hIcon = Me.Icon.Handle
b.hwnd = Me.hwnd
b.szTip = "Recent Files Deleter " + Versione & vbNullChar
b.uCallbackMessage = WM_MOUSEMOVE
b.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
b.uID = vbNull
Shell_NotifyIcon NIM_ADD, b
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim c1
Dim c2
Dim c3
Dim c4
Me.Hide

Me.Caption = "Deleter " + Versione

c1 = GetSetting("Deleter", "Prev", "Check1")
c2 = GetSetting("Deleter", "Prev", "Check2")
c3 = GetSetting("Deleter", "Prev", "Check3")
c4 = GetSetting("Deleter", "Prev", "Dir")

If c1 = 1 Then Check1.Value = 1
If c1 = 0 Then Check1.Value = 0
If c2 = 1 Then Check2.Value = 1
If c2 = 0 Then Check2.Value = 0
If c3 = 1 Then Check3.Value = 1
If c3 = 0 Then Check3.Value = 0
RecPath.Text = c4
File1.Path = c4
Check3_Click

Me.Top = -Screen.Height - 200
Me.Left = -Screen.Width - 200


10
'Cancella i file all'avvio
If c1 = 1 Then Speedy.Show
Me.Hide

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
msg = X \ Screen.TwipsPerPixelX
Select Case msg
Case WM_RBUTTONDOWN
     Me.File1.Refresh: deleteme.Enabled = False: propertiesme.Enabled = False: Command5.Enabled = False: Command7.Enabled = False
     Me.Show
Case WM_LBUTTONDOWN
     Speedy.Show
End Select
10
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
If UnloadMode = vbFormControlMenu Then Me.Hide: Unload frmView: Unload About
If UnloadMode = vbAppTaskManager Or UnloadMode = vbAppWindows Then
    Shell_NotifyIcon NIM_DELETE, b
    If Check1.Value = 1 Then SaveSetting "Deleter", "Prev", "Check1", "1"
    If Check1.Value = 0 Then SaveSetting "Deleter", "Prev", "Check1", "0"
    If Check2.Value = 1 Then SaveSetting "Deleter", "Prev", "Check2", "1"
    If Check2.Value = 0 Then SaveSetting "Deleter", "Prev", "Check2", "0"
    If Check3.Value = 1 Then SaveSetting "Deleter", "Prev", "Check3", "1"
    If Check3.Value = 0 Then SaveSetting "Deleter", "Prev", "Check3", "0"
    SaveSetting "Deleter", "Prev", "Dir", Form1.RecPath.Text
    Cancel = False
    End
End If
End Sub


Private Sub propertiesme_Click()
Command7_Click
End Sub

Private Sub refreshme_Click()
Command6_Click
End Sub
