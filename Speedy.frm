VERSION 5.00
Begin VB.Form Speedy 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   570
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3150
   ControlBox      =   0   'False
   Icon            =   "Speedy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1260
      Top             =   90
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   825
      Top             =   60
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "Speedy.frx":030A
      Stretch         =   -1  'True
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   0
      Top             =   180
      Width           =   2535
   End
End
Attribute VB_Name = "Speedy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_Load()
a = 0
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
On Error GoTo 10
If Form1.RecPath.Text = "" Then GoTo 10
If Form1.RecPath.Text = "Not found - Impossible delete files" Then GoTo 10

'Elimina i files lnk
Kill Form1.RecPath + "\*.lnk": GoTo 20
10 'Errore
Me.Visible = True
Timer1.Enabled = True
GoTo 30

20 'Ok
Me.Visible = True
Timer2.Enabled = True


30
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = vbRed
a = a + 1
If a = 1 Then Label1.Caption = ""
If a = 2 Then Label1.Caption = "R"
If a = 3 Then Label1.Caption = "Re"
If a = 4 Then Label1.Caption = "Rec"
If a = 5 Then Label1.Caption = "Rece"
If a = 6 Then Label1.Caption = "Recent"
If a = 7 Then Label1.Caption = "Recent "
If a = 8 Then Label1.Caption = "Recent f"
If a = 9 Then Label1.Caption = "Recent fi"
If a = 10 Then Label1.Caption = "Recent fil"
If a = 11 Then Label1.Caption = "Recent file"
If a = 12 Then Label1.Caption = "Recent files"
If a = 13 Then Label1.Caption = "Recent files "
If a = 14 Then Label1.Caption = "Recent files n"
If a = 15 Then Label1.Caption = "Recent files no"
If a = 16 Then Label1.Caption = "Recent files not"
If a = 17 Then Label1.Caption = "Recent files not "
If a = 18 Then Label1.Caption = "Recent files not f"
If a = 19 Then Label1.Caption = "Recent files not fo"
If a = 20 Then Label1.Caption = "Recent files not fou"
If a = 21 Then Label1.Caption = "Recent files not foun"
If a = 22 Then Label1.Caption = "Recent files not found"
If a = 50 Then Me.Height = Me.Height - 50
If a = 51 Then Me.Height = Me.Height - 50
If a = 52 Then Me.Height = Me.Height - 50
If a = 53 Then Me.Height = Me.Height - 50
If a = 54 Then Me.Height = Me.Height - 50
If a = 55 Then Me.Height = Me.Height - 50
If a = 56 Then Me.Height = Me.Height - 50
If a = 57 Then Me.Height = Me.Height - 50
If a = 58 Then Me.Width = Me.Width - 100
If a = 59 Then Me.Width = Me.Width - 100
If a = 60 Then Me.Width = Me.Width - 100
If a = 61 Then Me.Width = Me.Width - 100
If a = 62 Then Me.Width = Me.Width - 100
If a = 63 Then Me.Width = Me.Width - 100
If a = 64 Then Me.Width = Me.Width - 100
If a = 65 Then Me.Width = Me.Width - 100
If a = 66 Then Me.Width = Me.Width - 100
If a = 67 Then Me.Width = Me.Width - 100
If a = 68 Then Me.Width = Me.Width - 100
If a = 69 Then Me.Width = Me.Width - 100
If a = 58 Then Me.Width = Me.Width - 100
If a = 59 Then Me.Width = Me.Width - 100
If a = 60 Then Me.Width = Me.Width - 100
If a = 61 Then Me.Width = Me.Width - 100
If a = 62 Then Me.Width = Me.Width - 100
If a = 70 Then Me.Width = Me.Width - 100
If a = 71 Then Me.Width = Me.Width - 100
If a = 72 Then Me.Width = Me.Width - 100
If a = 73 Then Me.Width = Me.Width - 100
If a = 74 Then Me.Width = Me.Width - 100
If a = 75 Then Me.Width = Me.Width - 100
If a = 76 Then Me.Width = Me.Width - 100
If a = 77 Then Me.Width = Me.Width - 100
If a = 78 Then Me.Width = Me.Width - 100
If a = 79 Then Me.Width = Me.Width - 100
If a = 80 Then Me.Width = Me.Width - 100
If a = 81 Then Unload Me

End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = vbBlack
a = a + 1
If a = 1 Then Label1.Caption = ""
If a = 2 Then Label1.Caption = "R"
If a = 3 Then Label1.Caption = "Re"
If a = 4 Then Label1.Caption = "Rec"
If a = 5 Then Label1.Caption = "Rece"
If a = 6 Then Label1.Caption = "Recen"
If a = 7 Then Label1.Caption = "Recent"
If a = 8 Then Label1.Caption = "Recent "
If a = 9 Then Label1.Caption = "Recent f"
If a = 10 Then Label1.Caption = "Recent fi"
If a = 11 Then Label1.Caption = "Recent fil"
If a = 12 Then Label1.Caption = "Recent file"
If a = 13 Then Label1.Caption = "Recent files"
If a = 14 Then Label1.Caption = "Recent files "
If a = 15 Then Label1.Caption = "Recent files d"
If a = 16 Then Label1.Caption = "Recent files de"
If a = 17 Then Label1.Caption = "Recent files del"
If a = 18 Then Label1.Caption = "Recent files dele"
If a = 19 Then Label1.Caption = "Recent files delet"
If a = 20 Then Label1.Caption = "Recent files delete"
If a = 21 Then Label1.Caption = "Recent files deleted"
If a = 50 Then Me.Height = Me.Height - 50
If a = 51 Then Me.Height = Me.Height - 50
If a = 52 Then Me.Height = Me.Height - 50
If a = 53 Then Me.Height = Me.Height - 50
If a = 54 Then Me.Height = Me.Height - 50
If a = 55 Then Me.Height = Me.Height - 50
If a = 56 Then Me.Height = Me.Height - 50
If a = 57 Then Me.Height = Me.Height - 50
If a = 58 Then Me.Width = Me.Width - 100
If a = 59 Then Me.Width = Me.Width - 100
If a = 60 Then Me.Width = Me.Width - 100
If a = 61 Then Me.Width = Me.Width - 100
If a = 62 Then Me.Width = Me.Width - 100
If a = 63 Then Me.Width = Me.Width - 100
If a = 64 Then Me.Width = Me.Width - 100
If a = 65 Then Me.Width = Me.Width - 100
If a = 66 Then Me.Width = Me.Width - 100
If a = 67 Then Me.Width = Me.Width - 100
If a = 68 Then Me.Width = Me.Width - 100
If a = 69 Then Me.Width = Me.Width - 100
If a = 58 Then Me.Width = Me.Width - 100
If a = 59 Then Me.Width = Me.Width - 100
If a = 60 Then Me.Width = Me.Width - 100
If a = 61 Then Me.Width = Me.Width - 100
If a = 62 Then Me.Width = Me.Width - 100
If a = 70 Then Me.Width = Me.Width - 100
If a = 71 Then Me.Width = Me.Width - 100
If a = 72 Then Me.Width = Me.Width - 100
If a = 73 Then Me.Width = Me.Width - 100
If a = 74 Then Me.Width = Me.Width - 100
If a = 75 Then Me.Width = Me.Width - 100
If a = 76 Then Me.Width = Me.Width - 100
If a = 77 Then Me.Width = Me.Width - 100
If a = 78 Then Me.Width = Me.Width - 100
If a = 79 Then Me.Width = Me.Width - 100
If a = 80 Then Me.Width = Me.Width - 100
If a = 81 Then Unload Me: Form1.File1.Refresh: Form1.deleteme.Enabled = False: Form1.propertiesme.Enabled = False: Form1.Command5.Enabled = False: Form1.Command7.Enabled = False

End Sub

