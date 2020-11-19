VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   885
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5175
      Top             =   270
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 4.3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1365
      TabIndex        =   0
      Top             =   645
      Width           =   2880
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   -15
      Picture         =   "frmSplash.frx":0CCA
      Top             =   -15
      Width           =   6000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo 10
Versione = "4.3"
Label2.Caption = "Version " + Versione


Dim splashScr
splashScr = GetSetting("Deleter", "Prev", "Check2")
If splashScr = 1 Then Timer1.Enabled = True: Me.Visible = True: GoTo 15
If splashScr = 0 Then Unload Me: Load Form1:  GoTo 15

10
'Eseguito se non trova i valori nel registro (1° avvio)
SaveSetting "Deleter", "Prev", "Check1", "0"
SaveSetting "Deleter", "Prev", "Check2", "1"
SaveSetting "Deleter", "Prev", "Check3", "0"








Timer1.Enabled = True
Me.Visible = True
15
End Sub


Private Sub Timer1_Timer()
Unload Me: Load Form1
End Sub

