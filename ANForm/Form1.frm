VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Form"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6315
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00666666&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6315
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2370
      TabIndex        =   1
      Top             =   900
      Width           =   825
   End
   Begin ANFormX.ANForm ANForm1 
      Height          =   345
      Left            =   2250
      TabIndex        =   0
      Top             =   1800
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      StyleColor      =   5
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub
