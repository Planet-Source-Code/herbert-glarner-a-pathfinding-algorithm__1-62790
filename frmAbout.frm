VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "About"
   ClientHeight    =   3135
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2163.833
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAuthor 
      Appearance      =   0  '2D
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'Kein
      Height          =   675
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmAbout.frx":0442
      Top             =   1710
      Width           =   3915
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":049B
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Innen ausgefüllt
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "You are free to use this code for whatever purpose, as long as the following credits are included."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   2
      Top             =   1170
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "A+ Pathfinding Algorithm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version: Non-Constrained (2005)"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Use at own risk."
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   255
      TabIndex        =   3
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

