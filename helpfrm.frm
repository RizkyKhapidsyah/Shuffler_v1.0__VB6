VERSION 5.00
Begin VB.Form helpfrm 
   BackColor       =   &H0080C0FF&
   Caption         =   "Help"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   7800
      Picture         =   "helpfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label helptext 
      BackColor       =   &H0080C0FF&
      Caption         =   $"helpfrm.frx":018A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   5880
      Width           =   7935
   End
   Begin VB.Label helptext 
      BackColor       =   &H0080C0FF&
      Caption         =   $"helpfrm.frx":0372
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   7935
   End
   Begin VB.Label helptext 
      BackColor       =   &H0080C0FF&
      Caption         =   $"helpfrm.frx":050B
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   7935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      X1              =   120
      X2              =   8040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label helptext 
      BackColor       =   &H0080C0FF&
      Caption         =   "This game is simple, you shouldnt need help but it looks like you do so lets get on with it!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "helpfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
helpfrm.Visible = False

End Sub

