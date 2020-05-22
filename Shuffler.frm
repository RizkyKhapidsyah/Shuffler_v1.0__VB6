VERSION 5.00
Begin VB.Form Shuffler 
   BackColor       =   &H0080FF80&
   Caption         =   "Shuffler - By Rizky Khapidsyah"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   Icon            =   "Shuffler.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "&Help"
      Height          =   495
      Left            =   10800
      Picture         =   "Shuffler.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton stopbtn 
      Caption         =   "Stop"
      Height          =   615
      Left            =   960
      Picture         =   "Shuffler.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton playbtn 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "Shuffler.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Play the music"
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer musictime 
      Interval        =   29000
      Left            =   120
      Top             =   3480
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   18
      Top             =   5640
      Width           =   975
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   17
      Top             =   6240
      Width           =   975
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   16
      Top             =   5640
      Width           =   975
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox holder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton shuffbtn 
      Caption         =   "&SHUFFLE!"
      Height          =   615
      Left            =   720
      TabIndex        =   10
      ToolTipText     =   "Shuffle the numbers around"
      Top             =   6840
      Width           =   6255
   End
   Begin VB.CommandButton quitbtn 
      BackColor       =   &H0080FF80&
      Caption         =   "&Quit"
      Height          =   495
      Left            =   10680
      TabIndex        =   0
      ToolTipText     =   "Quit Shuffler"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Music Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label hiscr 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   6
      Left            =   10800
      TabIndex        =   37
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label hiscr 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   5
      Left            =   10800
      TabIndex        =   36
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label hiscr 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   4
      Left            =   10800
      TabIndex        =   35
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label hiscr 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   10800
      TabIndex        =   34
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label hiscr 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   10800
      TabIndex        =   33
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label hiscr 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   32
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label hiscr 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   31
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label highscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Richard Stoddart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   6
      Left            =   7800
      TabIndex        =   30
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label highscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Richard Stoddart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   29
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label highscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Richard Stoddart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   28
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label highscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Richard Stoddart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   27
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label highscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Richard Stoddart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   26
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label highscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Richard Stoddart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   25
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label highscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Richard Stoddart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   24
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "High - Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   7680
      TabIndex        =   23
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Shuffles left:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label shuffled 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   6360
      TabIndex        =   21
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Shuffles:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      TabIndex        =   20
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label shuffles 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   6480
      TabIndex        =   19
      ToolTipText     =   "The number of shuffles you have made"
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":0720
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   7
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Drag this to the main stack"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":0B62
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   6
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Drag this to the main stack"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":0FA4
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   5
      Left            =   4800
      TabIndex        =   7
      ToolTipText     =   "Drag this to the main stack"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":13E6
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   4
      Left            =   5520
      TabIndex        =   6
      ToolTipText     =   "Drag this to the main stack"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":1828
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   3
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Drag this to the main stack"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":1C6A
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Drag this to the main stack"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":20AC
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Drag this to the main stack"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "00"
      DragIcon        =   "Shuffler.frx":24EE
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   0
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "Drag this to the main stack"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label mainnum 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Main Number Stack"
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "Shuffler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Dim hin(0 To 6) As String
Dim hisc(0 To 6) As String
Dim SoundFile As String
Private Declare Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long

Private Sub buffy_Click()
SoundFile = "oooh.wav"
Result = sndPlaySound(SoundFile, 0)
If playbtn.Enabled = False Then
   SoundFile = "music.wav"
   Result = sndPlaySound(SoundFile, 1)
End If
End Sub

Private Sub Command1_Click()
helpfrm.Visible = True

End Sub

Private Sub Form_Click()
If playbtn.Enabled = True Then
   SoundFile = "fart.wav"
   Result = sndPlaySound(SoundFile, 1)
End If
End Sub

Private Sub Form_Load()
SoundFile = "music.wav"
Result = sndPlaySound(SoundFile, 1)

On Error GoTo errorhandler
Randomize
Open "c:\windows\system\shcon.cht" For Input As #1
For c = 0 To 6
    Input #1, hin(c), hisc(c)
Next c
For t = 0 To 6
    highscore(t).Caption = hin(t)
    hiscr(t) = Val(hisc(t))
Next t
Close #1
mainnum.Caption = Int(Rnd * 50) + 1
For X = 0 To 7
    num(X).Caption = Int(Rnd * 50) + 1
Next X

errorhandler:
If Err = 53 Then
   MsgBox ("The Hi-scores file is missing, attempting to repair...")
   u = 8
   For d = 0 To 6
       u = u - 1
       highscore(d) = "The Shuffler"
       hiscr(d) = u
   Next d
      Open "c:\windows\system\shcon.cht" For Output As #1
       For c = 0 To 6
           Write #1, highscore(c), hiscr(c)
       Next c
   Close #1
   MsgBox ("Please restart shuffler, the error has been fixed")
   End
End If
End Sub

Private Sub mainnum_DragDrop(Source As Control, X As Single, Y As Single)
doit = 0
If TypeOf Source Is Label Then
   pilenum = Val(mainnum.Caption)
   dragnum = Val(Source.Caption)
   If dragnum = pilenum Then doit = 1
   If dragnum - 1 = pilenum Then doit = 1
   If dragnum + 1 = pilenum Then doit = 1
   If doit = 1 Then
       mainnum.Caption = Source.Caption
       Source.Caption = Int(Rnd * 50) + 1
       shuffles.Caption = Val(shuffles.Caption) + 1
    End If
End If
End Sub

Private Sub musictime_Timer()
SoundFile = "music.wav"
Result = sndPlaySound(SoundFile, 1)
End Sub

Private Sub playbtn_Click()
SoundFile = "music.wav"
Result = sndPlaySound(SoundFile, 1)
playbtn.Enabled = False
stopbtn.Enabled = True
musictime.Enabled = True
End Sub

Private Sub quitbtn_Click()
SoundFile = "bye.wav"
Result = sndPlaySound(SoundFile, 0)
End
End Sub

Private Sub shuffbtn_Click()
If shuffled > 0 Then
   If playbtn.Enabled = True Then
      SoundFile = "shuf.wav"
      Result = sndPlaySound(SoundFile, 1)
   End If
   shuffled = shuffled - 1
   For X = 0 To 7
       If holder(X).Value = 0 Then
          num(X).Caption = Int(Rnd * 50) + 1
       End If
   Next X
Else
   MsgBox ("You are out of shuffles, but you managed to score " & shuffles.Caption & " Shuffles!!")
   If Val(shuffles.Caption) > Val(hiscr(6).Caption) Then
      dohighscores
   Else
      SoundFile = "loser.wav"
      Result = sndPlaySound(SoundFile, 1)
      Dim Msg$, Style, Tit$, Help, Ctxt, Response, MyString
      Msg$ = "Do you want to play again ?"
      Style = vbYesNo + vbQuestion + vbDefaultButton1
      Tit$ = "Leaving?"
      Response = MsgBox(Msg$, Style, Tit$)
      If Response = vbYes Then
         mainnum.Caption = Int(Rnd * 50) + 1
         For b = 0 To 7
             num(b).Caption = Int(Rnd * 50) + 1
         Next b
         shuffles.Caption = 0
         shuffled.Caption = 10
         For X = 0 To 7
             holder(X).Value = 0
         Next X
      Else
         SoundFile = "bye.wav"
         Result = sndPlaySound(SoundFile, 0)
         End
      End If
   End If
End If
End Sub

Public Sub dohighscores()
SoundFile = "crowd.wav"
Result = sndPlaySound(SoundFile, 1)
MsgBox ("Well done, you made it on to the High-Score table!!")
r = 7
For X = 0 To 6
    If Val(shuffles.Caption) >= Val(hiscr(X).Caption) Then
       Do
         r = r - 1
         If r <= X Then Exit Do
         highscore(r).Caption = highscore(r - 1).Caption
         highscore(r - 1).Caption = ""
         hiscr(r).Caption = hiscr(r - 1).Caption
         hiscr(r - 1).Caption = ""
       Loop
       Message$ = "Please Give me your name (max 20 chrs)"
       Title$ = "High-Score Name Entry"
       Defaul$ = "The Shuffler"
       Do
         p1name$ = InputBox(Message$, Title$, Defaul$)
         If Len(p1name$) <= 20 Then
            Exit Do
        End If
       Loop
       highscore(X).Caption = p1name$
       hiscr(X).Caption = shuffles.Caption
       Open "c:\windows\system\shcon.cht" For Output As #1
       For c = 0 To 6
           Write #1, highscore(c), hiscr(c)
       Next c
       Close #1
       Dim Msg$, Style, Tit$, Help, Ctxt, Response, MyString
       Msg$ = "Do you want to play again ?"
       Style = vbYesNo + vbQuestion + vbDefaultButton1
       Tit$ = "Leaving?"
       Response = MsgBox(Msg$, Style, Tit$)
       If Response = vbYes Then
          mainnum.Caption = Int(Rnd * 50) + 1
          For b = 0 To 7
              num(b).Caption = Int(Rnd * 50) + 1
          Next b
          shuffles.Caption = 0
          shuffled.Caption = 10
          Exit For
       Else
          End
       End If
    End If
Next X
End Sub

Private Sub stopbtn_Click()
SoundFile = "arrow1.wav"
Result = sndPlaySound(SoundFile, 1)
playbtn.Enabled = True
stopbtn.Enabled = False
musictime.Enabled = False
End Sub
