VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Data datFlightDesc 
      Caption         =   "Flight Descriptor Table"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame fraMessage 
      Caption         =   "Message Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5760
      TabIndex        =   47
      Top             =   10200
      Width           =   4815
      Begin VB.TextBox txtMessage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame fraInstructions 
      Caption         =   "Instruction Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10680
      TabIndex        =   43
      Top             =   10200
      Width           =   4215
      Begin VB.CommandButton cmdDown 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   60
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   59
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox cboNewHeading 
         Height          =   315
         ItemData        =   "frmBomberScaled.frx":0000
         Left            =   1680
         List            =   "frmBomberScaled.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cboNewAltitude 
         Height          =   315
         ItemData        =   "frmBomberScaled.frx":0050
         Left            =   1680
         List            =   "frmBomberScaled.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmdInstruction 
         Caption         =   "Issue Instruction"
         Height          =   645
         Left            =   2760
         TabIndex        =   52
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Change Heading to:"
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Change Altitude to:"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.Frame fraControls 
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   10200
      Width           =   5535
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
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
         Left            =   1920
         TabIndex        =   54
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtExit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   480
         Width           =   855
      End
      Begin VB.Timer tmrMovement 
         Interval        =   200
         Left            =   5160
         Top             =   0
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtAltitude 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtHeading 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
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
         Left            =   3720
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
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
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtCallSign 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Exit Route"
         Height          =   195
         Left            =   4440
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Entry Route"
         Height          =   195
         Left            =   3480
         TabIndex        =   40
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Altitude"
         Height          =   195
         Left            =   2760
         TabIndex        =   38
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Heading"
         Height          =   195
         Left            =   1800
         TabIndex        =   36
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Call Sign:"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "A7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   14520
      TabIndex        =   61
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lblAlert 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Caption         =   "Press ""Resume"" Button to Resume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5243
      TabIndex        =   58
      Top             =   9360
      Visible         =   0   'False
      Width           =   4875
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MURTALA MOHAMMED INTERNATIONAL AIRPORT'S AIRSPACE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   1950
      TabIndex        =   57
      Top             =   0
      Width           =   11640
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Caption         =   "Press ""Resume"" Button to Resume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   5123
      TabIndex        =   56
      Top             =   5940
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.Label lblPause 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Caption         =   "DISPLAY PAUSED"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Index           =   0
      Left            =   5123
      TabIndex        =   55
      Top             =   5220
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   19
      Left            =   1440
      Picture         =   "frmBomberScaled.frx":00B8
      Tag             =   "Double click to issue a command"
      Top             =   960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   18
      Left            =   13080
      Picture         =   "frmBomberScaled.frx":03D6
      Stretch         =   -1  'True
      Tag             =   "Double click to issue a command"
      Top             =   6000
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   17
      Left            =   13080
      Picture         =   "frmBomberScaled.frx":06F4
      Tag             =   "Double click to issue a command"
      Top             =   5640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   16
      Left            =   13080
      Picture         =   "frmBomberScaled.frx":0A23
      Tag             =   "Double click to issue a command"
      Top             =   5400
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   15
      Left            =   13080
      Picture         =   "frmBomberScaled.frx":0D52
      Tag             =   "Double click to issue a command"
      Top             =   5160
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   14
      Left            =   13080
      Picture         =   "frmBomberScaled.frx":1081
      Tag             =   "Double click to issue a command"
      Top             =   4920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   13
      Left            =   13080
      Picture         =   "frmBomberScaled.frx":13B0
      Tag             =   "Double click to issue a command"
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   12
      Left            =   13080
      Picture         =   "frmBomberScaled.frx":16CE
      Tag             =   "Double click to issue a command"
      Top             =   4320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   11
      Left            =   12480
      Picture         =   "frmBomberScaled.frx":19FD
      Tag             =   "Double click to issue a command"
      Top             =   5640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   10
      Left            =   12480
      Picture         =   "frmBomberScaled.frx":1D2C
      Tag             =   "Double click to issue a command"
      Top             =   5400
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   9
      Left            =   12480
      Picture         =   "frmBomberScaled.frx":205B
      Tag             =   "Double click to issue a command"
      Top             =   5160
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   8
      Left            =   12480
      Picture         =   "frmBomberScaled.frx":238A
      Tag             =   "Double click to issue a command"
      Top             =   4920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   7
      Left            =   12480
      Picture         =   "frmBomberScaled.frx":26B9
      Tag             =   "Double click to issue a command"
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   6
      Left            =   12480
      Picture         =   "frmBomberScaled.frx":29D7
      Tag             =   "Double click to issue a command"
      Top             =   4320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape5 
      FillStyle       =   4  'Upward Diagonal
      Height          =   1935
      Left            =   3960
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   4
      Left            =   360
      Picture         =   "frmBomberScaled.frx":2D06
      Tag             =   "Double click to issue a command"
      Top             =   4080
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   5
      Left            =   240
      Picture         =   "frmBomberScaled.frx":3024
      Tag             =   "Double click to issue a command"
      Top             =   3840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   3
      Left            =   8280
      Picture         =   "frmBomberScaled.frx":3342
      Tag             =   "Double click to issue a command"
      Top             =   3480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   2
      Left            =   840
      Picture         =   "frmBomberScaled.frx":3660
      Tag             =   "Double click to issue a command"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   1
      Left            =   240
      Picture         =   "frmBomberScaled.frx":397E
      Tag             =   "Double click to issue a command"
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgPlanePic 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   0
      Left            =   6720
      Picture         =   "frmBomberScaled.frx":3C9C
      Tag             =   "Double click to issue a command"
      Top             =   8520
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Caption         =   "to NW Africa Airspace (NWA)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   46
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lblGhana 
      Alignment       =   2  'Center
      Caption         =   "to GHANAIAN Airspace (GHA)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   45
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Label lblSA 
      Alignment       =   2  'Center
      Caption         =   "to Southern Africa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2520
      TabIndex        =   44
      Top             =   9480
      Width           =   855
   End
   Begin VB.Label lblLOSTakeOff 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4080
      TabIndex        =   34
      Top             =   6960
      Width           =   390
   End
   Begin VB.Label lblBEN 
      Caption         =   "to BEN ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14280
      TabIndex        =   33
      Top             =   5880
      Width           =   840
   End
   Begin VB.Label lblLOSLanding 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7200
      TabIndex        =   32
      Top             =   3960
      Width           =   345
   End
   Begin VB.Label lblPHC 
      AutoSize        =   -1  'True
      Caption         =   "to PHC ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14280
      TabIndex        =   31
      Top             =   7920
      Width           =   840
   End
   Begin VB.Label lblABJ 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "to ABJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11640
      TabIndex        =   30
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "LOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   29
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "B7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13920
      TabIndex        =   28
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "C7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13920
      TabIndex        =   27
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "D7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13920
      TabIndex        =   26
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "E7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13920
      TabIndex        =   25
      Top             =   7920
      Width           =   240
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   24
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   23
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   22
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   21
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9840
      TabIndex        =   20
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11880
      TabIndex        =   19
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13920
      TabIndex        =   18
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "A7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13920
      TabIndex        =   17
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "A6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11880
      TabIndex        =   16
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "A5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9840
      TabIndex        =   15
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "A4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   14
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "A3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   13
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "A2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   12
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "A1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "F0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   9840
      Width           =   225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "E0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   7920
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "D0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "C0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "B0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "A0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   240
   End
   Begin VB.Line Line12 
      X1              =   0
      X2              =   15000
      Y1              =   10000
      Y2              =   10000
   End
   Begin VB.Line Line11 
      X1              =   0
      X2              =   15000
      Y1              =   8000
      Y2              =   8000
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   15000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   15120
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   15000
      Y1              =   2000
      Y2              =   2000
   End
   Begin VB.Line Line7 
      X1              =   14000
      X2              =   14000
      Y1              =   0
      Y2              =   10000
   End
   Begin VB.Line Line6 
      X1              =   12000
      X2              =   12000
      Y1              =   0
      Y2              =   10000
   End
   Begin VB.Line Line5 
      X1              =   10000
      X2              =   10000
      Y1              =   0
      Y2              =   10000
   End
   Begin VB.Line Line4 
      X1              =   8000
      X2              =   8000
      Y1              =   0
      Y2              =   10000
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   6000
      Y1              =   0
      Y2              =   10000
   End
   Begin VB.Line Line2 
      X1              =   4000
      X2              =   4000
      Y1              =   0
      Y2              =   10000
   End
   Begin VB.Line Line1 
      X1              =   2000
      X2              =   2000
      Y1              =   0
      Y2              =   10000
   End
   Begin VB.Shape Shape2 
      FillStyle       =   2  'Horizontal Line
      Height          =   4215
      Left            =   8040
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      FillStyle       =   2  'Horizontal Line
      Height          =   2175
      Left            =   6000
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      FillStyle       =   4  'Upward Diagonal
      Height          =   3975
      Left            =   2040
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   -120
      Top             =   -120
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: CATCS (Computerized Air Traffic Control System) for
'Murtala Mohammed International Airport, Ikeja.
'Authored by: OGUNBANJO David Oluseyi as a Final Year (BSc.) Project at
'IGBINEDION UNIVERSITY, OKADA, EDO STATE, NIGERIA.
'Project Start Date: Wed. 12th July, 2006.
'Project Completion Date: Tues. 19th July, 2006.

Private aircraft(19) As New CPlane  'Variable array for aircraft object
                                    'The flight capacity of the airspace is 20 aircraft
Dim currentSelection As Integer     'holds index of the currently selected aircraft
Option Explicit
Dim xyz1 As Image
'Form Code

Private Sub cmdAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub cmdDown_Click()
 If simSpeed > 0 Then simSpeed = simSpeed - 1
End Sub

Private Sub cmdExit_Click()
 If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Quit?") = vbYes Then
    Unload Me
    End
 End If
End Sub

Private Sub cmdInstruction_Click()
  
  
 If txtCallSign = "" Then
    MsgBox "Please select an aircraft to instruct", vbInformation, "Information"
    Exit Sub
 End If
 
 'Ensure that blank parameters are not taken
 If cboNewAltitude.Text = "" Then cboNewAltitude.Text = "NIL"
 If cboNewHeading.Text = "" Then cboNewHeading.Text = "NIL"
 
 'If parameters are blank, display a message
 If (Val(cboNewAltitude.Text) = 0) And (cboNewHeading.Text = "NIL") Then
    MsgBox "You must specify a new altitude and/or heading", vbInformation, "Input Error"
 End If
 
 'If parameters are not blank, give instruction
 If Val(cboNewAltitude.Text) <> 0 Then
    aircraft(currentSelection).Altitude = cboNewAltitude.Text
    txtAltitude = aircraft(currentSelection).Altitude   'Update display text box
 End If
    
 If cboNewHeading.Text <> "NIL" Then
    aircraft(currentSelection).Heading = Val(cboNewHeading.Text)
    txtHeading = aircraft(currentSelection).Heading     'Update display text box
 End If

 'Inform the Controller that changes have been made thru the "message panel"
 
 
 'Clear Instruction panel's Combo boxes in readiness for next input
 cboNewAltitude.Text = "NIL"
 cboNewHeading.Text = "NIL"

End Sub

Private Sub cmdPause_Click()
 If cmdPause.Caption = "Pause" Then     'Display is currently running; pause it
    lblPause(0).Visible = True
    lblPause(1).Visible = True
    tmrMovement.Enabled = False
    cmdPause.Caption = "Unpause"
    fraInstructions.Enabled = False
 Else                                   'Display is currently paused; unpause it
    lblPause(0).Visible = False
    lblPause(1).Visible = False
    tmrMovement.Enabled = True
    cmdPause.Caption = "Pause"
    fraInstructions.Enabled = True
 End If
End Sub

Private Sub cmdUp_Click()
 If simSpeed < 6 Then simSpeed = simSpeed + 1
End Sub

Private Sub Form_Load()
 Me.Scale (19050, 0)-(0, 12000)
' Me.Line (2000, 2000)-(1000, 1000), vbGreen
 simSpeed = 0
 'Task 1: Define the aircraft objects
 Dim count As Integer       'Loop variable
 Dim xCoord As Integer, yCoord As Integer   'X and Y coordinates of aircraft on screen
 For count = 0 To 19
    aircraft(count).imagePlane = imgPlanePic(count)
 Next count
 
 'Task 2:   Initialize the 6 aircraft objects
 'This is divided into two sub-tasks as follows
 'Task 2.1: Connect to database.
 datFlightDesc.DatabaseName = App.Path & "\dbfiles\FlightInfo97.mdb"
 datFlightDesc.RecordSource = "AreaTC"
 datFlightDesc.Refresh
 
 'Task 2.2: Read info from DB.
 With datFlightDesc.Recordset
    .MoveFirst
    count = -1
'Uncomment the next line to use this form without the splash screen
'    If maxAircraft = 0 Then maxAircraft = 9
    Do While Not .EOF And count < maxAircraft 'UBound(aircraft)
        count = count + 1
        'Read aircraft data for next aircraft in db
        aircraft(count).CallSign = .Fields("Call_Sign")
        aircraft(count).Heading = .Fields("Heading")
        aircraft(count).Altitude = .Fields("Altitude")
        aircraft(count).EntryRoute = .Fields("Entry_Route")
        aircraft(count).ExitRoute = .Fields("Exit_Route")
        aircraft(count).FlightState = .Fields("FlightState")
        
        If aircraft(count).FlightState = 1 Then
            Call GetEntryLocation(aircraft(count).EntryRoute, xCoord, yCoord)
            imgPlanePic(count).Left = xCoord
            imgPlanePic(count).Top = yCoord
            imgPlanePic(count).Visible = True
        End If
        .MoveNext
    Loop
'    MsgBox Str(count)
 End With
 
 'Task 3: Open the Log-file for logging flight details.
 Open App.Path + "\CtrlReports\ATC_Log.txt" For Append As #1
 Write #1, "Report Type:  A.C.C. OPERATIONS"
 Write #1, "User name:    " + currentUser
 Write #1, "Date:         " + Str(Date)
 Write #1, "Time:         " + Str(Time)
 Write #1,                'Blank Line
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim X As Integer
 For X = 1 To 5
    Write #1,
 Next X
 Write #1, "-------End of Report----------"
 Close          'All open files
End Sub

Private Sub imgPlanePic_Click(Index As Integer)
 
 currentSelection = Index
 'Remove hightlight from all aircraft
 Dim i As Integer
 For i = 0 To 19
    If (aircraft(i).FlightState = 1) Then
        imgPlanePic(i).BorderStyle = 0     'NONE
    End If
 Next
 'Highlight the selected airplane object
 imgPlanePic(currentSelection).BorderStyle = 1     'FIXED SINGLE

 If lblPause(0).Visible = True Then
    'Display is currently paused
    Exit Sub
 End If
 
 'Display flight parameters for this aircraft
 With aircraft(Index)
    txtCallSign = .CallSign
    txtHeading = .Heading
    txtAltitude = .Altitude
    txtEntry = .EntryRoute
    txtExit = .ExitRoute
    '.imagePlane = App.Path & "\planeRED.jpg"
    'imgPlanePic(Index).Picture = App.Path & "\planeRED.jpg"
 End With
 
 'Clear Instruction panel's Combo boxes in readiness for next input
 cboNewAltitude.Text = "NIL"
 cboNewHeading.Text = "NIL"
 
 If aircraft(Index).FlightState = 1 Then
    fraInstructions.Enabled = True
 Else
    fraInstructions.Enabled = False
 End If

End Sub

Private Sub tmrMovement_Timer()
 lblSpeed.Caption = "X " + Str(2 ^ simSpeed)
 Dim countVar As Integer            'Loop variable
 Dim i As Integer, j As Integer     'Loop variables
 For countVar = 0 To 19             'Sector capacity is 20; range is 0 - 19
    'move all aircraft across the screen
                            
    If aircraft(countVar).FlightState = 1 Then
        'simulate movement for this aircraft w.r.t its heading
        Call aircraft(countVar).fly(aircraft(countVar).Heading)
        
        'Check if this aircraft is flying out of the controlled airspace.
        'If so, release it to the adjacent airspace
        Call ReleaseOutgoingAircraft(aircraft(countVar), countVar)
        
        'Check if this aircraft is flying within landing altitude over the landing beacon.
        'If so, release it to the approach control
        Call ReleaseLandingAircraft(aircraft(countVar), countVar)
        
    
        For i = 0 To 19
            For j = 0 To 19
                If i <> j And (aircraft(i).FlightState = 1) And (aircraft(j).FlightState = 1) And (aircraft(i).Altitude = aircraft(j).Altitude) Then
                    'Check if this aircraft has collided with any other.
                    If CheckForCollision(aircraft(i), aircraft(j)) = True Then
                        'Notify AreaTC of the collision, if such has occured
                        txtMessage.Text = aircraft(i).CallSign + " has been involved in a collision with " + aircraft(j).CallSign
                        'Record the Collision in the log file
                        Write #1, Str(Time), aircraft(i).CallSign + " was involved in a collision with " + aircraft(j).CallSign
                        'Update flight status information
                        aircraft(i).FlightState = -1
                        aircraft(j).FlightState = -1
                        
                        'Remove Highlight from aircraft descriptors
                        imgPlanePic(i).BorderStyle = 0     'NONE
                        imgPlanePic(j).BorderStyle = 0     'NONE
                        'Disable aircraft descriptors
                        imgPlanePic(i).Enabled = False
                        imgPlanePic(j).Enabled = False
                        'Draw a cirle around the collided aircraft descriptors
                        With imgPlanePic(i)
                            Me.PSet (.Left + 1, .Top - 1)
                            Me.Circle (.Left + 1, .Top - 1), 500, vbRed
                        End With
                        
                        With imgPlanePic(j)
                            Me.PSet (.Left + 1, .Top - 1)
                            Me.Circle (.Left + 1, .Top - 1), 500, vbBlue
                        End With
                        
                        lblAlert.Visible = False
                        Beep
                    End If
                    'Check if this aircraft is on a conflicting flight path with any other
                    'If DetectCollisionPath(aircraft(i), aircraft(j)) = True Then
                        'Notify AreaTC of the aircraft involved in the conflict
                    '    lblAlert.Caption = aircraft(i).CallSign + " might soon be involved in a collision with " + aircraft(j).CallSign
                    '    lblAlert.Visible = True
                    'End If
                End If
            Next j
        Next i
    
    End If
    
 Next countVar
End Sub

