VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   120
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image Image1 
         Height          =   495
         Left            =   5640
         Picture         =   "CATCS_Splash.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Loading . . ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   2580
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "For Microsoft Windows 2000/XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4275
         TabIndex        =   6
         Top             =   2400
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MURTALA MOHAMMED INT'L AIRPORT, IKEJA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   840
         TabIndex        =   7
         Top             =   2160
         Width           =   5415
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "Air Traffic Control System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   645
         TabIndex        =   8
         Top             =   1800
         Width           =   5805
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1305
         Left            =   2880
         Picture         =   "CATCS_Splash.frx":88D6
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   3180
         Width           =   1695
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         Caption         =   "(c) CoDemon Inc. 2002 -2006"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4800
         TabIndex        =   3
         Top             =   3390
         Width           =   2115
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         Caption         =   "Warning: This computer program is protected by copyright law. Do not reproduce illegally."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   3660
         Width           =   6495
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: CATCS (Computerized Air Traffic Control System) for
'Murtala Mohammed International Airport (MMIA), Ikeja.
'Authored by: OGUNBANJO David Oluseyi as a Final Year (BSc.) Project at
'IGBINEDION UNIVERSITY, OKADA, EDO STATE, NIGERIA.
'Project Start Date: Wed. 12th July, 2006.
'Project Completion Date: Tues. 19th July, 2006.


Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.MousePointer = 11 'vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmLogin.Show
End Sub

Private Sub Frame1_Click()
    'Unload Me
End Sub

Private Sub Timer1_Timer()
 Unload Me
End Sub

Private Sub Timer2_Timer()
 If ProgressBar1 <= 98 Then ProgressBar1 = ProgressBar1 + 2
End Sub
