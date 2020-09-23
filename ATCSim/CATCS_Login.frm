VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CATCS - Login"
   ClientHeight    =   6075
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4350
   Icon            =   "CATCS_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3589.309
   ScaleMode       =   0  'User
   ScaleWidth      =   4084.415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDifficulty 
      Caption         =   "Difficulty Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   3615
      Begin VB.OptionButton optLevel 
         Caption         =   "Difficult (More than 10  Aircraft)"
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
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   3255
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Medium (5 - 10 Aircraft)"
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
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Easy (5 Aircraft)"
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
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Data datFlightDesc 
      Caption         =   "Control Database Users Table"
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
      Top             =   4320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ComboBox cboUserName 
      Height          =   315
      ItemData        =   "CATCS_Login.frx":030A
      Left            =   1560
      List            =   "CATCS_Login.frx":0317
      TabIndex        =   5
      Text            =   "administrator"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   735
      TabIndex        =   3
      Top             =   3900
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   390
      Left            =   2340
      TabIndex        =   4
      Top             =   3900
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "351"
      Top             =   3405
      Width           =   2325
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MMIA - CONTROL TERMINAL"
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
      Left            =   495
      TabIndex        =   8
      Top             =   1800
      Width           =   3360
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "CATCS"
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
      Height          =   435
      Left            =   1553
      TabIndex        =   7
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   1508
      Picture         =   "CATCS_Login.frx":0332
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"CATCS_Login.frx":10C3
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
      Height          =   795
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   3675
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   3030
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   345
      TabIndex        =   1
      Top             =   3420
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
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


Dim loginAttempts As Integer
Option Explicit

Private Sub cmdCancel_Click()
 If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Quit?") = vbYes Then
    Unload Me
    End
 End If
End Sub

Private Sub cmdOK_Click()
 'check for correct password
 loginAttempts = loginAttempts - 1
 LoginSucceeded = False
 With datFlightDesc.Recordset
    .MoveFirst
    Do While Not .EOF And Not LoginSucceeded
        If (UCase(cboUserName) = UCase(.Fields("user"))) And Trim(txtPassword) = .Fields("pass") Then
            'Login authenticated
            MsgBox "Login successful. Welcome to the system.", vbInformation
            LoginSucceeded = True
            currentUser = UCase(cboUserName)
            If optLevel(0).Value = True Then
                'Easy
                maxAircraft = 4
            ElseIf optLevel(1) = True Then
                'Medium
                maxAircraft = 9
            Else
                'Hard
                maxAircraft = 19
            End If
            Unload Me
            frmMain.Show
        End If
        .MoveNext
    Loop
    If (LoginSucceeded = False) And (loginAttempts > 0) Then
        MsgBox "Login failed; You have " & loginAttempts & " more attempts!", vbInformation, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If LoginSucceeded = False And loginAttempts = 0 Then
        MsgBox "Unauthorized personnel are not allowed to access this system." + vbCrLf + "The system will now exit." + vbCrLf + "Goodbye.", vbCritical, "Login Failed"
        End
    End If
 End With
End Sub

Private Sub Form_Load()
 datFlightDesc.DatabaseName = App.Path & "\DBFiles\FlightInfo97.mdb"
 datFlightDesc.RecordSource = "CATCS_Users"
 datFlightDesc.Refresh
 loginAttempts = 3
End Sub

