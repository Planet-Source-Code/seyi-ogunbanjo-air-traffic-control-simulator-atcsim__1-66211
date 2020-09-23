VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: CATCS (Computerized Air Traffic Control System) for
'Murtala Mohammed International Airport (MMIA), Ikeja.
'Authored by: OGUNBANJO David Oluseyi as a Final Year (BSc.) Project at
'IGBINEDION UNIVERSITY, OKADA, EDO STATE, NIGERIA.
'Project Start Date: Wed. 12th July, 2006.
'Project Completion Date: Tues. 19th July, 2006.


Private m_imgPlane As Image     'Image control associated with plane;
                                'Is a write-only property
Private m_CallSign As String
Private m_Heading As Integer
Private m_Altitude As Integer
Private m_EntryRoute As String
Private m_ExitRoute As String
Private m_FlightState As Integer        '1 = flying within the controlled airspace,
                                        '0 = has flown out of the controlled airspace
                                        '-1 = has crashed into another aircraft
Private m_CrashReport As Integer        '0 = Not yet involved in a crash
                                        'Crash detected and reported.

Option Explicit
'Class Module for CPlane

Property Let imagePlane(newPlane As Image)
 Set m_imgPlane = newPlane
End Property

Public Property Get CallSign() As String
 CallSign = m_CallSign
End Property

Public Property Let CallSign(ByVal vNewCallSign As String)
 m_CallSign = vNewCallSign
End Property

Public Property Get Heading() As Integer
 Heading = m_Heading
End Property

Public Property Let Heading(ByVal vNewHeading As Integer)
 m_Heading = vNewHeading
End Property

Public Property Get Altitude() As Integer
 Altitude = m_Altitude
End Property

Public Property Let Altitude(ByVal vNewAltitude As Integer)
 m_Altitude = vNewAltitude
End Property

Public Property Get EntryRoute() As String
 EntryRoute = m_EntryRoute
End Property

Public Property Let EntryRoute(ByVal vNewEntryRoute As String)
 m_EntryRoute = vNewEntryRoute
End Property

Public Property Get ExitRoute() As String
 ExitRoute = m_ExitRoute
End Property

Public Property Let ExitRoute(ByVal vNewExitRoute As String)
 m_ExitRoute = vNewExitRoute
End Property

Public Property Get FlightState() As Integer
 FlightState = m_FlightState
End Property

Public Property Let FlightState(ByVal vNewFlightState As Integer)
 m_FlightState = vNewFlightState
End Property

Public Function X() As Integer
 X = m_imgPlane.Left
End Function

Public Function Y() As Integer
 Y = m_imgPlane.Top
End Function

Public Function W() As Integer
 W = m_imgPlane.Width
End Function

Public Function H() As Integer
 H = m_imgPlane.Height
End Function

Public Sub fly(ByVal bearing As Integer)
 'm_imgPlane.Visible = True
 
 'bearing is the direction of airplane object(in degrees)
 Select Case bearing
 Case 0
        'aircraft is headed NORTH
        m_imgPlane.Top = m_imgPlane.Top - 12 * 2 ^ simSpeed
 Case 45
        'aircraft is headed NORTH-EAST
        m_imgPlane.Top = m_imgPlane.Top - 6 * 2 ^ simSpeed
        m_imgPlane.Left = m_imgPlane.Left + 6 * 2 ^ simSpeed
 Case 90
        'aircraft is headed EAST
        m_imgPlane.Left = m_imgPlane.Left + 12 * 2 ^ simSpeed
 Case 135
        'aircraft is headed SOUTH-EAST
        m_imgPlane.Top = m_imgPlane.Top + 6 * 2 ^ simSpeed
        m_imgPlane.Left = m_imgPlane.Left + 6 * 2 ^ simSpeed
 Case 180
        'aircraft is headed SOUTH
        m_imgPlane.Top = m_imgPlane.Top + 12 * 2 ^ simSpeed
 Case 225
        'aircraft is headed SOUTH-WEST
        m_imgPlane.Top = m_imgPlane.Top + 6 * 2 ^ simSpeed
        m_imgPlane.Left = m_imgPlane.Left - 6 * 2 ^ simSpeed
 Case 270
        'aircraft is headed WEST
        m_imgPlane.Left = m_imgPlane.Left - 12 * 2 ^ simSpeed
 Case 315
        'aircraft is headed NORTH-WEST
        m_imgPlane.Top = m_imgPlane.Top - 6 * 2 ^ simSpeed
        m_imgPlane.Left = m_imgPlane.Left - 6 * 2 ^ simSpeed
 
 End Select
End Sub

Private Sub Class_Initialize()
 m_CrashReport = 0
End Sub

Private Sub Class_Terminate()
 Set m_imgPlane = Nothing
End Sub

Public Property Get ReportCrash() As Integer
 ReportCrash = m_CrashReport
End Property

Public Property Let ReportCrash(ByVal vReportCrash As Integer)
 m_CrashReport = vReportCrash
End Property
