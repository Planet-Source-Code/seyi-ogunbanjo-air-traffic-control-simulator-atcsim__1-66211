Attribute VB_Name = "modGlobalVars"
'Project: CATCS (Computerized Air Traffic Control System) for
'Murtala Mohammed International Airport (MMIA), Ikeja.
'Authored by: OGUNBANJO David Oluseyi as a Final Year (BSc.) Project at
'IGBINEDION UNIVERSITY, OKADA, EDO STATE, NIGERIA.
'Project Start Date: Wed. 12th July, 2006.
'Project Completion Date: Tues. 19th July, 2006.

Public LoginSucceeded As Boolean        'Login status variable
Public maxAircraft As Integer
Public simSpeed As Integer              'Speed of the simulation; is set to 1 at the start (when frmMain is loaded)
Public currentUser As String               'Name of the current system user
Option Explicit
'General Declarations Module


Public Sub GetEntryLocation(routeCode As String, XVal As Integer, YVal As Integer)
 'Map an "entry route code" to co-ordinates on the control airspace
    
 Select Case UCase(routeCode)
    Case "A0"
        XVal = 240
        YVal = 0
        
    Case "A1"
        XVal = 1920
        YVal = 600
        
    Case "A2"
        XVal = 3840
        YVal = 600
        
    Case "A3"
        XVal = 5880
        YVal = 600
        
    Case "A4"
        XVal = 7800
        YVal = 600
        
    Case "A5"
        XVal = 9840
        YVal = 600
        
    Case "A6", "ABJ"  'ABJ
        XVal = 11880
        YVal = 600
        
    Case "A7"
        XVal = 13680
        YVal = 0
        
    Case "F0"
        XVal = 240
        YVal = 9600
        
    Case "F1"
        XVal = 1920
        YVal = 9600
        
    Case "F2"
        XVal = 3840
        YVal = 9600
        
    Case "F3"
        XVal = 5880
        YVal = 9600
        
    Case "F4"
        XVal = 7800
        YVal = 9600
        
    Case "F5"
        XVal = 9840
        YVal = 9600
        
    Case "F6"
        XVal = 11800
        YVal = 9600
        
    Case "F7"
        XVal = 13920
        YVal = 9600
        
    Case "B0"
        XVal = 240
        YVal = 1920       'Change to 1120 to test collision path wahala
        'Also, change ARC213CRBJ's altitude from 370 to 310 (in the database)
        'Although, the collision thingy is not yet complete... :)
        
    Case "C0"
        XVal = 240
        YVal = 3840
        
    Case "D0"
        XVal = 240
        YVal = 5880
        
    Case "E0"
        XVal = 240
        YVal = 7920
        
    Case "B7"
        XVal = 13680
        YVal = 1920
        
    Case "C7"
        XVal = 13680
        YVal = 3840

    Case "D7", "BEN"  'BEN
        XVal = 13680
        YVal = 5880

    Case "E7", "PHC"  'PHC
        XVal = 13680
        YVal = 7920

    Case "NWA"
        XVal = 840
        YVal = 720

    Case "GHA"
        XVal = 840
        YVal = 6600
        
    Case "SA"
        XVal = 2520
        YVal = 9480
        
    Case "LOS"  'Taking -off from Lagos
        XVal = 4080
        YVal = 6960
        
    Case Else
        XVal = 0
        YVal = 0
    
 End Select


End Sub
