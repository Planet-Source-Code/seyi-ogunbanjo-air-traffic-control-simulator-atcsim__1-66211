Attribute VB_Name = "modRoutines"
'Project: CATCS (Computerized Air Traffic Control System) for
'Murtala Mohammed International Airport, Ikeja.
'Authored by: OGUNBANJO David Oluseyi as a Final Year (BSc.) Project at
'IGBINEDION UNIVERSITY, OKADA, EDO STATE, NIGERIA.
'Project Start Date: Wed. 12th July, 2006.
'Project Completion Date: Tues. 19th July, 2006.

Option Explicit

Public Function CheckForCollision(a1 As CPlane, a2 As CPlane) As Boolean
 If (a1.X < a2.X) And (a1.X + a1.W > a2.X) And (a1.Y < a2.Y) And (a1.Y + a1.H > a2.Y) Or _
    (a2.X < a1.X) And (a2.X + a2.W > a1.X) And (a2.Y < a1.Y) And (a2.Y + a2.H > a1.Y) Then
        CheckForCollision = True
 Else
        CheckForCollision = False
 End If
End Function

Public Sub ReleaseOutgoingAircraft(a As CPlane, indx As Integer)
 'indx is the index of the current airplane. is used to hide the corresponding image control off the screen.
 
 'Check if aircraft is flying off the screen (out of controlled airspace)
 With frmMain
    If (a.X < 0) Or (a.X > .Width) Or (a.Y < 0) Or (a.Y > 10100) Then
    
        'Update flight status information
        a.FlightState = 0
                
        'Remove Highlight from aircraft descriptors
        .imgPlanePic(indx).BorderStyle = 0     'NONE
        
        'make aircraft image disappear
        .imgPlanePic(indx).Visible = False
            
        'Notify AreaTC that aircraft has been transferred to adjacent airspace
        .txtMessage.Text = a.CallSign + " has been transferred to an adjacent Area Control Center."
        
        'Log the transfer in the log file
        Write #1, Str(Time), a.CallSign + " transferred to adjacent ACC."
        
        'Clear data of aircraft that just flew out of our airspace from text boxes in "Control Panel"
        If .txtCallSign = a.CallSign Then
            .txtCallSign = ""
            .txtHeading = ""
            .txtAltitude = ""
            .txtEntry = ""
            .txtExit = ""
        End If
    
        'Update DB to reflect transfer
        With .datFlightDesc.Recordset
            .MoveFirst
            'search DB for this aircraft's record
            Do While (.EOF = False) And (.Fields("Call_Sign") <> a.CallSign)
                .MoveNext
            Loop
            .Edit
            '.Fields("FlightState") = a.FlightState
            .Update
        End With
    End If
 End With
End Sub

Public Sub ReleaseLandingAircraft(a As CPlane, indx As Integer)
 'indx is the index of the current airplane. is used to hide the corresponding image control off the screen.
 If a.Altitude > 300 Then Exit Sub      'i.e, planes must fly below 300nmi to be within approach control's coverage.
 
 With frmMain
    'Check if aircraft is flying over landing beacon
    If (a.X > .lblLOSLanding.Left) And (a.X < .lblLOSLanding.Left + .lblLOSLanding.Width) And _
        (a.Y > .lblLOSLanding.Top) And (a.Y < .lblLOSLanding.Top + .lblLOSLanding.Height) Or _
        (.lblLOSLanding.Left > a.X) And (.lblLOSLanding.Left < a.X + a.W) And _
        (.lblLOSLanding.Top > a.Y) And (.lblLOSLanding.Top < a.Y + a.H) Then
        'i.e. aircraft object is inside landing beacon
        
        'Update flight status information
        a.FlightState = 0
                
        'Remove Highlight from aircraft descriptors
        .imgPlanePic(indx).BorderStyle = 0     'NONE
                
        'make aircraft image disappear
        .imgPlanePic(indx).Visible = False
            
        'Notify AreaTC that aircraft has been transferred to approach control
        .txtMessage.Text = a.CallSign + " has been transferred to Approach Control."
                
        'Log the Landing in the log file
        Write #1, Str(Time), a.CallSign + " transferred to Approach Control for Landing."
        
        'Clear data of aircraft that just flew out of our airspace from text boxes in "Control Panel"
        If .txtCallSign = a.CallSign Then
            .txtCallSign = ""
            .txtHeading = ""
            .txtAltitude = ""
            .txtEntry = ""
            .txtExit = ""
        End If

        'Update DB to reflect transfer
        With .datFlightDesc.Recordset
            .MoveFirst
            'search DB for this aircraft's record
            Do While (.EOF = False) And (.Fields("Call_Sign") <> a.CallSign)
                .MoveNext
            Loop
            .Edit
            '.Fields("FlightState") = a.FlightState
            .Update
        End With
    End If
 End With
End Sub

Public Function DetectCollisionPath(a1 As CPlane, a2 As CPlane) As Boolean
 Dim rX As Integer
 Dim rY As Integer
 Dim rH As Integer
 Dim rW As Integer
 
 rX = a1.X - 500
 rY = a1.Y - 500
 rH = a1.H + 1000
 rW = a1.W + 1000
 
 If (rX < a2.X) And (rX + rW > a2.X) And (rY < a2.Y) And (rY + rH > a2.Y) Or _
    (a2.X < rX) And (a2.X + a2.W > rX) And (a2.Y < rY) And (a2.Y + a2.H > rY) Then
        DetectCollisionPath = True
 Else
        DetectCollisionPath = False
 End If
 
End Function
