Dim objOL   'As Outlook.Application
    Dim objAppt 'As Outlook.AppointmentItem
    Const olAppointmentItem = 1
    Const olMeeting = 1
    
    Set objOL = CreateObject("Outlook.Application")
    Set objAppt = objOL.CreateItem(olAppointmentItem)
	
strInput = InputBox( "Enter Subject:")
strInput2 = InputBox("Enter Date:","",Now() )

	
    With objAppt
        .Subject = strInput
	.Body = strInput
    		.Start = strInput2
	.End = DateAdd("h", 1, .Start)
	       
        ' make it a meeting request
        .MeetingStatus = olMeeting
        .RequiredAttendees = ""
        .Send
    End With
    
    Set objAppt = Nothing
    Set objOL = Nothing
