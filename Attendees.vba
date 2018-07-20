Sub Attendees()

'Set Variables 
Dim objApp As Outlook.Application
Dim objItem As Object
Dim objAttendees As Outlook.Recipients
Dim objAttendeeReq As String
Dim objAttendeeOpt As String
Dim dtStart As Date
Dim dtEnd As Date
Dim strSubject As String
Dim strLocation As String
Dim strNotes As String
Dim strMeetStatus As String
Dim strCopyData As String
Dim strCount  as String
Dim strAttendeeCost as Integer
Dim totalExpense as Integer
Dim strAttendeeList as String
Dim MeetingDur as Outlook.Duration
Dim TotalDur as Integer

On Error Resume Next
 
Set objApp = CreateObject("Outlook.Application")
Set objItem = GetCurrentItem()
Set objAttendees = objItem.Recipients
 
' Get the data
dtStart = objItem.Start
dtEnd = objItem.End
'Reset variables
objAttendeeReq = ""
objAttendeeOpt = ""
strAttendeeList = ""
MeetingDur = ""

' Get Meeting Duration
MeetingDur = objItem.Duration
' Get The Attendee List
For x = 1 To objAttendees.Count
   strMeetStatus = ""
   Select Case objAttendees(x).MeetingResponseStatus
     Case 0
       strMeetStatus = "No Response (or Organizer)"
       ino = ino + 1
     Case 1
       strMeetStatus = "Organizer"
       ino = ino + 1
     Case 2
       strMeetStatus = "Tentative"
       it = it + 1
     Case 3
       strMeetStatus = "Accepted"
       ia = ia + 1
     Case 4
       strMeetStatus = "Declined"
       ide = ide + 1
   End Select
  
   If objAttendees(x).Type = olRequired Then
      objAttendeeReq = objAttendees(x).Name
	  'Append Attendee to List
	   strAttendeeList = strAttendeeList & ";" & objAttendeeReq
	   'Evaluate Cost of Attendee
		Dim Minute From HR_Data In HR_DATA
                  Where HR_DATA.ED_EMP_NAME = objAttendeeReq
	    totalExpense = totalExpense + HR_DATA.Minute
   Else
      objAttendeeOpt = objAttendees(x).Name
	  'Append Attendee to List
       strAttendeeList = strAttendeeList & ";" & objAttendeeOpt
	   	'Evaluate Cost of Attendee
	   	 Dim Minute From HR_Data In HR_DATA
                  Where HR_DATA.ED_EMP_NAME = objAttendeeOpt
	     totalExpense = totalExpense + HR_DATA.Minute
   End If




Next

totalExpense = totalExpense * MeetingDur

End Sub
