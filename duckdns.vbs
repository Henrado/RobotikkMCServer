Call LogEntry()

Sub LogEntry()
	On Error Resume Next
	Dim objRequest
	Dim URL

	URL = "https://www.duckdns.org/update?domains=mcrobotikk&token=77275d2a-6e23-410f-8796-2849762aae40&ip="

	Set objRequest = CreateObject("Microsoft.XMLHTTP")
	objRequest.open "GET", URL , false
	objRequest.Send
	Set objRequest = Nothing
End Sub