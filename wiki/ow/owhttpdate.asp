<%
' Functions in this file are initially taken from http://www.motobit.com/tips/detpg_DateToHTTPDate/

' Converts date (19991022 11:08:38)
' to http form (Fri, 22 Oct 1999 12:08:38 GMT)
function DateToHTTPDate(ByVal OleDATE)
  OleDATE = OleDATE
	DateToHTTPDate = engWeekDayName(OleDATE) & _
	", " & Right("0" & Day(OleDATE),2) & " " & engMonthName(OleDATE) & _
	" " & Year(OleDATE) & " " & Right("0" & Hour(OleDATE),2) & _
	":" & Right("0" & Minute(OleDATE),2) & ":" & Right("0" & Second(OleDATE),2) & " GMT"
End Function 

Function engWeekDayName(dt)
	Dim Out
	Select Case WeekDay(dt,1)
		Case 1:Out="Sun"
		Case 2:Out="Mon"
		Case 3:Out="Tue"
		Case 4:Out="Wed"
		Case 5:Out="Thu"
		Case 6:Out="Fri"
		Case 7:Out="Sat"
	End Select
	engWeekDayName = Out
End Function

Function engMonthName(dt)
	Dim Out
	Select Case Month(dt)
		Case 1:Out="Jan"
		Case 2:Out="Feb"
		Case 3:Out="Mar"
		Case 4:Out="Apr"
		Case 5:Out="May"
		Case 6:Out="Jun"
		Case 7:Out="Jul"
		Case 8:Out="Aug"
		Case 9:Out="Sep"
		Case 10:Out="Oct"
		Case 11:Out="Nov"
		Case 12:Out="Dec"
	End Select
	engMonthName = Out
End Function
%>