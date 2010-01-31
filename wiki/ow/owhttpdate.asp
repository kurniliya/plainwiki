<%
' Functions in this file are initially taken from http://www.motobit.com/tips/detpg_DateToHTTPDate/
Const GMTDiff = #04:00:00#

' Converts date (19991022 11:08:38)
' to http form (Fri, 22 Oct 1999 12:08:38 GMT)
Function DateToHTTPDate(ByVal OleDATE)
  OleDATE = OleDATE - GMTDiff
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

Public Function DateFromHTTPDate(HTTPDate)
  Dim Swd, d, Sm, y, h, m, s, g, Out
  HTTPDate = LCase(HTTPDate)

  If Mid(HTTPDate, 27, 3) = "gmt" Then
    Swd = Left(HTTPDate, 3)
    d = Mid(HTTPDate, 6, 2)
    Sm = Mid(HTTPDate, 9, 3)
    y = Mid(HTTPDate, 13, 4)
    h = Mid(HTTPDate, 18, 2)
    m = Mid(HTTPDate, 21, 2)
    s = Mid(HTTPDate, 24, 2)
'    on error resume Next
    Out = DateSerial(y, mFromSm(Sm), d) + TimeSerial(h, m, s) + GMTDiff
'    on error goto 0
  End If
  DateFromHTTPDate = Out
End Function

Function wdFromSwd(Swd)
  Dim Out
  Select Case LCase(Swd)
    Case "sun": Out = 1: Case "mon": Out = 2: Case "tue": Out = 3: Case "wed": Out = 4: Case "thu": Out = 5: Case "fri": Out = 6: Case "sat": Out = 7
  End Select
  wdFromSwd = Out
End Function

Function mFromSm(Sm)
  Dim Out
  Select Case LCase(Sm)
    Case "jan": Out = 1: Case "feb": Out = 2: Case "mar": Out = 3: Case "apr": Out = 4
    Case "may": Out = 5: Case "jun": Out = 6: Case "jul": Out = 7: Case "aug": Out = 8
    Case "sep": Out = 9: Case "oct": Out = 10: Case "nov": Out = 11: Case "dec": Out = 12
  End Select
  mFromSm = Out
End Function

Function GetLatestDate(Date1, Date2)
	If DateDiff("s", Date1, Date2) > 0 Then
		GetLatestDate = Date2
	Else
		GetLatestDate = Date1
	End If
End Function

Function SecondDateIsLater(Date1, Date2)
	If DateDiff("s", Date1, Date2) >= 0  Then
		SecondDateIsLater = True
	Else
		SecondDateIsLater = False
	End If	
End Function
%>