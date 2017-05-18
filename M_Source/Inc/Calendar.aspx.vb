Imports System.Data
Imports System.Data.SqlClient


Partial Class calendar
    Inherits System.Web.UI.Page

    Public sTextBoxID As String
    Public dDate     ' Date we're displaying calendar for
    Public iDIM      ' Days In Month
    Public iDOW      ' Day Of Week that month starts on
    Public iCurrent  ' Variable we use to hold current day of month as we write table
    Public iPosition ' Variable we use to hold current position in table
    Public iLooper   ' Variable used for misc. loops
    Public return_value As String



    Function GetDaysInMonth(ByVal iMonth, ByVal iYear)
        Dim dTemp
        dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
        GetDaysInMonth = Day(dTemp)
    End Function
    Function GetWeekdayMonthStartsOn(ByVal dAnyDayInTheMonth)
        Dim dTemp
        dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
        GetWeekdayMonthStartsOn = Weekday(dTemp)
    End Function

    Function SubtractOneMonth(ByVal dDate)
        SubtractOneMonth = DateAdd(DateInterval.Month, -1, dDate)
    End Function

    Function AddOneMonth(ByVal dDate)
        AddOneMonth = DateAdd("m", 1, dDate)
    End Function

    Function LoadCalendar(ByVal dDate)
        iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
        iDOW = GetWeekdayMonthStartsOn(dDate)

        PlaceHolder1.Controls.Add(New LiteralControl("<table align='center' border='0' cellspacing='0' cellpadding='0'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr><td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<table border='0' cellspacing='0' cellpadding='1' bgcolor='#ffffcc' width='100%'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td bgcolor='#000099' align='center' colspan='7'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<table width='100%' border='0' cellspacing='0' cellpadding='0'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='right'><a href=""calendar.aspx?TextBoxId=" & sTextBoxID & "&date=" & SubtractOneMonth(dDate) & """><font color='#FFFF00' size='-1'>&lt;&lt;</font></a></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<select name='year' onchange='CalendarChage()'>"))
        For iLooper = Year(dDate) - 10 To Year(dDate) + 10
            PlaceHolder1.Controls.Add(New LiteralControl("<option value='" & iLooper & "'"))
            If iLooper = Year(dDate) Then PlaceHolder1.Controls.Add(New LiteralControl(" selected='selected'"))
            PlaceHolder1.Controls.Add(New LiteralControl(">" & iLooper & "</option>"))
        Next
        PlaceHolder1.Controls.Add(New LiteralControl("</select>"))

        PlaceHolder1.Controls.Add(New LiteralControl("<font color='#FFFF00'><strong>年</strong></font>"))

        PlaceHolder1.Controls.Add(New LiteralControl("<select name='month' onchange='CalendarChage()'>"))
        PlaceHolder1.Controls.Add(New LiteralControl(vbCrLf))
        For iLooper = 1 To 12
            PlaceHolder1.Controls.Add(New LiteralControl(vbTab & vbTab & vbTab & vbTab))
            PlaceHolder1.Controls.Add(New LiteralControl("<option value='" & iLooper & "'"))
            If iLooper = Month(dDate) Then PlaceHolder1.Controls.Add(New LiteralControl(" selected='selected'"))
            PlaceHolder1.Controls.Add(New LiteralControl(">"))
            PlaceHolder1.Controls.Add(New LiteralControl(iLooper))
            PlaceHolder1.Controls.Add(New LiteralControl("</option>" & vbCrLf))
        Next
        PlaceHolder1.Controls.Add(New LiteralControl("</select>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<font color='#FFFF00'><strong>月</strong></font>"))

        PlaceHolder1.Controls.Add(New LiteralControl("<div style='display:none'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<select name='day'>"))
        For iLooper = 1 To 31
            PlaceHolder1.Controls.Add(New LiteralControl("<option value='" & iLooper & "'"))
            If iLooper = Day(dDate) Then PlaceHolder1.Controls.Add(New LiteralControl(" selected='selected'"))
            PlaceHolder1.Controls.Add(New LiteralControl(">" & iLooper & "</option>"))
        Next
        PlaceHolder1.Controls.Add(New LiteralControl("</select>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<font color='#FFFF00'><strong>日</strong></font></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("</div>"))


        PlaceHolder1.Controls.Add(New LiteralControl("<td align='left' ><a href=""calendar.aspx?TextBoxId=" & sTextBoxID & "&date=" & AddOneMonth(dDate) & """><font color='#FFFF00' size='-1'>&gt;&gt;</font></a></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("</tr></table></td></tr><tr>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>日</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>一</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>二</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>三</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>四</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>五</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>六</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("</tr>"))
        If iDOW <> 1 Then
            PlaceHolder1.Controls.Add(New LiteralControl(vbTab & "<tr>" & vbCrLf))
            iPosition = 1
            Do While iPosition < iDOW
                PlaceHolder1.Controls.Add(New LiteralControl(vbTab & vbTab & "<td>&nbsp;</td>" & vbCrLf))
                iPosition = iPosition + 1
            Loop
        End If

        iCurrent = 1
        iPosition = iDOW
        Do While iCurrent <= iDIM
            If iPosition = 1 Then
                PlaceHolder1.Controls.Add(New LiteralControl(vbTab & "<tr>" & vbCrLf))
            End If
            Dim sHref As String
            Dim sFont As String
            Dim sBgcolor As String
            sHref = "<a href=""javascript:LoadCalendar('" & CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)) & "')"" style='text-decoration: none;'>"
            sFont = "<font size='3'>" & iCurrent & "</font>"
            sBgcolor = ""
            If iPosition = 1 Or iPosition = 7 Then
                sFont = "<font size='3' Color='red'>" & iCurrent & "</font>"
            End If
            If iCurrent = Day(dDate) Then
                sBgcolor = "bgcolor='#00FFFF'"
            Else
            End If
            PlaceHolder1.Controls.Add(New LiteralControl(vbTab & vbTab & "<td " & sBgcolor & ">" & sHref & sFont & "</a></td>" & vbCrLf))

            If iPosition = 7 Then
                PlaceHolder1.Controls.Add(New LiteralControl(vbTab & "</tr>" & vbCrLf))
                iPosition = 0
            End If

            iCurrent = iCurrent + 1
            iPosition = iPosition + 1
        Loop

        If iPosition <> 1 Then
            Do While iPosition <= 7
                PlaceHolder1.Controls.Add(New LiteralControl(vbTab & vbTab & "<td>&nbsp;</td>" & vbCrLf))
                iPosition = iPosition + 1
            Loop
            PlaceHolder1.Controls.Add(New LiteralControl(vbTab & "</tr>" & vbCrLf))
        End If
        PlaceHolder1.Controls.Add(New LiteralControl("</table></td></tr></table>"))

        LoadCalendar = ""
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'server端 呼叫月曆的寫法
        'Dim sScript As String
        'Dim sPath As String
        'Dim strFeatures As String
        'sPath = "../calendar.aspx?TextBoxId=" & Me.FindControl("Sdate").ClientID
        'strFeatures = "dialogWidth=250px;dialogHeight=168px;help=no;status=no;resizable=yes;scroll=no;dialogTop="+y+";dialogLeft="+x;
        'sScript = "showModalDialog('" & sPath & "',self,'" & strFeatures & "')"
        'CalBtn1.Attributes("onclick") = sScript


        '取得要輸入日期的 TextBox   
        sTextBoxID = Request.QueryString("TextBoxId")
        return_value = Request.QueryString("return_value")
        'server端 將日期設給 TextBox，並將視窗關閉的寫法
        'Dim sScript As String
        'sScript = "window.dialogArguments.document.getElementById('" & sTextBoxID & "').value='" & Calendar1.SelectedDate.Date & "';"
        'sScript = "window.close();"
        'Me.ClientScript.RegisterStartupScript(Me.GetType(), "_Calendar", sScript, True)


        If IsDate(Request.QueryString("date")) Then
            dDate = CDate(Request.QueryString("date"))
        Else
            If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
                dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
            Else
                dDate = Now()
                If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
                    Response.Write("The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>")
                End If
            End If
        End If
        LoadCalendar(dDate)
    End Sub
End Class
