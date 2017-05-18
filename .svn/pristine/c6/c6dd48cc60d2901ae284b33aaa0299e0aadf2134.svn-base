Imports System.Data
Imports System.Data.SqlClient


Partial Class Source_00_MOA00107
    Inherits System.Web.UI.Page

    Public dDate     ' Date we're displaying calendar for
    Public iDIM      ' Days In Month
    Public iDOW      ' Day Of Week that month starts on
    Public iCurrent  ' Variable we use to hold current day of month as we write table
    Public iPosition ' Variable we use to hold current position in table
    Public iLooper   ' Variable used for misc. loops
    Public butColor1 = "#000000" '非假日
    Public butColor2 = "red" '假日
    Public dv As DataView
    Dim user_id, org_uid As String

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

    Function LoadCalendar(ByVal p As PlaceHolder, ByVal dDate As Date)
        iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
        iDOW = GetWeekdayMonthStartsOn(dDate)

        p.Controls.Add(New LiteralControl("<table border=0 bgcolor=#E0E0E0 width=605>"))
        p.Controls.Add(New LiteralControl("<tr bgcolor=#99cccc><td width=100% colspan=8>"))
        p.Controls.Add(New LiteralControl("<table border=0 bgcolor=#E0E0E0 width=100% CELLSPACING=0 CELLPADDING=0>"))
        p.Controls.Add(New LiteralControl("<tr bgcolor=#99cccc><td width=100%>"))
        p.Controls.Add(New LiteralControl("<a href=MOA00107.aspx?Date=" & dDate.AddMonths(-1) & ">"))
        p.Controls.Add(New LiteralControl("<img align='absmiddle' src=../../Image/arrowleft_w.gif title='上一個月' border=0></a>"))
        p.Controls.Add(New LiteralControl("<font size=4><b> " & Year(dDate) & " 年 " & MonthName(Month(dDate)) & "</b></font>"))
        p.Controls.Add(New LiteralControl("<a href=MOA00107.aspx?Date=" & dDate.AddMonths(1) & ">"))
        p.Controls.Add(New LiteralControl("<img align='absmiddle' src=../../Image/arrowright_w.gif title='下一個月' border=0></a>"))
        p.Controls.Add(New LiteralControl("</td><td width=20% align=right>"))
        p.Controls.Add(New LiteralControl("<form method=get action=/calendar/0,mid1283538 name=jumpdate>"))

        p.Controls.Add(New LiteralControl("<select name='yy'>"))
        For iLooper = Year(dDate) - 10 To Year(dDate) + 10
            p.Controls.Add(New LiteralControl("<option value='" & iLooper & "'"))
            If iLooper = Year(dDate) Then p.Controls.Add(New LiteralControl(" selected='selected'"))
            p.Controls.Add(New LiteralControl(">" & iLooper & "</option>"))
        Next
        p.Controls.Add(New LiteralControl("</select>"))

        p.Controls.Add(New LiteralControl("<select name='mm'>"))
        For iLooper = 1 To 12
            p.Controls.Add(New LiteralControl("<option value='" & iLooper & "'"))
            If iLooper = Month(dDate) Then p.Controls.Add(New LiteralControl(" selected='selected'"))
            p.Controls.Add(New LiteralControl(">"))
            p.Controls.Add(New LiteralControl(iLooper))
            p.Controls.Add(New LiteralControl("</option>" & vbCrLf))
        Next
        p.Controls.Add(New LiteralControl("</select>"))

        p.Controls.Add(New LiteralControl("<input type='submit' value='Go' onclick='CalendarChage()'/>"))
        p.Controls.Add(New LiteralControl("</td></form></tr></table></tr>"))
        p.Controls.Add(New LiteralControl("<tr bgcolor=#666666>"))
        p.Controls.Add(New LiteralControl("<td align=center width=80><font color=white size=2>日</font></td>"))
        p.Controls.Add(New LiteralControl("<td align=center width=80><font color=white size=2>一</font></td>"))
        p.Controls.Add(New LiteralControl("<td align=center width=80><font color=white size=2>二</font></td>"))
        p.Controls.Add(New LiteralControl("<td align=center width=80><font color=white size=2>三</font></td>"))
        p.Controls.Add(New LiteralControl("<td align=center width=80><font color=white size=2>四</font></td>"))
        p.Controls.Add(New LiteralControl("<td align=center width=80><font color=white size=2>五</font></td>"))
        p.Controls.Add(New LiteralControl("<td align=center width=80><font color=white size=2>六</font></td></tr>"))

        If iDOW <> 1 Then
            p.Controls.Add(New LiteralControl("<tr bgcolor=ffffff height=60><td colspan=" & iDOW - 1 & " bgcolor=e0e0e0></td>"))
        End If

        iCurrent = 1
        iPosition = iDOW
        Dim row As Int16 = 0

        Do While iCurrent <= iDIM
            If iPosition = 1 Then
                p.Controls.Add(New LiteralControl("<tr bgcolor=ffffff height=60>"))
            End If
            Dim sHref As String
            sHref = "<Label id='D" & CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)).ToString("yyyyMMdd") & "'  style='FONT-SIZE: 20px;color:" & butColor1 & "'><b>" & iCurrent & "</b></Label>"

            p.Controls.Add(New LiteralControl("<td align=left  valign=top>"))
            p.Controls.Add(New LiteralControl("<table border=0 cellpadding=0 cellspacing=0 width=100%>"))
            p.Controls.Add(New LiteralControl("<tr><td width=50%>"))
            p.Controls.Add(New LiteralControl(sHref))
            p.Controls.Add(New LiteralControl("<a href='MOA00108.aspx?Date=" & CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)) & "'>"))
            If Session("Role") = "1" Or Session("Role") = "2" Then
                p.Controls.Add(New LiteralControl("<img src=../../Image/memo.gif border=0 title='修改行事曆'></a> "))
            End If
            p.Controls.Add(New LiteralControl("</td><td valign=center align=right width=50%>"))
            p.Controls.Add(New LiteralControl("</td></tr></table>"))

            Dim Sql As String = "select NoteContent from P_1201 where ORG_UID='" & Session("ORG_UID") & "' and NoteDate='" & CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)) & "'"
            SqlDataSource1.SelectCommand = Sql
            dv = CType(SqlDataSource1.Select(DataSourceSelectArguments.Empty), DataView)
            Dim i As Integer = 0
            Dim count As Int16 = dv.Count
            Dim Str As String = ""
            Do While i < count
                If (dv.Table.Rows(i)(0).ToString.Length > 5) Then
                    Str = dv.Table.Rows(i)(0).ToString.Substring(0, 5)
                Else
                    Str = dv.Table.Rows(i)(0).ToString
                End If
                p.Controls.Add(New LiteralControl("<font color='blue' title='" & dv.Table.Rows(i)(0).ToString & "'>" & Str & "</font></br>"))
                i += 1
            Loop
            dv.Dispose()
            p.Controls.Add(New LiteralControl("</td>"))

            If iPosition = 7 Then
                p.Controls.Add(New LiteralControl("</tr>"))
                iPosition = 0
            End If

            iCurrent = iCurrent + 1
            iPosition = iPosition + 1
        Loop

        If iPosition <> 1 Then
            p.Controls.Add(New LiteralControl("<td colspan=" & 7 - iPosition & " bgcolor=e0e0e0></td>"))
            p.Controls.Add(New LiteralControl("</tr>"))
        End If

        p.Controls.Add(New LiteralControl("</td></table>"))

        LoadCalendar = ""
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00107") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00107.aspx")
                    Response.End()
                End If

                Dim AddDate As String = Request.QueryString("AddDate")

                Dim yy As Integer
                Dim mm As Integer
                Dim sql As String = ""

                If IsDate(Request.QueryString("Date")) Then
                    dDate = CDate(Request.QueryString("Date"))
                Else
                    dDate = Now()
                End If
                yy = Year(dDate)
                mm = Month(dDate)

                LoadCalendar(Me.FindControl("PlaceHolder1"), CDate(mm & "-01-" & yy))

                sql = "select fod_date,fod_year from P_12 where fod_year=" & yy & " and month(fod_date)=" & mm
                SqlDataSource1.SelectCommand = sql
                dv = CType(SqlDataSource1.Select(DataSourceSelectArguments.Empty), DataView)

            End If

        Catch ex As Exception

        End Try


    End Sub
End Class
