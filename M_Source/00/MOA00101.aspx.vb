Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_00_MOA00101
    Inherits System.Web.UI.Page

    Public dDate     ' Date we're displaying calendar for
    Public iDIM      ' Days In Month
    Public iDOW      ' Day Of Week that month starts on
    Public iCurrent  ' Variable we use to hold current day of month as we write table
    Public iPosition ' Variable we use to hold current position in table
    Public iLooper   ' Variable used for misc. loops
    Public butColor1 = "#99cccc" '非假日
    Public butColor2 = "#ff6666" '假日
    Public dv As DataView
    Public Shared HistoryGo As Int16
    Public yy As Integer
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

        p.Controls.Add(New LiteralControl("<table  topmargin='0' align='center' border='0' cellspacing='0' cellpadding='0'>"))
        p.Controls.Add(New LiteralControl("<tr><td>"))
        p.Controls.Add(New LiteralControl("<table align='center' border='0' cellspacing='0' cellpadding='1' width='100%'>"))
        p.Controls.Add(New LiteralControl("<tr>"))
        p.Controls.Add(New LiteralControl("<td align='center' colspan='7'>"))
        p.Controls.Add(New LiteralControl("<table width='100%' border='0' cellspacing='0' cellpadding='0'>"))
        p.Controls.Add(New LiteralControl("<tr>"))
        p.Controls.Add(New LiteralControl("<td align='center'>"))
        p.Controls.Add(New LiteralControl("<font color='#0000CC'><strong>" & Year(dDate) & "年</strong></font>"))
        p.Controls.Add(New LiteralControl(vbCrLf))
        p.Controls.Add(New LiteralControl("<font color='#0000CC'><strong>" & Month(dDate) & "月</strong></font>"))
        p.Controls.Add(New LiteralControl("</tr></table></td></tr><tr>"))
        p.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>日</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        p.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>一</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        p.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>二</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        p.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>三</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        p.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>四</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        p.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>五</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        p.Controls.Add(New LiteralControl("<td align='center' bgcolor='#0000CC'><font color='#FFFF00'><strong>六</strong></font><br /><img src='images/spacer.gif' width='14%' height='1' border='0'></td>"))
        p.Controls.Add(New LiteralControl("</tr>"))
        If iDOW <> 1 Then
            p.Controls.Add(New LiteralControl("<tr>"))
            iPosition = 1
            Do While iPosition < iDOW
                p.Controls.Add(New LiteralControl("<td>&nbsp;</td>"))
                iPosition = iPosition + 1
            Loop
        End If

        iCurrent = 1
        iPosition = iDOW
        Dim row As Int16 = 0

        Do While iCurrent <= iDIM
            If iPosition = 1 Then
                p.Controls.Add(New LiteralControl("<tr>"))
            End If
            Dim sHref As String
            sHref = "<input type='submit' name='D" & CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)).ToString("yyyyMMdd") & "' value='" & iCurrent & "' style='width:100%;background:" & butColor1 & ";' onclick=""CalendarChage(this,'" & CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)) & "')""/>"
            p.Controls.Add(New LiteralControl("<td>" & sHref & "</td>"))

            If iPosition = 7 Then
                p.Controls.Add(New LiteralControl("</tr>"))
                iPosition = 0
                row += 1
            End If

            iCurrent = iCurrent + 1
            iPosition = iPosition + 1
        Loop

        If iPosition <> 1 Then
            Do While iPosition <= 7
                p.Controls.Add(New LiteralControl("<td>&nbsp;</td>"))
                iPosition = iPosition + 1
            Loop
            p.Controls.Add(New LiteralControl("</tr>"))
            row += 1
        End If
        Do While row <= 5
            p.Controls.Add(New LiteralControl("<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"))
            row += 1
        Loop
        p.Controls.Add(New LiteralControl("</table></td></tr></table>"))

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

                If LoginCheck.LoginCheck(user_id, "MOA00100") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00101.aspx")
                    Response.End()
                End If

                Dim AddDate As String = Request.QueryString("AddDate")
                Dim sql As String = ""

                If IsPostBack Then
                    yy = txtSelYear.Text
                Else
                    yy = Request.QueryString("year")
                    txtSelYear.Text = yy
                End If

                If IsDate(Request.QueryString("Date")) Then
                    dDate = CDate(Request.QueryString("Date"))
                    yy = Year(dDate)
                    If (AddDate.Equals("1")) Then
                        sql = "insert into P_12(fod_date,fod_year) values(@DATE,@YEAR)"
                        SqlDataSource1.InsertCommand = sql
                        SqlDataSource1.InsertParameters.Clear()
                        SqlDataSource1.InsertParameters.Add("DATE", dDate)
                        SqlDataSource1.InsertParameters.Add("YEAR", yy)
                        Try
                            SqlDataSource1.Insert()
                        Catch ex As Exception
                        End Try
                    ElseIf (AddDate.Equals("0")) Then
                        sql = "Delete from P_12 where fod_date=@DATE and fod_year=@YEAR"
                        SqlDataSource1.DeleteCommand = sql
                        SqlDataSource1.DeleteParameters.Clear()
                        SqlDataSource1.DeleteParameters.Add("DATE", dDate)
                        SqlDataSource1.DeleteParameters.Add("YEAR", yy)
                        SqlDataSource1.Delete()
                    End If
                ElseIf yy = 0 Then
                    dDate = Now()
                    yy = Year(dDate)
                End If
                sql = "select fod_date,fod_year from P_12 where fod_year=" & yy
                SqlDataSource1.SelectCommand = sql
                dv = CType(SqlDataSource1.Select(DataSourceSelectArguments.Empty), DataView)

                Dim mm As Int16 = 1
                Do While mm <= 12
                    LoadCalendar(Me.FindControl("PlaceHolder" & mm), CDate(mm & "-01-" & yy))
                    mm += 1
                Loop

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub btnImgLeft_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgLeft.Click

        '上年度
        Server.Transfer("MOA00101.aspx?year=" & txtSelYear.Text - 1)

    End Sub

    Protected Sub btnImgRight_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgRight.Click

        '下年度
        Server.Transfer("MOA00101.aspx?year=" & txtSelYear.Text + 1)

    End Sub

    Protected Sub btnImgBack_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgBack.Click

        Server.Transfer("MOA00100.aspx")

    End Sub
End Class
