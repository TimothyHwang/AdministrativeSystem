Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_00_MOA00100
    Inherits System.Web.UI.Page

    Public iDIM      ' Days In Month
    Public iDOW      ' Day Of Week that month starts on
    Public iCurrent  ' Variable we use to hold current day of month as we write table
    Public iPosition ' Variable we use to hold current position in table
    Public iLooper   ' Variable used for misc. loops
    Dim sql As String = ""
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

    Sub SetCalendar(ByVal dDate As Date)
        iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
        iDOW = GetWeekdayMonthStartsOn(dDate)

        iCurrent = 1
        iPosition = iDOW

        Do While iCurrent <= iDIM
            If iPosition = 1 Or iPosition = 7 Then
                sql = "insert into P_12(fod_date,fod_year) values(@DATE,@YEAR)"
                SqlDataSource1.InsertCommand = sql
                SqlDataSource1.InsertParameters.Clear()
                SqlDataSource1.InsertParameters.Add("DATE", CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)))
                SqlDataSource1.InsertParameters.Add("YEAR", Sel_Y.Text)
                Try
                    SqlDataSource1.Insert()
                Catch ex As Exception
                    'MsgBox(CDate(Month(dDate) & "-" & iCurrent & "-" & Year(dDate)))
                    'MsgBox(ex.Message)
                    Exit Sub

                End Try
            End If

            iCurrent = iCurrent + 1
            If iPosition = 7 Then
                iPosition = 0
            End If
            iPosition = iPosition + 1
        Loop

    End Sub

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
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00100.aspx")
                    Response.End()
                End If

                Dim d As Date = Now()
                Dim yy As Integer = Year(d)

                If (Sel_Y.Items.Count = 0) Then
                    Sel_Y.Items.Insert(0, yy)
                    Sel_Y.Items.Insert(1, yy + 1)
                    Sel_Y.Items.Insert(2, yy + 2)
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Server.Transfer("MOA00101.aspx?year=" & Sel_Y.Text)
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        sql = "Delete from P_12 where fod_year=@YEAR"
        SqlDataSource1.DeleteCommand = sql
        SqlDataSource1.DeleteParameters.Clear()
        SqlDataSource1.DeleteParameters.Add("YEAR", TypeCode.Int32, Sel_Y.Text)
        SqlDataSource1.Delete()

        Dim mm As Int16 = 1
        Do While mm <= 12
            SetCalendar(CDate(mm & "-01-" & Sel_Y.Text))
            mm += 1
        Loop

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim strsqlAll As String = "SELECT fod_year,CONVERT(nvarchar,fod_date,111) as fod_date FROM P_12 WHERE fod_year='" & Sel_Y.SelectedValue & "'"

        '設定檔案路徑
        Dim path As String = "", filename As String = ""

        path = Server.MapPath("~/Drs/")
        filename = Date.Today.ToString("yyyyMMdd") & "_00100.csv"
        path = path & filename

        Dim colname As String = ""
        Dim data As String = ""

        colname = "年度,休假日期"

        colname = Left(colname, Len(colname) - 1)

        Dim sw As New System.IO.StreamWriter(path, False, Encoding.GetEncoding("big5"))
        sw.WriteLine(colname)

        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        db.Open()
        Dim dt2 As DataTable = New DataTable("")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(strsqlAll, db)
        da2.Fill(dt2)
        db.Close()

        For y As Integer = 0 To dt2.Rows.Count - 1

            data += dt2.Rows(y).Item("fod_year").ToString & "," & dt2.Rows(y).Item("fod_date").ToString & ","

            data = Left(data, Len(data) - 1)
            sw.WriteLine(data)
            data = ""

        Next

        sw.WriteLine("")
        sw.Close()

        '匯出檔案
        Response.Clear()
        Response.ContentType = "application/download"
        Response.AddHeader("content-disposition", "attachment;filename=" & filename)
        Response.WriteFile(path)
        Response.End()

    End Sub
End Class
