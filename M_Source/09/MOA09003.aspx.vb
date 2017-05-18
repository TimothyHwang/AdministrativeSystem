Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_09_MOA09003
    Inherits System.Web.UI.Page
    Dim sql_function As New C_SQLFUN
    Dim connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            SettingDepartments() '設定部門名稱搜尋下拉選單
            Dim strStartDate As String = DateAdd(DateInterval.Day, -3, DateTime.Now).ToString("yyyy/MM/dd")
            txtMeetingDateStart.Text = strStartDate
            Dim strEndDate As String = DateAdd(DateInterval.Day, 3, DateTime.Now).ToString("yyyy/MM/dd")
            txtMeetingDateEnd.Text = strEndDate
            gvP_09.DataSource = GetP_09GroupByDept(strStartDate, strEndDate, "")
            gvP_09.DataBind()
        End If
    End Sub

    '設定部門名稱搜尋下拉選單
    Private Sub SettingDepartments()
        ddlDepartment.Items.Add(New ListItem("全部", ""))
        db.Open()
        Dim sqlComm As New SqlCommand("SELECT DISTINCT(Sponsor) FROM P_09", db)
        Dim reader As SqlDataReader = sqlComm.ExecuteReader()
        While reader.Read()
            ddlDepartment.Items.Add(New ListItem(reader("Sponsor").ToString(), reader("Sponsor").ToString()))
        End While
        reader.Close()
        sqlComm.Dispose()
        db.Close()
    End Sub

    ''' <summary>
    ''' 取得門禁會議管制申請單(P_09)部門統計資料
    ''' </summary>
    ''' <param name="startDate">搜尋條件開會起始日期</param>
    ''' <param name="endDate">搜尋條件開會結束日期</param>
    ''' <param name="strDept">部門名稱搜尋字串</param>
    ''' <returns>P_09部門統計資料DataTable</returns>
    ''' <remarks></remarks>
    Private Function GetP_09GroupByDept(ByVal startDate As String, ByVal endDate As String, ByVal strDept As String) As DataTable
        db.Open()
        Dim strSQL As String = "SELECT Sponsor, COUNT(Sponsor) AS NumberOfMeetings, SUM(EnteringPeopleNumber) AS NumberOfPeople FROM P_09 WHERE Status = 1"
        If Not String.IsNullOrEmpty(startDate) Then
            strSQL = strSQL + " AND MeetingDate >= '" + startDate + "'"
        End If
        If Not String.IsNullOrEmpty(endDate) Then
            strSQL = strSQL + " AND MeetingDate <= '" + endDate + "'"
        End If
        If Not String.IsNullOrEmpty(strDept) Then
            strSQL = strSQL + " AND Sponsor = '" + strDept + "' "
        End If
        strSQL = strSQL + " GROUP BY Sponsor"
        Dim sqlComm As New SqlCommand(strSQL, db)
        Dim ds = New DataSet()
        Dim da = New SqlDataAdapter(sqlComm)
        da.Fill(ds)
        db.Close()
        Return ds.Tables(0)
    End Function

    Protected Sub ImgSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        gvP_09.DataSource = GetP_09GroupByDept(txtMeetingDateStart.Text.Trim(), txtMeetingDateEnd.Text.Trim(), ddlDepartment.SelectedValue)
        gvP_09.DataBind()
    End Sub

    Protected Sub gvP_09_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvP_09.RowCommand
        Dim currentCommand As String = e.CommandName
        Select Case currentCommand
            Case "Detail"
                Server.Transfer("MOA09004.aspx?Sponsor=" & e.CommandArgument.ToString() & "&SDate=" & txtMeetingDateStart.Text & "&EDate=" & txtMeetingDateEnd.Text)
        End Select
    End Sub

    Protected Sub gvP_09_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvP_09.PageIndexChanging
        gvP_09.DataSource = GetP_09GroupByDept(txtMeetingDateStart.Text.Trim(), txtMeetingDateEnd.Text.Trim(), ddlDepartment.SelectedValue)
        gvP_09.PageIndex = e.NewPageIndex
        gvP_09.DataBind()
    End Sub

    Protected Sub Calendar1_SelectionChanged(sender As Object, e As System.EventArgs) Handles Calendar1.SelectionChanged
        txtMeetingDateStart.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose1_Click(sender As Object, e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar2_SelectionChanged(sender As Object, e As System.EventArgs) Handles Calendar2.SelectionChanged
        txtMeetingDateEnd.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False
    End Sub

    Protected Sub btnClose2_Click(sender As Object, e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub

    Protected Sub ImgDate1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "70px"
    End Sub

    Protected Sub ImgDate2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "130px"
    End Sub
End Class
