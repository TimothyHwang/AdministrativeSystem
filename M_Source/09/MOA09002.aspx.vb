﻿Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_09_MOA09002
    Inherits System.Web.UI.Page
    Dim sql_function As New C_SQLFUN
    Dim connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim strStartDate As String = DateAdd(DateInterval.Day, -3, DateTime.Now).ToString("yyyy/MM/dd")
            txtMeetingDateStart.Text = strStartDate
            Dim strEndDate As String = DateAdd(DateInterval.Day, 3, DateTime.Now).ToString("yyyy/MM/dd")
            txtMeetingDateEnd.Text = strEndDate
            gvP_09.DataSource = GetP_09ByDate(strStartDate, strEndDate)
            gvP_09.DataBind()
        End If
    End Sub

    ''' <summary>
    ''' 取得門禁會議管制申請單(P_09)資料
    ''' </summary>
    ''' <param name="startDate">搜尋條件開會起始日期</param>
    ''' <param name="endDate">搜尋條件開會結束日期</param>
    ''' <returns>P_09資料DataTable</returns>
    ''' <remarks></remarks>
    Private Function GetP_09ByDate(ByVal startDate As String, ByVal endDate As String) As DataTable
        db.Open()
        Dim strSQL As String = "SELECT P.*,A.ORG_NAME,E.emp_chinese_name FROM P_09 AS P "
        strSQL = strSQL + "JOIN EMPLOYEE AS E ON P.CreateBy = E.employee_id "
        strSQL = strSQL + "JOIN ADMINGROUP AS A ON E.ORG_UID = A.ORG_UID WHERE 1=1"
        If Not String.IsNullOrEmpty(startDate) Then
            strSQL = strSQL + " AND MeetingDate >= '" + startDate + "'"
        End If
        If Not String.IsNullOrEmpty(endDate) Then
            strSQL = strSQL + " AND MeetingDate <= '" + endDate + "'"
        End If
        strSQL = strSQL + " ORDER BY P.MeetingDate DESC, P.MeetingHour DESC"
        Dim sqlComm As New SqlCommand(strSQL, db)
        Dim ds = New DataSet()
        Dim da = New SqlDataAdapter(sqlComm)
        da.Fill(ds)
        db.Close()
        Return ds.Tables(0)
    End Function

    ''' <summary>
    ''' 轉換開會日期顯示文字
    ''' </summary>
    ''' <param name="meetingDate">開會日期</param>
    ''' <returns>開會日期顯示文字</returns>
    ''' <remarks></remarks>
    Public Function ShowMeetingDate(ByVal meetingDate As DateTime) As String
        Dim strChnDayOfWeek As String = ""
        Select Case meetingDate.DayOfWeek
            Case DayOfWeek.Sunday
                strChnDayOfWeek = "星期日"
            Case DayOfWeek.Monday
                strChnDayOfWeek = "星期一"
            Case DayOfWeek.Tuesday
                strChnDayOfWeek = "星期二"
            Case DayOfWeek.Wednesday
                strChnDayOfWeek = "星期三"
            Case DayOfWeek.Thursday
                strChnDayOfWeek = "星期四"
            Case DayOfWeek.Friday
                strChnDayOfWeek = "星期五"
            Case DayOfWeek.Saturday
                strChnDayOfWeek = "星期六"
        End Select
        Dim strDate As String = String.Format("{0}年{1}月{2}日({3})", meetingDate.Year, meetingDate.Month, meetingDate.Day, strChnDayOfWeek)
        Return strDate
    End Function

    ''' <summary>
    ''' 轉換進出營門顯示文字
    ''' </summary>
    ''' <param name="eformsn">表單資料ID</param>
    ''' <returns>營門顯示文字</returns>
    ''' <remarks></remarks>
    Public Function ShowEnteringGate(ByVal eformsn As String) As String
        db.Open()
        Dim strSQL As String = "SELECT * FROM P_0901 WHERE EFORMSN = @EFORMSN"
        Dim sqlComm As New SqlCommand(strSQL, db)
        sqlComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eformsn
        Dim ds = New DataSet()
        Dim da = New SqlDataAdapter(sqlComm)
        da.Fill(ds)
        db.Close()
        Dim strShowEnteringGate As String = ""
        Dim dt As DataTable = ds.Tables(0)
        For Each dr As DataRow In dt.Rows
            Select Case dr("GateNumber")
                Case 1
                    strShowEnteringGate = strShowEnteringGate + "博愛營區一號門<br />"
                Case 2
                    strShowEnteringGate = strShowEnteringGate + "博愛營區二號門<br />"
                Case 3
                    strShowEnteringGate = strShowEnteringGate + "博愛營區三號門<br />"
                Case 4
                    strShowEnteringGate = strShowEnteringGate + "博愛營區四號門<br />"
                Case 5
                    strShowEnteringGate = strShowEnteringGate + "採購中心<br />"
                Case 6
                    strShowEnteringGate = strShowEnteringGate + "博一大樓<br />"
                Case 7
                    strShowEnteringGate = strShowEnteringGate + "博二大樓<br />"
            End Select
        Next
        Return strShowEnteringGate
    End Function

    ''' <summary>
    ''' 轉換資料狀態顯示文字
    ''' </summary>
    ''' <param name="status">資料狀態</param>
    ''' <returns>資料狀態顯示文字</returns>
    ''' <remarks></remarks>
    Public Function ShowStatus(ByVal status As Integer) As String
        Select Case status
            Case 0
                Return "未登管"
            Case 1
                Return "已登管"
            Case 2
                Return "退件"
            Case 3
                Return "博愛營區四號門"
            Case 4
                Return "撤銷"
            Case Else
                Return "狀態編號異常"
        End Select
    End Function

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

    Protected Sub Calendar1_SelectionChanged(sender As Object, e As System.EventArgs) Handles Calendar1.SelectionChanged
        txtMeetingDateStart.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar2_SelectionChanged(sender As Object, e As System.EventArgs) Handles Calendar2.SelectionChanged
        txtMeetingDateEnd.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False
    End Sub

    Protected Sub btnClose1_Click(sender As Object, e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose2_Click(sender As Object, e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub

    Protected Sub ImgSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        gvP_09.DataSource = GetP_09ByDate(txtMeetingDateStart.Text, txtMeetingDateEnd.Text)
        gvP_09.DataBind()
    End Sub

    Protected Sub gvP_09_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvP_09.PageIndexChanging
        gvP_09.DataSource = GetP_09ByDate(txtMeetingDateStart.Text, txtMeetingDateEnd.Text)
        gvP_09.PageIndex = e.NewPageIndex
        gvP_09.DataBind()
    End Sub

    Protected Sub gvP_09_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvP_09.RowCommand
        Dim currentCommand As String = e.CommandName
        Select Case currentCommand
            Case "Detail"
                Dim doorAndMeetingControlID As String = GetEFormId("門禁會議管制申請單") '取得門禁會議管制申請單ID
                Dim strPath As String = "../00/MOA00020.aspx?x=MOA09001&y=" & doorAndMeetingControlID & "&Read_Only=1&EFORMSN=" & e.CommandArgument
                Response.Write(" <script language='javascript'>")
                Response.Write(" sPath = '" & strPath & "';")
                Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=700px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
                Response.Write(" showModalDialog(sPath,self,strFeatures);")
                Response.Write(" </script>")
        End Select
    End Sub

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        db.Open()
        Dim sqlcomm As New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db)
        GetEFormId = sqlcomm.ExecuteScalar()
        db.Close()
    End Function

    Protected Sub imbOutSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imbOutSearch.Click
        Try
            Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                conn.Open()
                ''取出門禁會議管制申請單且未成功匯出之資料
                Dim cmd As SqlCommand
                cmd = New SqlCommand("select * from P_09 P JOIN EMPLOYEE AS E ON P.CreateBy = E.employee_id JOIN ADMINGROUP AS B ON E.ORG_UID = B.ORG_UID right join V_EformFlow2 V on P.eformsn=V.eformsn where P.pendflag='E' and P.nCheckDT is null", conn)

                Dim DR As SqlDataReader
                DR = cmd.ExecuteReader
                Dim dt As New DataTable
                dt.Load(DR)
                If (dt.Rows.Count > 0) Then
                    ViewState("HasOutputData") = True
                    imbExport.Visible = True
                End If
                gvP_09.DataSourceID = ""
                gvP_09.DataSource = dt
                gvP_09.DataBind()

            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub

    Protected Sub imbExport_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imbExport.Click
        Try
            For Each GridViewRow In gvP_09.Rows
                Select Case GridViewRow.RowType
                    Case DataControlRowType.DataRow
                        ''insert外部資料庫
                        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                            conn.Open()
                            Dim tran As SqlTransaction
                            tran = conn.BeginTransaction
                            ''門禁會議管制申請單
                            Dim P_09AData As New P_09A(GridViewRow.Cells(0).Text)
                            P_09AData.Insert(tran, conn)
                            tran.Commit()
                        End Using
                        ''更新門禁欄位已匯出時間
                        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                            conn.Open()
                            Dim tran As SqlTransaction
                            tran = conn.BeginTransaction
                            Dim cmd As New SqlCommand("UPDATE P_09 SET nCheckDT=GetDate() WHERE EFORMSN='" & GridViewRow.Cells(0).Text & "'", conn, tran)
                            Dim iSuccess As Integer = cmd.ExecuteNonQuery
                            tran.Commit()
                        End Using
                End Select
            Next
            imbExport.Visible = False
            MessageBox.Show("資料匯出完成")
            Me.imbOutSearch_Click(sender, e)
        Catch sqlex As SqlException
            If (Not "-1".Equals(sqlex.Message.IndexOf("無法開啟至 SQL Server 的連接").ToString) Or Not "-1".Equals(sqlex.Message.IndexOf("登入失敗").ToString)) Then
                MessageBox.Show("門禁系統連線錯誤，請連繫管理員")
            Else
                MessageBox.Show(sqlex.Message)
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub

    Protected Sub gvP_09_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvP_09.RowCreated
        Try
            If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
                '隱藏eformsn,eformid
                e.Row.Cells(0).Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class