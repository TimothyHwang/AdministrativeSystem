Imports System.Data.SqlClient
Imports System.Data
Imports System.Collections.Generic

Partial Class M_Source_01_MOA01012
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
    Dim conn As New C_SQLFUN
    Dim connstr As String = conn.G_conn_string
    Dim db As New SqlConnection(connstr)
    Dim tool As New C_Public
    Dim user_id, org_uid, org_name As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load        
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        org_name = Session("ORG_NAME")

        If Not IsPostBack Then
            ViewState("user_id") = Session("user_id") '登入者ID
            ViewState("org_uid") = Session("ORG_UID") '登入者單位

            If ViewState("user_id") = "" Then 'session被清空回首頁
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            Else
                Dim dtYear = GetYears() '取得資料庫內已有資料之年度種類
                Dim thisYear As String = DateTime.Now.Year.ToString()
                If dtYear.Select("Year=" + thisYear).Length = 0 Then '若當中不包含今年
                    ddlYear.Items.Add(New ListItem(thisYear, thisYear)) '則在年度下拉選單中加入今年之選項
                Else
                    For Each drYear As DataRow In dtYear.Rows
                        ddlYear.Items.Add(New ListItem(drYear("Year"), drYear("Year")))
                    Next
                End If

                ViewState("leavelTwoOrg") = tool.getUporg(ViewState("org_uid"), 1) '找出登入者(單位管理員)的一級單位
                '檢查單位內是否有尚無今年慰勞假資料,但已有去年資料之人員,若有則將之去年之慰勞假天數複製至今年
                Dim checkResult As String = CheckNewYearHolidays(ViewState("leavelTwoOrg"))
                If Not checkResult.Equals("OK") Then
                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('" + checkResult + "');")
                    Response.Write(" </script>")
                    Exit Sub
                End If

                ''目前休假天數設定只開放給管理員、單位管理員及處單位人事員
                If Session("Role") = "1" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                ElseIf tool.CheckUserInRoleGroup(user_id, "單位管理員") Then
                    SqlDataSource1.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" + tool.getchildorg(ViewState("leavelTwoOrg")) + ") ORDER BY [ORG_NAME]"
                Else
                    'SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") ORDER BY ORG_NAME"
                    SqlDataSource1.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" + tool.SelectWholeTreeORG_UID(user_id, 5) + ") ORDER BY [ORG_NAME]"
                End If
                ''查詢字串設定完後先行繫結一次
                ddlOrgSel.DataBind()

                '取得單位內成員之慰勞假資料
                Dim dtEmp As DataTable = GetEmployees(ViewState("leavelTwoOrg"), thisYear,ddlOrgSel.SelectedValue)
                GridView1.DataSource = dtEmp
                GridView1.DataBind()
                '儲存目前之慰勞假TextBox欄位值,分頁時保存使用者輸入值,及最後全部資料Submit至資料庫時使用
                ViewState("CurrentValue") = SaveCurrentTable(dtEmp)
            End If

            
        End If
    End Sub

    Private Function GetYears() As DataTable
        Dim ds As New DataSet()
        db.Open()
        Dim adapter As New SqlDataAdapter()
        adapter.SelectCommand = New SqlCommand("SELECT DISTINCT Year FROM P_0106 ORDER BY Year DESC", db)
        adapter.Fill(ds)
        db.Close()
        Return ds.Tables(0)
    End Function

    ''' <summary>
    ''' 檢查今年慰勞假尚無資料但已有去年資料之人員,將去年之慰勞假天數複製至今年
    ''' </summary>
    ''' <param name="strDeptID">登入者單位ID</param>
    ''' <returns>執行成功或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function CheckNewYearHolidays(ByVal strDeptID As String) As String
        Dim msg As String = "OK"
        Dim ds As New DataSet()
        Dim trans As SqlTransaction = Nothing

        Try
            db.Open()
            trans = db.BeginTransaction() '取得單位內今年慰勞假尚無資料但已有去年資料之人員
            Dim strSQL As String = "SELECT E.employee_id, E.emp_chinese_name, P.Holidays, PL.Holidays AS LastYear FROM EMPLOYEE AS E " + _
                "LEFT JOIN P_0106 AS P ON E.employee_id = P.employee_id AND P.Year = '" + DateTime.Now.Year.ToString() + "' AND P.Status='LATEST' " + _
                "LEFT JOIN P_0106 AS PL ON E.employee_id = PL.employee_id AND PL.Year = '" + (DateTime.Now.Year - 1).ToString() + "' AND PL.Status='LATEST' " + _
                "WHERE E.ORG_UID IN (" + tool.getchildorg(strDeptID) + ") AND P.Holidays IS NULL AND PL.Holidays IS NOT NULL"
            Dim comm As New SqlCommand(strSQL, db, trans)
            Dim adapter As New SqlDataAdapter(comm)
            adapter.Fill(ds)
            Dim dt As DataTable = ds.Tables(0)
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows '複製去年慰勞假天數為該員之今年慰勞假天數
                    Dim strInsert As String = "INSERT INTO P_0106(employee_id,Holidays,Year,Status,Creator,CreateDate) " + _
                        "VALUES(@employee_id,@Holidays,@Year,'LATEST',@Creator,GETDATE())"
                    Dim insertComm As New SqlCommand(strInsert, db, trans)
                    insertComm.Parameters.Add("@employee_id", SqlDbType.VarChar, 10).Value = dr("employee_id")
                    insertComm.Parameters.Add("@Holidays", SqlDbType.Int).Value = dr("LastYear")
                    insertComm.Parameters.Add("@Year", SqlDbType.Int).Value = DateTime.Now.Year
                    insertComm.Parameters.Add("@Creator", SqlDbType.VarChar, 10).Value = ViewState("user_id").ToString()
                    insertComm.ExecuteNonQuery()
                Next
            End If

            trans.Commit()
            db.Close()
        Catch ex As Exception
            trans.Rollback()
            msg = ex.Message
        End Try
        
        Return msg
    End Function

    ''' <summary>
    ''' 取得目標年度及單位內之人員資料
    ''' </summary>
    ''' <param name="strDeptID">目標單位ID字串</param>
    ''' <param name="strYear">目標年度</param>
    ''' <returns>人員資料</returns>
    ''' <remarks></remarks>
    Private Function GetEmployees(ByVal strDeptID As String, ByVal strYear As String, ByVal strOrgID As String) As DataTable
        Dim ds As New DataSet()
        db.Open()
        Dim adapter As New SqlDataAdapter()
        Dim strSQL As String = ""
        ''目前休假天數設定只開放給管理員、單位管理員及處單位人事員
        If Session("Role") = "1" Then
            strSQL = "SELECT E.employee_id, E.emp_chinese_name, P.Holidays FROM EMPLOYEE AS E " + _
            "LEFT JOIN P_0106 AS P ON E.employee_id = P.employee_id AND P.Year = '" + strYear + "' " + _
            "AND P.Status='LATEST' WHERE E.LEAVE='Y' AND E.ORG_UID IN (" + IIf((Len(strOrgID) > 0 AndAlso strOrgID <> "ALL"), "'" + strOrgID + "'", tool.getchildorg(strDeptID)) + ")"
        ElseIf tool.CheckUserInRoleGroup(user_id, "單位管理員") Then
            strSQL = "SELECT E.employee_id, E.emp_chinese_name, P.Holidays FROM EMPLOYEE AS E " + _
            "LEFT JOIN P_0106 AS P ON E.employee_id = P.employee_id AND P.Year = '" + strYear + "' " + _
            "AND P.Status='LATEST' WHERE E.LEAVE='Y' AND E.ORG_UID IN (" + IIf((Len(strOrgID) > 0 AndAlso strOrgID <> "ALL"), "'" + strOrgID + "'", tool.getchildorg(strDeptID)) + ")"
        Else
            strSQL = "SELECT E.employee_id, E.emp_chinese_name, P.Holidays FROM EMPLOYEE AS E " + _
                        "LEFT JOIN P_0106 AS P ON E.employee_id = P.employee_id AND P.Year = '" + strYear + "' " + _
                        "AND P.Status='LATEST' WHERE E.LEAVE='Y' AND E.ORG_UID IN (" + IIf(Len(strOrgID) > 0, "'" + strOrgID + "'", "''") + ")"
        End If        
        adapter.SelectCommand = New SqlCommand(strSQL, db)
        adapter.Fill(ds)
        db.Close()
        Return ds.Tables(0)
    End Function

    Private Function SaveCurrentTable(ByVal dtSource As DataTable) As DataTable
        Dim dtDestination As New DataTable()
        dtDestination.Columns.Add("employee_id", GetType(String))
        dtDestination.Columns.Add("emp_chinese_name", GetType(String))
        dtDestination.Columns.Add("Holidays", GetType(String))

        For Each drSource As DataRow In dtSource.Rows
            Dim drDestination As DataRow = dtDestination.NewRow()
            drDestination("employee_id") = drSource("employee_id")
            drDestination("emp_chinese_name") = drSource("emp_chinese_name")
            drDestination("Holidays") = drSource("Holidays").ToString()
            dtDestination.Rows.Add(drDestination)
        Next
        Return dtDestination
    End Function

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim tb As TextBox = e.Row.Cells(1).FindControl("txtHolidays")
            Dim lb As Label = e.Row.Cells(1).FindControl("lbHolidays")
            If ddlYear.SelectedValue.Equals(DateTime.Now.Year.ToString()) Then
                tb.Visible = True '今年的資料才可做更改
                lb.Visible = False
            Else
                tb.Visible = False '其餘年分僅做資料顯示
                lb.Visible = True
            End If
        End If
    End Sub

    Protected Sub ImgSubmit_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgSubmit.Click
        Dim dateTimeNow As DateTime = DateTime.Now
        Dim trans As SqlTransaction = Nothing
        Dim dt As DataTable = ViewState("CurrentValue")
        ViewState("CurrentValue") = SaveCurrentPage(dt) '更新當前頁面之資料

        Try
            db.Open()
            trans = db.BeginTransaction()
            For Each dr As DataRow In dt.Rows
                Dim employeeID As String = dr("employee_id") '人員ID
                Dim holidays As String = dr("Holidays") '慰勞假天數
                Dim year As String = ddlYear.SelectedValue '年度
                Dim Result As String = "OK"

                Dim dtP0106 As DataTable = QueryP_0106(db, trans, employeeID, ddlYear.SelectedValue)
                If dtP0106.Rows.Count = 0 Then '若該人員尚無該年度之資料,新增該人員於目標年度之慰勞假資料
                    Result = InsertP_0106(db, trans, employeeID, holidays.Trim(), year, ViewState("user_id").ToString(), dateTimeNow)
                    If Not Result.Equals("OK") Then
                        Throw New Exception(Result)
                    End If
                ElseIf Not dtP0106.Rows(0).Item("Holidays").ToString().Equals(holidays.Trim()) Then
                    '已有資料,且慰勞假輸入值與原本的不相同,亦即資料庫須做資料變更處理
                    '將原本最新的資料之Status值由LATEST變更為EDITED,並記錄變更者於Editor欄位
                    Result = UpdateP_0106(db, trans, employeeID, year, ViewState("user_id").ToString(), dateTimeNow)
                    If Not Result.Equals("OK") Then
                        Throw New Exception(Result)
                    End If
                    '新增該人員於目標年度之經變更後最新資料,Status值為LATEST
                    Result = InsertP_0106(db, trans, employeeID, holidays, year, ViewState("user_id").ToString(), dateTimeNow)
                    If Not Result.Equals("OK") Then
                        Throw New Exception(Result)
                    End If
                End If
            Next
            trans.Commit()
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('資料儲存成功');")
            Response.Write(" </script>")
        Catch ex As Exception
            trans.Rollback()
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('" + ex.Message + "');")
            Response.Write(" </script>")
            Exit Sub
        Finally
            db.Close()
        End Try
    End Sub

    ''' <summary>
    ''' Query From P_0106 By employee_id and Year
    ''' </summary>
    ''' <param name="conn">SqlConnection</param>
    ''' <param name="employee_id">慰勞假資料人員ID</param>
    ''' <param name="Year">慰勞假年分</param>
    ''' <returns>目標人員於目標年分之慰勞假資料</returns>
    ''' <remarks></remarks>
    Private Function QueryP_0106(ByRef conn As SqlConnection, ByRef trans As SqlTransaction, ByVal employee_id As String, ByVal Year As Integer) As DataTable
        Dim ds As New DataSet()
        Dim comm As New SqlCommand("SELECT * FROM P_0106 WHERE employee_id = @employee_id AND Year = @Year AND Status = 'LATEST'", conn, trans)
        comm.Parameters.Add("@employee_id", SqlDbType.VarChar).Value = employee_id
        comm.Parameters.Add("@Year", SqlDbType.Int).Value = Year
        Dim da As New SqlDataAdapter(comm)
        da.Fill(ds)
        Return ds.Tables(0)
    End Function

    ''' <summary>
    ''' 新增資料至P_0106
    ''' </summary>
    ''' <param name="conn">SqlConnection</param>
    ''' <param name="trans">SqlTransaction</param>
    ''' <param name="employee_id">慰勞假資料人員ID</param>
    ''' <param name="Holidays">慰勞假天數</param>
    ''' <param name="Year">年分</param>
    ''' <param name="Creator">資料新增人員ID</param>
    ''' <param name="CreateDate">資料新增日期</param>
    ''' <returns>執行成功或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function InsertP_0106(ByRef conn As SqlConnection, ByRef trans As SqlTransaction, ByVal employee_id As String, _
                                    ByVal Holidays As String, ByVal Year As Integer, ByVal Creator As String, _
                                    ByVal CreateDate As DateTime) As String
        Dim result As String = "OK"
        Try
            Dim comm As New SqlCommand("INSERT INTO P_0106(employee_id,Holidays,Year,Status,Creator,CreateDate) " + _
                                   "VALUES(@employee_id,@Holidays,@Year,'LATEST',@Creator,@CreateDate)", conn, trans)
            comm.Parameters.Add("@employee_id", SqlDbType.VarChar).Value = employee_id
            Dim num As Integer = 0
            If Not String.IsNullOrEmpty(Holidays) Then
                num = Convert.ToInt32(Holidays)
            End If
            comm.Parameters.Add("@Holidays", SqlDbType.Int).Value = num
            comm.Parameters.Add("@Year", SqlDbType.Int).Value = Year
            comm.Parameters.Add("@Creator", SqlDbType.VarChar).Value = Creator
            comm.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = CreateDate
            comm.ExecuteNonQuery()
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 變更目標資料之Status欄位為EDITED狀態
    ''' </summary>
    ''' <param name="conn">SqlConnection</param>
    ''' <param name="trans">SqlTransaction</param>
    ''' <param name="employee_id">慰勞假資料人員ID</param>
    ''' <param name="Year">年分</param>
    ''' <param name="Editor">資料變更人員ID</param>
    ''' <param name="EditDate">資料變更日期</param>
    ''' <returns>執行成功或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function UpdateP_0106(ByRef conn As SqlConnection, ByRef trans As SqlTransaction, ByVal employee_id As String, _
                                    ByVal Year As Integer, ByVal Editor As String, ByVal EditDate As DateTime) As String
        Dim result As String = "OK"
        Try
            Dim comm As New SqlCommand("UPDATE P_0106 SET Status = 'EDITED',Editor = @Editor,EditDate = @EditDate " + _
                                   "WHERE employee_id = @employee_id AND Year = @Year AND Status = 'LATEST'", conn, trans)
            comm.Parameters.Add("@employee_id", SqlDbType.VarChar).Value = employee_id
            comm.Parameters.Add("@Year", SqlDbType.Int).Value = Year
            comm.Parameters.Add("@Editor", SqlDbType.VarChar).Value = Editor
            comm.Parameters.Add("@EditDate", SqlDbType.DateTime).Value = EditDate
            comm.ExecuteNonQuery()
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    Protected Sub GridView1_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Dim dt As DataTable = ViewState("CurrentValue")
        '儲存目前之慰勞假TextBox欄位值,分頁時保存使用者輸入值,及最後全部資料Submit至資料庫時使用
        ViewState("CurrentValue") = SaveCurrentPage(dt) '更新當前頁面之資料

        GridView1.PageIndex = e.NewPageIndex '分頁資料顯示
        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    ''' <summary>
    ''' 更新當前GridView分頁使用者輸入之慰勞假欄位值
    ''' </summary>
    ''' <param name="dt">當前資料DataTable</param>
    ''' <returns>更新後之資料</returns>
    ''' <remarks></remarks>
    Private Function SaveCurrentPage(ByVal dt As DataTable) As DataTable
        Dim i As Integer = 0
        For Each gvr As GridViewRow In GridView1.Rows
            Dim tb As TextBox = gvr.Cells(1).FindControl("txtHolidays")
            dt.Rows(GridView1.PageIndex * 10 + i)("Holidays") = tb.Text
            i = i + 1
        Next
        Return dt
    End Function

    Protected Sub ddlYear_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlYear.SelectedIndexChanged, ddlOrgSel.SelectedIndexChanged
        Dim dt As DataTable = GetEmployees(ViewState("leavelTwoOrg").ToString(), ddlYear.SelectedValue, ddlOrgSel.SelectedValue)
        GridView1.DataSource = dt '重新載入目標年度之資料
        GridView1.DataBind()
        ViewState("CurrentValue") = SaveCurrentTable(dt)

        If ddlYear.SelectedValue.Equals(DateTime.Now.Year.ToString()) Then
            ImgSubmit.Visible = True '今年的資料才可做更改
        Else
            ImgSubmit.Visible = False '非顯示今年之資料,則隱藏修改按鈕
        End If
    End Sub

    Protected Sub ddlOrgSel_DataBound(sender As Object, e As System.EventArgs) Handles ddlOrgSel.DataBound
        Dim ddl As DropDownList = CType(sender, DropDownList)
        ddl.Items.Insert(0, New ListItem("全部", "ALL"))
    End Sub
End Class
