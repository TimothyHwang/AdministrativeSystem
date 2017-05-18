Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_01_MOA01013
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
    Dim connstr, user_id, org_uid As String
    Dim str_ORG_UID As String = ""
    Dim n_table As New System.Data.DataTable
    Dim sAction As String
    Dim sP_NUM As String

    ''' <summary>
    ''' 檢查是否超過管理員作業時間15:30
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function CheckTime() As Boolean
        If CType(ApplyDate.Text, Date).ToShortDateString = Now.ToShortDateString Then
            CheckTime = CType(CType(ApplyDate.Text, Date).ToShortDateString & " " & Now.Hour & ":" & Now.Minute & ":00", DateTime) >= CType(Now.Date & " 15:30:00", DateTime)
        Else
            CheckTime = CType(ApplyDate.Text, Date).ToShortDateString < Now.Date
        End If
        'Dim boolReturn As Boolean = False
        'If CType(ApplyDate.Text, Date) = CType(Now.ToShortDateString(), Date) Then
        '    boolReturn = True
        '    If Now.Hour = 15 Then
        '        boolReturn = True
        '        If Now.Minute >= 30 Then
        '            boolReturn = True
        '        Else
        '            boolReturn = False
        '        End If
        '    ElseIf Now.Hour > 15 Then
        '        boolReturn = True
        '    ElseIf Now.Hour < 15 Then
        '        boolReturn = False
        '    End If
        'ElseIf CType(ApplyDate.Text, Date) > CType(Now.ToShortDateString(), Date) Then
        '    boolReturn = True
        'End If        
        'CheckTime = boolReturn
    End Function
    ''' <summary>
    ''' 檢查頁面資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function CheckForm() As Boolean
        Dim boolReturn As Boolean = True
        If Lab_ORG_NAME_1.Text.Length = 0 Then boolReturn = False
        If Lab_emp_chinese_name.Text.Length = 0 Then boolReturn = False
        If Lab_title_name_1.Text.Length = 0 Then boolReturn = False
        If Lab_ORG_NAME_2.Text.Length = 0 Then boolReturn = False
        If Lab_title_name_2.Text.Length = 0 Then boolReturn = False
        If DrDown_emp_chinese_name.Text.Length = 0 Then boolReturn = False
        If ApplyDate.Text.Length = 0 Then boolReturn = False
        If ddlStartHour.Text.Length = 0 Then boolReturn = False
        If ddlEndHour.Text.Length = 0 Then boolReturn = False
        If txtLocation.Text.Length = 0 Then boolReturn = False
        If TXT_nREASON.Text.Length = 0 Then boolReturn = False
        CheckForm = boolReturn
    End Function

    ''' <summary>
    ''' 檢查輸入資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function CheckData() As Boolean
        Dim boolReturn As Boolean = False
        Dim strSql As String = ""
        strSql = "SELECT * FROM P_0107 WHERE "
        strSql = strSql & "PAIDNO='" & DrDown_emp_chinese_name.SelectedValue & "'"
        strSql = strSql & " AND STARTTIME='" & ApplyDate.Text & " 00:00:00'"
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            boolReturn = True
        Else
            boolReturn = False
        End If
        DC.Dispose()
        'strSql = strSql & ddlStartHour
        'strSql = strSql & ddlEndHour
        CheckData = boolReturn
    End Function

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

        Dim stmt As String = ""
        Dim p As Integer = 0
        Dim K As Integer = 0
        Dim tool As New C_Public
        do_sql.G_errmsg = ""
        do_sql.G_user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        sAction = Request("Action")
        sP_NUM = Request("P_NUM")

        'session被清空回首頁
        If user_id = "" And org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '新增連線
            connstr = do_sql.G_conn_string

            If IsPostBack Then
                Exit Sub
            End If

            ApplyDate.Text = Now.ToString("yyyy/MM/dd")
            If do_sql.select_urname(do_sql.G_user_id) = False Then
                Lab_ORG_NAME_1.Text = ""
                Exit Sub
            End If
            If do_sql.G_usr_table.Rows.Count > 0 Then
                Lab_ORG_NAME_1.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
                Lab_emp_chinese_name.Text = do_sql.G_usr_table.Rows(0).Item("emp_chinese_name").ToString.Trim
                Lab_title_name_1.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
                str_ORG_UID = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
            End If

            If str_ORG_UID <> "" Then
                Dim Org_Down As New C_Public

                'stmt = "select * from Employee where ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") order by emp_chinese_name"
                stmt = "select * from Employee where ORG_UID IN (" + tool.SelectWholeTreeORG_UID(user_id, 5) + ") order by emp_chinese_name"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    n_table = do_sql.G_table
                    p = 0
                    K = 0
                    DrDown_emp_chinese_name.Items.Clear()

                    For Each dr As DataRow In n_table.Rows
                        DrDown_emp_chinese_name.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_emp_chinese_name.Items(p).Value = Trim(dr("employee_id").ToString)
                        If UCase(do_sql.G_user_id) = UCase(Trim(dr("employee_id").ToString)) Then
                            K = p
                        End If
                        p += 1
                    Next
                    If p > 0 Then
                        DrDown_emp_chinese_name.SelectedIndex = K
                        Lab_ORG_NAME_2.Text = Org_Down.GetOrgNameByIDNo(DrDown_emp_chinese_name.SelectedValue)
                        Call DrDown_emp_chinese_name_SelectedIndexChanged(sender, e)
                    End If
                End If
            End If
            ddlStartHour.Text = "17"
            ddlEndHour.Text = "18"
            If sAction = "2" AndAlso sP_NUM.Length > 0 Then
                stmt = "SELECT * FROM P_0107 WHERE P_NUM='" & sP_NUM & "'"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    btnModify.Visible = False
                    btnCancel.Visible = False
                    Exit Sub
                End If
                btnModify.Visible = True
                btnCancel.Visible = True
                btnSubmit.Visible = False
                If do_sql.G_table.Rows.Count > 0 Then
                    n_table = do_sql.G_table
                    For Each dr As DataRow In n_table.Rows
                        ApplyDate.Text = CType(dr("STARTTIME"), Date).ToString("yyyy/MM/dd")
                        ddlStartHour.Text = dr("STHOUR")
                        ddlEndHour.Text = dr("ETHOUR")
                        txtLocation.Text = dr("LOCATION")
                        TXT_nREASON.Text = dr("REASON")
                        Exit For
                    Next
                End If
            End If
        End If
    End Sub

    Protected Sub DrDown_emp_chinese_name_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_emp_chinese_name.SelectedIndexChanged

        'str_app_id = ""

        If do_sql.select_urname(DrDown_emp_chinese_name.Items(DrDown_emp_chinese_name.SelectedIndex).Value) = False Then
            Lab_ORG_NAME_2.Text = ""
            Exit Sub
        Else
            Dim tool As New C_Public
            Lab_ORG_NAME_2.Text = tool.GetOrgNameByIDNo(DrDown_emp_chinese_name.SelectedValue)
            Lab_title_name_2.Text = tool.GetADTitleByIDNo(DrDown_emp_chinese_name.SelectedValue)
        End If

    End Sub

    Protected Sub btnSubmit_Click(sender As Object, e As System.EventArgs) Handles btnSubmit.Click
        Dim boolInsertFlag As Boolean = True
        If Not CheckForm() Then
            boolInsertFlag = False
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('頁面資料不得空白，請重新輸入');")
            Response.Write(" </script>")
            Exit Sub
        End If
        If CheckData() Then
            boolInsertFlag = False
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('同日期已有加班紀錄，請確認');")
            Response.Write(" </script>")
            Exit Sub
        End If
        If CheckTime() Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('已超過當日加班申請呈核作業時間，請自行補單上呈！');")
            Response.Write(" </script>")
            'MessageBox.Show("已超過當日加班申請呈核作業時間，請自行補單上呈！")
        End If
        If boolInsertFlag Then
            Dim strSql As String = ""
            strSql = "INSERT INTO P_0107 (PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PATITLE,PANAME,PAIDNO,STARTTIME,STHOUR,ETHOUR,LOCATION,REASON) VALUES ("
            strSql += "'" + Lab_ORG_NAME_1.Text + "'"
            strSql += ",'" + Lab_title_name_1.Text + "'"
            strSql += ",'" + Lab_emp_chinese_name.Text + "'"
            strSql += ",'" + do_sql.G_user_id + "'"
            strSql += ",'" + Lab_ORG_NAME_2.Text + "'"
            strSql += ",'" + Lab_title_name_2.Text + "'"
            strSql += ",'" + DrDown_emp_chinese_name.SelectedItem.Text + "'"
            strSql += ",'" + DrDown_emp_chinese_name.SelectedValue + "'"
            strSql += ",'" + ApplyDate.Text + "'"
            strSql += ",'" + ddlStartHour.Text + "'"
            strSql += ",'" + ddlEndHour.Text + "'"
            strSql += ",'" + txtLocation.Text + "'"
            strSql += ",'" + TXT_nREASON.Text + "'"

            strSql += ")"
            If do_sql.db_exec(strSql, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            'Response.Redirect("MOA01014.aspx")
            Response.Write(" <script language='javascript'>")
            Response.Write(" window.location='MOA01014.aspx';")
            Response.Write(" </script>")
        End If
    End Sub

    Protected Sub btnModify_Click(sender As Object, e As System.EventArgs) Handles btnModify.Click
        sAction = Request("Action")
        sP_NUM = Request("P_NUM")
        If sAction = "2" AndAlso sP_NUM.Length > 0 Then
        If Not CheckForm() Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('頁面資料不得空白，請重新輸入');")
            Response.Write(" </script>")
            Exit Sub
        End If
        If CheckTime() Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('已超過當日加班申請呈核作業時間，請自行補單上呈！');")
            Response.Write(" </script>")
            'MessageBox.Show("已超過當日加班申請呈核作業時間，請自行補單上呈！")
        End If
            Dim strSql As String = ""
            strSql = "UPDATE P_0107 SET "
            strSql += "STARTTIME='" + ApplyDate.Text + "'"
            strSql += ",STHOUR='" + ddlStartHour.Text + "'"
            strSql += ",ETHOUR='" + ddlEndHour.Text + "'"
            strSql += ",LOCATION='" + txtLocation.Text + "'"
            strSql += ",REASON='" + TXT_nREASON.Text + "'"

            strSql += " WHERE P_NUM='" & sP_NUM & "'"
            If do_sql.db_exec(strSql, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            'Response.Redirect("MOA01014.aspx")
            Response.Write(" <script language='javascript'>")
            Response.Write(" window.location='MOA01014.aspx';")
            Response.Write(" </script>")
        End If
    End Sub

    Protected Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        Response.Redirect("MOA01014.aspx")
    End Sub
End Class
