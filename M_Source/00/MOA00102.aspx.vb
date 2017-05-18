Imports System.Data.sqlclient
Partial Class Source_00_MOA00102
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Public Shared OrgChange As String      '判斷組織是否變更
    Dim chk As New C_CheckFun
    Dim len As New Integer
    Dim connstr, user_id, org_uid, PARENT_ORG_UID As String
    Dim ParentFlag As String = ""
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        PARENT_ORG_UID = Session("PARENT_ORG_UID")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            Dim connstr As String = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷是否有下一級單位
            db.Open()
            Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                ParentFlag = "Y"
            End If
            db.Close()

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA00102") <> "" And ParentFlag = "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00102.aspx")
                Response.End()
            End If

            '先設定起始日期
            Dim dt As Date = Now()
            If (Agent_SDate.Text = "") Then
                Agent_SDate.Text = dt.Date
            End If

            If (Agent_EDate.Text = "") Then
                Agent_EDate.Text = dt.AddDays(30).Date
            End If

            '找出上一級單位以下全部單位
            Dim Org_Down As New C_Public

            '找出登入者的一級單位
            Dim strParentOrg As String = ""
            strParentOrg = Org_Down.getUporg(org_uid, 1)

            '判斷登入者權限
            If Session("Role") = "1" Then
                SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                SqlDataSource3.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(Dep.SelectedValue) & ") ORDER BY [emp_chinese_name]"
                SqlDataSource4.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(Dep.SelectedValue) & ") ORDER BY [emp_chinese_name]"
            ElseIf Session("Role") = "2" Then
                SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") ORDER BY ORG_NAME"
                SqlDataSource3.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(Dep.SelectedValue) & ") ORDER BY [emp_chinese_name]"
                SqlDataSource4.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(Dep.SelectedValue) & ") ORDER BY [emp_chinese_name]"
            Else
                SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID ='" & Session("ORG_UID") & "' OR ORG_UID ='" & Session("PARENT_ORG_UID") & "' ORDER BY ORG_NAME"

                If ParentFlag = "Y" Then
                    SqlDataSource3.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(Dep.SelectedValue) & ") ORDER BY [emp_chinese_name]"
                    SqlDataSource4.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(Dep.SelectedValue) & ") ORDER BY [emp_chinese_name]"
                Else
                    SqlDataSource3.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(Dep.SelectedValue) & ") ORDER BY [emp_chinese_name]"
                    SqlDataSource4.SelectCommand = "SELECT employee_id,emp_chinese_name,ORG_UID FROM EMPLOYEE WHERE EMPLOYEE.employee_id='" & Session("user_id") & "' ORDER BY [emp_chinese_name]"
                End If
            End If

            If IsPostBack = False Then
                '登入馬上查詢
                ImgSearch_Click(Nothing, Nothing)

            End If

        End If

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView1.RowDeleted
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)

        If (sort) Then
            GridView1.Sort("DepName,Agent_SDate", SortDirection.Ascending)
            Return
        End If

        Dim Org_Down As New C_Public
        Dim sql As String

        '找出登入者的一級單位
        Dim strParentOrg As String = ""
        strParentOrg = Org_Down.getUporg(org_uid, 1)

        sql = "SELECT DepName,Agent_Name1,Agent_Name2,convert(nvarchar,Agent_SDate,111) Agent_SDate,convert(nvarchar,Agent_EDate,111) Agent_EDate,Agent_Num"
        sql += " FROM AGENT"
        sql += " where 1=1 "

        If Session("Role") = "2" Then
            sql += " and AGENT.Dep IN (" & Org_Down.getchildorg(strParentOrg) & ")"
        ElseIf Session("Role") = "3" Then
            sql += " and (AGENT.Dep ='" & Session("ORG_UID") & "' or AGENT.Dep ='" & PARENT_ORG_UID & "')"
            sql += " and ((AGENT.Agent1 = '" + Session("user_id") + "'" '代理人
            sql += " OR AGENT.Agent1 = '" + Session("user_id") + "')" '被代理人
            sql += " OR (AGENT.Agent2 = '" + Session("user_id") + "'" '代理人
            sql += " OR AGENT.Agent2 = '" + Session("user_id") + "'))" '被代理人
        End If

        If (Dep.Text <> "") Then sql += " and AGENT.Dep = '" + Dep.Text + "'" '代理單位
        If (Agent1.Text <> "") Then sql += " and AGENT.Agent1 = '" + Agent1.Text + "'" '代理人
        If (Agent2.Text <> "") Then sql += " and AGENT.Agent2 = '" + Agent2.Text + "'" '被代理人

        '代理時間
        If (Agent_SDate.Text <> "" And Agent_EDate.Text <> "") Then
            sql += " and AGENT.Agent_EDate>=CONVERT(datetime,'" & Agent_SDate.Text & " 00:00:00') and AGENT.Agent_SDate<=CONVERT(datetime,'" & Agent_EDate.Text & " 23:59:59')"
        ElseIf (Agent_SDate.Text <> "") Then
            sql += " and AGENT.Agent_EDate>=CONVERT(datetime,'" & Agent_SDate.Text & " 00:00:00') and AGENT.Agent_SDate<=CONVERT(datetime,'" & Agent_SDate.Text & " 23:59:59')"
        ElseIf (Agent_EDate.Text <> "") Then
            sql += " and AGENT.Agent_EDate>=CONVERT(datetime,'" & Agent_EDate.Text & " 00:00:00') and AGENT.Agent_SDate<=CONVERT(datetime,'" & Agent_EDate.Text & " 23:59:59')"
        End If

        SqlDataSource2.SelectCommand = sql
        OrgChange = ""
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        ErrMsg.Text = ""
        Search(True)
    End Sub

    Protected Sub Agent1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Agent1.SelectedIndexChanged
        OrgChange = "" '判斷組織是否變更
    End Sub

    Protected Sub Agent2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Agent2.SelectedIndexChanged
        OrgChange = "" '判斷組織是否變更
    End Sub

    Protected Sub Dep_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Dep.SelectedIndexChanged
        OrgChange = "1" '判斷組織是否變更
        If (Dep.Text = "") Then
            Agent1.Items.Clear()
            Agent2.Items.Clear()
        End If
    End Sub

    Protected Sub Agent1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Agent1.PreRender
        If OrgChange = "1" Then
            Dim tLItm As New ListItem("", "")
            Agent1.Items.Insert(0, tLItm)
            If Agent1.Items.Count > 1 Then
                Agent1.Items(1).Selected = False
            End If
        End If
    End Sub

    Protected Sub Agent2_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Agent2.PreRender
        If OrgChange = "1" Then
            Dim tLItm As New ListItem("", "")
            Agent2.Items.Insert(0, tLItm)
            If Agent2.Items.Count > 1 Then
                Agent2.Items(1).Selected = False
            End If
        End If
    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete
        If Not IsPostBack Then
            Dim tLItm As New ListItem("", "")
            Agent1.Items.Insert(0, tLItm)
            If Agent1.Items.Count > 1 Then
                Agent1.Items(1).Selected = False
            End If
            Agent2.Items.Insert(0, tLItm)
            If Agent2.Items.Count > 1 Then
                Agent2.Items(1).Selected = False
            End If
        End If
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        Dim tLItm As New ListItem("", "")
        ErrMsg.Text = ""
        Dep.Text = ""
        Agent1.Text = ""
        Agent2.Text = ""
        'Agent_SDate.Text = ""
        'Agent_EDate.Text = ""
        Agent1.Items.Clear()
        Agent2.Items.Clear()
        Agent1.Items.Insert(0, tLItm)
        Agent2.Items.Insert(0, tLItm)
    End Sub

    Protected Sub ImgInsert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgInsert.Click
        ErrMsg.Text = ""
        Try

            Dim AgentFlag As New C_Public

            '同一個人不可當自己的代理人
            If Agent1.SelectedValue = Agent2.SelectedValue Then
                ErrMsg.Text = "同一個人不得互相代理"
            Else

                '主官管不可以設定代理其他人
                If UCase(Agent1.SelectedValue) <> UCase(user_id) And UCase(Agent2.SelectedValue) <> UCase(user_id) And ParentFlag = "Y" Then
                    ErrMsg.Text = "代理人或被代理人其中之一須為本人"
                Else

                    If AgentFlag.DouBleAgent(Agent2.SelectedValue, Agent_SDate.Text, Agent_EDate.Text) = "" And AgentFlag.DouBleAgent(Agent1.SelectedValue, Agent_SDate.Text, Agent_EDate.Text) = "" Then

                        chk.CheckDataLen(Dep.Text, 10, "新增時：<單位>", True)
                        chk.CheckDataLen(Dep.SelectedItem.Text, 50, "新增時：<單位名稱>", True)
                        chk.CheckDataLen(Agent1.Text, 10, "新增時：<代理人帳號>", True)
                        chk.CheckDataLen(Agent2.Text, 10, "新增時：<被代理人帳號>", True)
                        chk.CheckDataLen(Agent_SDate.Text, 10, "新增時：<代理時間(起)>", True)
                        chk.CheckDataLen(Agent_EDate.Text, 10, "新增時：<代理時間(迄)>", True)
                        chk.CheckDataLen(Agent1.SelectedItem.Text, 20, "新增時：<代理人名稱>", True)
                        chk.CheckDataLen(Agent2.SelectedItem.Text, 20, "新增時：<被代理人名稱>", True)

                        Dim sql As String
                        sql = "insert into AGENT(Dep,DepName,Agent1,Agent2,Agent_SDate,Agent_EDate,Agent_Name1,Agent_Name2,Ins_Userid)"
                        sql += " values(@Dep,@DepName,@Agent1,@Agent2,@Agent_SDate,@Agent_EDate,@Agent_Name1,@Agent_Name2,@Ins_Userid)"
                        SqlDataSource2.InsertCommand = sql
                        SqlDataSource2.InsertParameters.Clear()
                        SqlDataSource2.InsertParameters.Add("Dep", Dep.Text)
                        SqlDataSource2.InsertParameters.Add("DepName", Dep.SelectedItem.Text)
                        SqlDataSource2.InsertParameters.Add("Agent1", Agent1.SelectedValue)
                        SqlDataSource2.InsertParameters.Add("Agent2", Agent2.SelectedValue)
                        SqlDataSource2.InsertParameters.Add("Agent_SDate", Agent_SDate.Text)
                        SqlDataSource2.InsertParameters.Add("Agent_EDate", Agent_EDate.Text)
                        SqlDataSource2.InsertParameters.Add("Agent_Name1", Agent1.SelectedItem.Text)
                        SqlDataSource2.InsertParameters.Add("Agent_Name2", Agent2.SelectedItem.Text)
                        SqlDataSource2.InsertParameters.Add("Ins_Userid", user_id)
                        SqlDataSource2.Insert()

                        ImgClearAll_Click(Nothing, Nothing)

                        Search(False)
                    Else
                        ErrMsg.Text = "代理人或被代理人這段時間已被代理，不可重複。"
                    End If


                End If

            End If

        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try
    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "90px"
        Div_grid.Style("left") = "90px"

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "90px"
        Div_grid2.Style("left") = "190px"

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Agent_SDate.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        Agent_EDate.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

End Class
