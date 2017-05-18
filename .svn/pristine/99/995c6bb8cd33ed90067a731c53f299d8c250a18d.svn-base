Imports System.Data.SqlClient
Partial Class Source_00_MOA00104
    Inherits System.Web.UI.Page

    Public OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid As String

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

                If LoginCheck.LoginCheck(user_id, "MOA00104") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00104.aspx")
                    Response.End()
                End If

                If IsPostBack = False Then

                    '先設定起始日期
                    Dim dt As Date = Now()
                    If (Sdate.Text = "") Then
                        Sdate.Text = dt.AddDays(-7).Date
                    End If

                    If (Edate.Text = "") Then
                        Edate.Text = dt.Date
                    End If

                    '找出上一級單位以下全部單位
                    Dim Org_Down As New C_Public

                    '找出登入者的一級單位
                    Dim strParentOrg As String = ""
                    strParentOrg = Org_Down.getUporg(org_uid, 1)

                    '判斷登入者權限
                    If Session("Role") = "1" Then
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = ""
                    ElseIf Session("Role") = "2" Then
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") ORDER BY ORG_NAME"
                    Else
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID ='" & Session("ORG_UID") & "' ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE employee_id ='" & Session("User_id") & "'"
                    End If

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇
            If Session("Role") = "1" Then

                OrgSel.Items.Insert(0, tLItm)
                If OrgSel.Items.Count > 1 Then
                    OrgSel.Items(1).Selected = False
                End If

            End If

            '人員加請選擇
            UserSel.Items.Insert(0, tLItm)
            If UserSel.Items.Count > 1 Then
                UserSel.Items(1).Selected = False
            End If

            '重新整理GridView
            Dim strOrd As String

            strOrd = " ORDER BY flowctl.appdate DESC"

            SqlDataSource2.SelectCommand = SQLALL(Sdate.Text, Edate.Text) & strOrd

            OrgChange = ""

        End If
    End Sub

    Protected Sub UserSel_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles UserSel.PreRender

        If OrgChange = "1" Then

            Dim tLItm As New ListItem("請選擇", "")

            '人員加請選擇
            UserSel.Items.Insert(0, tLItm)
            If UserSel.Items.Count > 1 Then
                UserSel.Items(1).Selected = False
            End If

        End If

    End Sub

    Protected Sub OrgSel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OrgSel.SelectedIndexChanged

        '判斷組織是否變更
        OrgChange = "1"

        '清空User重新讀取
        UserSel.Items.Clear()


    End Sub

    Public Function SQLALL(ByVal SDate, ByVal EDate)

        '呼叫查詢單位函數
        Dim Org_Down As New C_Public

        '找出登入者的一級單位
        Dim strParentOrg As String = ""
        strParentOrg = Org_Down.getUporg(org_uid, 1)

        '整合SQL搜尋字串
        Dim strsql, strDate, strOrg As String

        '判斷登入者權限
        If Session("Role") = "1" Then

            strsql = "SELECT flowctl.flowsn,flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid,flowctl.emp_chinese_name, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent,V_EformShow.PWNAME,V_EformShow.PANAME FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') "

            '申請日期搜尋
            strDate = " AND (flowctl.appdate between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"

            '找出一級單位所屬單位
            strOrg = ""

        ElseIf Session("Role") = "2" Then

            strsql = "SELECT flowctl.flowsn,flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid,flowctl.emp_chinese_name, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent,V_EformShow.PWNAME,V_EformShow.PANAME FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') "

            '申請日期搜尋
            strDate = " AND (flowctl.appdate between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"

            '找出一級單位所屬單位
            strOrg = " AND deptcode IN (" & Org_Down.getchildorg(strParentOrg) & ") "

        Else
            strsql = ""
            strDate = ""
            strOrg = ""
        End If

        SQLALL = strsql & strDate & strOrg

    End Function

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Dim strOrd As String

        strOrd = " ORDER BY flowctl.appdate DESC"

        SqlDataSource2.SelectCommand = SQLALL(Sdate.Text, Edate.Text) & strOrd

        OrgChange = ""

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY flowctl.appdate DESC"

        SqlDataSource2.SelectCommand = SQLALL(Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '隱藏flowsn
            e.Row.Cells(1).Visible = False

        End If

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub ImgAssign_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgAssign.Click

        Try

            Dim struser_id As String = ""

            struser_id = UserSel.SelectedValue

            If struser_id = "" Then

                ErrName.Visible = True

            Else

                ErrName.Visible = False

                '表單重新分派
                Dim FC As New C_FlowSend.C_FlowSend

                Dim i As Integer = 0
                For i = 0 To GridView1.Rows.Count - 1
                    If CType(GridView1.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then

                        Dim strflowsn As String = ""

                        strflowsn = GridView1.Rows(i).Cells(1).Text

                        Dim Val_P As String = ""

                        Dim SendVal As String = struser_id & "," & strflowsn

                        Val_P = FC.F_Assign(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

                    End If
                Next

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('重新分派完成!!!');")
                Response.Write(" </script>")

                '一登入馬上查詢
                ImgSearch_Click(Nothing, Nothing)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "60px"

        Calendar1.SelectedDate = Sdate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "230px"

        Calendar2.SelectedDate = Edate.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Sdate.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        Edate.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub
End Class
