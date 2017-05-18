
Partial Class Source_00_MOA00061
    Inherits System.Web.UI.Page

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

                If LoginCheck.LoginCheck(user_id, "MOA00061") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00061.aspx")
                    Response.End()
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub
    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete
        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")
            DropDownList1.Items.Insert(0, tLItm)
            DropDownList1.Items(1).Selected = False
        End If
    End Sub

    Protected Sub searchbtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles searchbtn.Click

        Dim strOrd As String

        strOrd = " ORDER BY SYSTEMOBJ.object_name,ADMINGROUP.ORG_NAME"

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Public Function SQLALL(ByVal StepSel, ByVal PerSel)

        '整合SQL搜尋字串
        Dim strsql As String = "SELECT SYSTEMOBJUSE.object_num, SYSTEMOBJ.object_name, EMPLOYEE.emp_chinese_name, EMPLOYEE.employee_id, ADMINGROUP.ORG_NAME FROM SYSTEMOBJ INNER JOIN SYSTEMOBJUSE ON SYSTEMOBJ.object_uid = SYSTEMOBJUSE.object_uid INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID "

        '關卡搜尋
        If DropDownList1.SelectedIndex <> 0 Then
            strsql += " WHERE SYSTEMOBJ.object_Uid = '" & StepSel & "'"
        End If

        '人員搜尋
        If PerName.Text <> "" Then
            strsql += " AND EMPLOYEE.emp_chinese_name LIKE '%" & PerSel & "%'"
        End If

        SQLALL = strsql

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY SYSTEMOBJ.object_name,ADMINGROUP.ORG_NAME"

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序條件
        Dim strOrd As String

        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Server.Transfer("MOA00062.aspx")

    End Sub
End Class
