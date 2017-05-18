
Partial Class Source_00_MOA00054
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

                If LoginCheck.LoginCheck(user_id, "MOA00054") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00054.aspx")
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

        strOrd = " ORDER BY ROLEGROUP.Group_Name"

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Public Function SQLALL(ByVal GroupSel, ByVal PerSel)

        '整合SQL搜尋字串

        Dim strsql As String = "SELECT ROLEGROUP.Group_Name, EMPLOYEE.emp_chinese_name, ROLEGROUPITEM.Role_Num, ROLEGROUPITEM.Group_Uid, ROLEGROUPITEM.employee_id,ADMINGROUP.ORG_NAME FROM ROLEGROUPITEM INNER JOIN ROLEGROUP ON ROLEGROUPITEM.Group_Uid = ROLEGROUP.Group_Uid INNER JOIN EMPLOYEE ON ROLEGROUPITEM.employee_id = EMPLOYEE.employee_id INNER Join ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID"

        '查詢表單
        If GroupSel = "" Then
            strsql += " WHERE 1=1"
        Else
            strsql += " WHERE ROLEGROUPITEM.Group_Uid = '" & GroupSel & "'"
        End If

        '人員搜尋
        If PerSel <> "" Then
            strsql += " AND EMPLOYEE.emp_chinese_name LIKE '%" & PerSel & "%'"
        End If

        SQLALL = strsql

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY ROLEGROUP.Group_Name,ADMINGROUP.ORG_NAME"

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit

        '取消修改
        Dim strOrd As String

        strOrd = " ORDER BY ROLEGROUP.Group_Name"

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView1.RowDeleted

        '刪除
        Dim strOrd As String

        strOrd = " ORDER BY ROLEGROUP.Group_Name"

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd


    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing

        '修改
        Dim strOrd As String

        strOrd = " ORDER BY ROLEGROUP.Group_Name"

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource1.SelectCommand = SQLALL(DropDownList1.SelectedValue, PerName.Text) & strOrd

    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Server.Transfer("MOA00055.aspx")

    End Sub
End Class
