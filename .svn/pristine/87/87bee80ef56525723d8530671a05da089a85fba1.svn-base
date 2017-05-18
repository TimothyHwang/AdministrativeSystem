Imports System.Data.SqlClient

Partial Class M_Source_10_MOA10004
    Inherits Page

    ReadOnly sql_function As New C_SQLFUN
    ReadOnly connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

#Region "Custom Function"
    ''**********************  以下為Custom Function  **************************
#End Region

#Region "Form Function"
    ''**********************  以下為Form Function  ****************************
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Page.Header.DataBind()
    End Sub

    Protected Sub gvManagerEmployee_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvManagerEmployee.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim lblManager As Label = CType(e.Row.FindControl("lblManager"), Label)
                Dim lblManagerID As Label = CType(e.Row.FindControl("lblManagerID"), Label)
                Dim ibnSelectEdit As ImageButton = CType(e.Row.FindControl("ibnSelectEdit"), ImageButton)
                Dim SqlDataSourceP_1001 As SqlDataSource = CType(e.Row.FindControl("SqlDataSourceP_1001"), SqlDataSource)
                Dim gvEmployee As GridView = CType(e.Row.FindControl("gvEmployee"), GridView)

                Dim strSql As String
                Dim dr As SqlDataReader
                strSql = "SELECT * FROM P_1001 WHERE MANAGER_ID='" & IIf(lblManagerID.Text.Length = 0, "", lblManagerID.Text) & "'"
                SqlDataSourceP_1001.SelectCommand = strSql
                SqlDataSourceP_1001.DataBind()
                gvEmployee.DataBind()

                db.Open()
                strSql = "SELECT * FROM P_10 A WHERE P_NUM='" & IIf(lblManagerID.Text.Length = 0, "", lblManagerID.Text) & "'"
                dr = New SqlCommand(strSql, db).ExecuteReader
                If (dr.Read) Then
                    'lblManagerID.Text = lblManager.Text
                    ibnSelectEdit.CommandArgument = lblManagerID.Text & "|" & e.Row.RowIndex
                    lblManager.Text = dr("NAME").ToString()
                End If

                db.Close()
        End Select
    End Sub


    Protected Sub gvManagerEmployee_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvManagerEmployee.RowCommand
        ''繫結按鈕事件
        Select Case e.CommandName
            Case "Add"
            Case "Select"
                Dim data() = Split(e.CommandArgument, "|")
                Dim pk As Integer = CInt(data(0))
                Dim strManagerName = CType(CType(sender, GridView).Rows(data(1)).FindControl("lblManager"), Label).Text
                'Response.Redirect("MOA10003.aspx?MID=" & pk & "&MName=" & strManagerName)
                Response.Redirect("MOA10008.aspx?MID=" & pk & "&MName=" & strManagerName)
        End Select
    End Sub

    Protected Sub imbClear_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imbClear.Click
        Dim strSql As String = "SELECT * FROM [P_10] A "
        SqlDataSourceP_10.SelectCommand = strSql
        SqlDataSourceP_10.DataBind()
        gvManagerEmployee.DataBind()

        ddlManager.Items(ddlManager.SelectedIndex).Selected = False
        ddlManager.Items(0).Selected = True
        ddlEmployee.Items(ddlEmployee.SelectedIndex).Selected = False
        ddlEmployee.Items(0).Selected = True
    End Sub

    Protected Sub imbQuery_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imbQuery.Click
        Dim strCondition As String = ""
        If ddlManager.SelectedValue <> 0 AndAlso ddlEmployee.SelectedValue <> 0 Then
            'strCondition += "AND B.MANAGER_ID='" & ddlManager.SelectedValue & "'"
            SqlDataSourceP_10.SelectCommand = "SELECT * FROM P_10 A LEFT JOIN P_1001 B ON A.P_NUM=B.MANAGER_ID WHERE A.P_NUM='" & ddlManager.SelectedValue & "' AND B.EMPLOYEE_ID='" & ddlEmployee.SelectedValue & "' ORDER BY A.[RANKID]"
        ElseIf ddlManager.SelectedValue <> 0 AndAlso ddlEmployee.SelectedValue = 0 Then
            SqlDataSourceP_10.SelectCommand = "SELECT * FROM P_10 A WHERE A.P_NUM='" & ddlManager.SelectedValue & "' ORDER BY [RANKID]"
        ElseIf ddlManager.SelectedValue = 0 AndAlso ddlEmployee.SelectedValue <> 0 Then
            'strCondition += "AND B.EMPLOYEE_ID='" & ddlEmployee.SelectedValue & "'"
            SqlDataSourceP_10.SelectCommand = "SELECT * FROM P_10 A RIGHT JOIN P_1001 B ON A.P_NUM=B.MANAGER_ID WHERE B.EMPLOYEE_ID='" & ddlEmployee.SelectedValue & "' ORDER BY A.[RANKID]"
        Else
            SqlDataSourceP_10.SelectCommand = "SELECT * FROM [P_10] A "
        End If

        SqlDataSourceP_10.DataBind()

    End Sub

    Protected Sub ddlManager_DataBound(sender As Object, e As System.EventArgs) Handles ddlManager.DataBound
        ddlManager.Items.Insert(0, New ListItem("----", "0"))
    End Sub

    Protected Sub ddlEmployee_DataBound(sender As Object, e As System.EventArgs) Handles ddlEmployee.DataBound
        ddlEmployee.Items.Insert(0, New ListItem("----", "0"))
    End Sub

    Protected Sub gvEmployee_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim strSql As String
                Dim dr As SqlDataReader
                db.Open()
                strSql = "SELECT EMP_CHINESE_NAME FROM EMPLOYEE WHERE EMPLOYEE_ID='" & IIf(e.Row.Cells(2).Text.Length = 0, 0, e.Row.Cells(2).Text) & "'"
                dr = New SqlCommand(strSql, db).ExecuteReader
                If (dr.Read) Then e.Row.Cells(2).Text = dr("EMP_CHINESE_NAME").ToString()
                db.Close()
        End Select
    End Sub

#End Region


End Class
