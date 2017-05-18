'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Finished at 2010/06/17
'   Description : First version of development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/18
'   Description : 修改 GridView 單筆 Record 修改的資料輸入模式 (另開新窗操作之)。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/21
'   Description : 修正查詢功能，該以傳遞參數的方式執行，提升安全性。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/24
'   Description : 修正執行查詢後，GridView 操作異動作業時，呈現錯誤問題。
'==================================================================================================================================================================================
Imports System.Data
Imports WebUtilities.Functions

Partial Class Source_04_MOA04010
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'Session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA04010") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04010.aspx")
                Response.End()
            End If

        End If

        If Not Page.IsPostBack Then

            If Not String.IsNullOrEmpty(Request.Form("query_element_no")) Then

                Me.ViewState("element_no") = Request.Form("query_element_no").Trim()
                tb_element_no.Text = Me.ViewState("element_no")

            End If

            If Not String.IsNullOrEmpty(Request.Form("query_element_code")) Then

                Me.ViewState("element_code") = Request.Form("query_element_code").Trim()
                tb_element_code.Text = Me.ViewState("element_code")

            End If

            GridView1_DataRebinding()

        End If

    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        tb_element_no.Text = tb_element_no.Text.Trim()
        tb_element_code.Text = tb_element_code.Text.Trim()

        Me.ViewState.Remove("element_no")
        Me.ViewState.Remove("element_code")

        If tb_element_no.Text <> "" Then
            Me.ViewState("element_no") = tb_element_no.Text.Trim()
        End If

        If tb_element_code.Text <> "" Then
            Me.ViewState("element_code") = tb_element_code.Text.Trim()
        End If

        GridView1_DataRebinding()

    End Sub

    Protected Sub ibtnAdd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim parameters As New SortedList()

        parameters.Add("Action", "MOA04011.aspx")

        If Not String.IsNullOrEmpty(Me.ViewState("element_no")) Then
            parameters.Add("query_element_no", Me.ViewState("element_no"))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("element_code")) Then
            parameters.Add("query_element_code", Me.ViewState("element_code"))
        End If

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()

    End Sub

    Private Sub GridView1_DataRebinding()

        SqlDataSource1.SelectParameters.Clear()

        Dim sql_statement As String = SqlDataSource1.SelectCommand

        sql_statement = sql_statement.Substring(0, sql_statement.ToUpper().IndexOf("ORDER BY"))
        sql_statement += "Where 1 = 1"

        If Not String.IsNullOrEmpty(Me.ViewState("element_no")) Then

            sql_statement += " And element_no = @element_no"
            SqlDataSource1.SelectParameters.Add(New Parameter("element_no", DbType.String, Me.ViewState("element_no")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("element_code")) Then

            sql_statement += " And element_code Like @element_code + '%'"
            SqlDataSource1.SelectParameters.Add(New Parameter("element_code", DbType.String, Me.ViewState("element_code")))

        End If

        sql_statement += " Order by element_no"

        SqlDataSource1.SelectCommand = sql_statement

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing

        Dim current_row As GridViewRow = GridView1.Rows(e.NewEditIndex)
        Dim parameters As New SortedList()

        parameters.Add("Action", "MOA04011.aspx")
        parameters.Add("element_no", current_row.Cells(0).Text.Trim())
        parameters.Add("element_code", current_row.Cells(1).Text.Trim())
        parameters.Add("bd_code", current_row.Cells(2).Text.Trim())
        parameters.Add("fl_code", current_row.Cells(3).Text.Trim())
        parameters.Add("rnum_code", current_row.Cells(4).Text.Trim())
        parameters.Add("wa_code", current_row.Cells(5).Text.Trim())
        parameters.Add("bg_code", current_row.Cells(6).Text.Trim())
        parameters.Add("it_code", current_row.Cells(7).Text.Trim())

        If Not String.IsNullOrEmpty(Me.ViewState("element_no")) Then
            parameters.Add("query_element_no", Me.ViewState("element_no"))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("element_code")) Then
            parameters.Add("query_element_code", Me.ViewState("element_code"))
        End If

        current_row = Nothing

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        GridView1_DataRebinding()
    End Sub

End Class