'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Created at 2010/08/23
'   Description : First Version for Development
'==================================================================================================================================================================================
Imports System.Data.Common

Partial Class Source_04_MOA04021
    Inherits System.Web.UI.Page
    Dim chk As New C_CheckFun
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

            If LoginCheck.LoginCheck(user_id, "MOA04021") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04021.aspx")
                Response.End()
            End If

        End If

        If Not Me.IsPostBack Then
            Me.ibtnSearch_Click(Me.ibtnSearch, Nothing)
        End If

        Me.ImageButtons_ClientScriptsPreparing()

    End Sub

    Protected Function ShowStocksStateText(ByRef stock_value As String) As String

        Select Case stock_value

            Case "0"
                Return "庫存"

            Case "1"
                Return "待出庫"

            Case "2"
                Return "出庫"

            Case Else
                Return String.Empty

        End Select

    End Function

    Private Sub ImageButtons_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_year.ClientID) _
               .Append("numeric_validation: [{ message: '倉儲年份請輸入數字' }] } ") _
               .Append("]});")

        ibtnSearch.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

    End Sub

    Protected Sub sdsDataRecords_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles sdsDataRecords.Selecting

        If e.Command.Parameters.Contains("@shcode_partial") Then

            e.Command.Parameters("@shcode_partial").Value = IIf(String.IsNullOrEmpty(Me.ViewState(tb_it_code.ID)), "%", Me.ViewState(tb_it_code.ID).ToString()) & _
                                                            IIf(String.IsNullOrEmpty(Me.ViewState(tb_year.ID)), "%", Me.ViewState(tb_year.ID).ToString())

        End If

        If e.Command.Parameters.Contains("@stocks_status") Then
            e.Command.Parameters("@stocks_status").Value = Me.ViewState(ddl_stocks_status.ID).ToString()
        End If

        For Each element As DbParameter In e.Command.Parameters

            If String.IsNullOrEmpty(element.Value) Then
                element.Value = DBNull.Value
            End If

        Next

    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        Me.tb_it_code.Text = tb_it_code.Text.Trim()
        Me.tb_year.Text = tb_year.Text.Trim()

        Me.ViewState(tb_it_code.ID) = tb_it_code.Text
        Me.ViewState(tb_year.ID) = tb_year.Text
        Me.ViewState(ddl_stocks_status.ID) = ddl_stocks_status.SelectedValue
        If (Not e Is Nothing And tb_year.Text.Trim().Length < 4 And tb_year.Text.Trim().Length > 0) Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('查詢輸入西元年份錯誤，請輸入4碼西元年份重新查詢！');")
            Response.Write("window.close();")
            Response.Write("</script>")
            Return
        End If
        Me.gvDataRecords.DataBind()

    End Sub
End Class