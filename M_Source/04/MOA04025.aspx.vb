﻿Imports System.Data
Imports WebUtilities.Functions
Imports System.Web
Partial Class M_Source_04_MOA04025
    Inherits System.Web.UI.Page
    Dim user_id, org_uid As String

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

            If LoginCheck.LoginCheck(user_id, "MOA04025") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04025.aspx")
                Response.End()
            End If

        End If

        If Not Page.IsPostBack Then

            If Not String.IsNullOrEmpty(Request.Form("query_it_code")) Then

                Me.ViewState("shcode") = Request.Form("query_it_code").Trim()
                Me.tb_it_code.Text = Me.ViewState.Item("it_code")

            End If

            If Not String.IsNullOrEmpty(Request.Form("query_it_name")) Then

                Me.ViewState("it_name") = Request.Form("query_it_name").Trim()
                Me.tb_itname.Text = Me.ViewState.Item("it_name")

            End If

            If Not String.IsNullOrEmpty(Request.Form("query_usecheck")) Then

                Me.ViewState("usecheck") = Request.Form("query_usecheck").Trim()
                Me.ddlUseCheck.SelectedValue = Me.ViewState.Item("usecheck")

            End If

            Me.DataRecords_DataRebinding()

        End If

    End Sub

    Protected Function ShowStocksStateText(ByRef stock_value As String) As String

        Select Case stock_value
            Case "0"
                Return "庫存"
            Case "2"
                Return "出庫"
            Case Else
                Return String.Empty
        End Select
    End Function

    Private Sub DataRecords_DataRebinding()

        Me.sdsDataRecords.SelectParameters.Clear()

        Dim where_conditions As String = String.Empty

        If Not String.IsNullOrEmpty(Me.ViewState("it_code")) Then

            where_conditions += " And shcode Like (@it_code + '%')"
            sdsDataRecords.SelectParameters.Add(New Parameter("it_code", DbType.String, Me.ViewState("it_code")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("it_name")) Then

            where_conditions += " And b.it_name Like ('%' + @it_name + '%')"
            sdsDataRecords.SelectParameters.Add(New Parameter("it_name", DbType.String, Me.ViewState("it_name")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("usecheck")) Then

            where_conditions += " And usecheck = (@usecheck)"
            sdsDataRecords.SelectParameters.Add(New Parameter("usecheck", DbType.String, Me.ViewState("usecheck")))

        End If

        Me.sdsDataRecords.SelectCommand = sdsDataRecords.SelectCommand.Insert _
                                                         ( _
                                                              sdsDataRecords.SelectCommand.IndexOf(" group by"), _
                                                               where_conditions _
                                                         )

    End Sub

    Protected Sub sdsDataRecords_Deleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) Handles sdsDataRecords.Deleting

        If e.Command.Parameters.Contains("@it_code") Then

            e.Command.Parameters("@it_code").Value = Left _
                                                    ( _
                                                         e.Command.Parameters("@it_code") _
                                                                  .Value _
                                                                  .ToString() _
                                                                  .PadLeft(18, " "c), _
                                                         14 _
                                                    ) _
                                                   .Trim()

        End If

    End Sub

    Protected Sub gvDataRecords_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvDataRecords.PageIndexChanging
        Me.DataRecords_DataRebinding()
    End Sub

    Protected Sub gvDataRecords_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvDataRecords.Sorted
        Me.DataRecords_DataRebinding()
    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        tb_it_code.Text = tb_it_code.Text.Trim()
        tb_itname.Text = tb_itname.Text.Trim()

        Me.ViewState.Remove("it_code")
        Me.ViewState.Remove("it_name")
        Me.ViewState.Remove("usecheck")

        If Not String.IsNullOrEmpty(tb_it_code.Text) Then
            Me.ViewState.Add("it_code", tb_it_code.Text)
        End If

        If Not String.IsNullOrEmpty(tb_itname.Text) Then
            Me.ViewState.Add("it_name", tb_itname.Text)
        End If

        If Not String.IsNullOrEmpty(ddlUseCheck.SelectedValue) Then
            If ddlUseCheck.SelectedValue <> "-1" Then
                Me.ViewState.Add("usecheck", ddlUseCheck.SelectedValue)
            End If
        End If

        Me.DataRecords_DataRebinding()

    End Sub

  
    Protected Sub ibtnAdd_Click(sender As Object, e As System.EventArgs) Handles ibtnAdd.Click
        Dim parameters As New SortedList : With parameters
            .Add("Action", "MOA04026.aspx")
        End With

        If Not IsNothing(Me.ViewState("shcode")) Then
            parameters.Add("query_shcode", Me.ViewState("shcode"))
        End If

        If Not IsNothing(Me.ViewState("it_name")) Then
            parameters.Add("query_it_name", Me.ViewState("it_name"))
        End If

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()
    End Sub
End Class