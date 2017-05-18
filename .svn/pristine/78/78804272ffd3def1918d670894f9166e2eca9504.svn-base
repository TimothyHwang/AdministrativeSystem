'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Created at 2010/08/17
'   Description : First Version for Development
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_04_MOA04022
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

            If LoginCheck.LoginCheck(user_id, "MOA04022") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04022.aspx")
                Response.End()
            End If

        End If

        If Not IsPostBack Then
            '先設定起始日期
            Dim dt As Date = Now()
            If (Sdate.Text = "") Then
                Sdate.Text = dt.AddDays(-14).Date
            End If

            If (Edate.Text = "") Then
                Edate.Text = dt.Date
            End If
        End If
        

        'Me.ImageButtonSearch_ClientScriptsPreparing()

    End Sub

    Private Sub ImageButtonSearch_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_eFormSn.ClientID) _
               .Append("numeric_validation: [{ message: '請輸入報修單單號' }] } ") _
               .Append("]});")

        ibtnSearch.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

    End Sub

    Private Sub DataRecords_DataRebinding()

        Me.sdsDataRecords.SelectParameters.Clear()

        Dim where_conditions As String = String.Empty

        If Not String.IsNullOrEmpty(Me.ViewState("sEFormSn")) Then

            where_conditions += " And EFORMSN Like ('%' +@EFORMSN + '%')"
            sdsDataRecords.SelectParameters.Add(New Parameter("EFORMSN", DbType.String, Me.ViewState("sEFormSn")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("sDateS")) And Not String.IsNullOrEmpty(Me.ViewState("sDateE")) Then

            where_conditions += " And (nAPPTIME between @DateS and @DateE)"
            sdsDataRecords.SelectParameters.Add(New Parameter("DateS", DbType.String, Me.ViewState("sDateS")))
            sdsDataRecords.SelectParameters.Add(New Parameter("DateE", DbType.String, Me.ViewState("sDateE")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("nAppStockStatus")) Then
            If Me.ViewState("nAppStockStatus") = "Null" Then
                where_conditions += " And nAppStockStatus Is null"
            Else
                where_conditions += " And nAppStockStatus = @nAppStockStatus"
                sdsDataRecords.SelectParameters.Add(New Parameter("nAppStockStatus", DbType.String, Me.ViewState("nAppStockStatus")))
            End If



        End If



        Me.sdsDataRecords.SelectCommand = sdsDataRecords.SelectCommand.Insert _
                                                         ( _
                                                               sdsDataRecords.SelectCommand.IndexOf(" Order by"), _
                                                               where_conditions _
                                                         )

    End Sub

    Protected Sub gvDataRecords_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvDataRecords.PageIndexChanging
        Me.DataRecords_DataRebinding()
    End Sub

    Protected Sub gvDataRecords_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvDataRecords.Sorted
        Me.DataRecords_DataRebinding()
    End Sub

    'Protected Sub ibtnApply_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnApply.Click

    '    Dim has_selection As Boolean = False

    '    For Each data_row As GridViewRow In gvDataRecords.Rows

    '        If DirectCast(data_row.Cells(0).FindControl("cbSelection"), CheckBox).Checked Then
    '            has_selection = True
    '            Exit For
    '        End If

    '    Next

    '    If has_selection Then

    '        scripts.Append("Update P_0414 Set UseCheck = '2', UseDate = GetDate() ") _
    '               .Append("Where shcode = @shcode;")

    '        Dim command As New SqlCommand : With command

    '            .CommandType = CommandType.Text
    '            .CommandText = scripts.ToString()
    '            .Connection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

    '            .Parameters.Add(New SqlParameter("shcode", SqlDbType.VarChar, 18)).Value = DBNull.Value

    '        End With

    '        scripts.Remove(0, scripts.Length)

    '        Try
    '            command.Connection.Open()
    '            command.Transaction = command.Connection.BeginTransaction(IsolationLevel.ReadCommitted)

    '        Catch ex As Exception

    '            scripts.Append(ex.Message.Replace("'"c, """"c).Replace(Environment.NewLine, "\n"))

    '            If command.Connection.State.Equals(ConnectionState.Open) Then
    '                command.Connection.Close()
    '            End If

    '            command.Dispose()
    '            command = Nothing

    '        End Try

    '        If Not IsNothing(command) Then

    '            Try

    '                For Each data_row As GridViewRow In gvDataRecords.Rows

    '                    If DirectCast(data_row.Cells(0).FindControl("cbSelection"), CheckBox).Checked Then
    '                        command.Parameters.Item("shcode").Value = data_row.Cells(1).Text.Trim()
    '                        command.ExecuteNonQuery()
    '                    End If

    '                Next

    '                command.Transaction.Commit()

    '            Catch ex As Exception

    '                command.Transaction.Rollback()
    '                scripts.Append(ex.Message.Replace("'"c, """"c).Replace(Environment.NewLine, "\n"))

    '            Finally

    '                If command.Connection.State.Equals(ConnectionState.Open) Then
    '                    command.Connection.Close()
    '                End If

    '                command.Dispose()
    '                command = Nothing

    '            End Try

    '        End If

    '    End If

    '    If scripts.Length.Equals(0) Then
    '        Me.DataRecords_DataRebinding()
    '    Else

    '        scripts.Insert(0, "alert('").Append("');")

    '        ClientScript.RegisterStartupScript(Me.GetType(), "ApplyProcess_Error", scripts.ToString(), True)
    '        scripts.Remove(0, scripts.Length)

    '    End If

    'End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click
        Dim sEFormSn, sDateS, sDateE, nAppStockStatus As String
        Dim EFORMSN As TextBox = Nothing
        Dim nAppStockDateS As TextBox = Nothing
        Dim nAppStockDateE As TextBox = Nothing
        Dim ddl_nAppStockStatus2 As DropDownList = Nothing
        For Each element As Control In plConditionsContainer.Controls
            If TypeOf element Is TextBox Then
                EFORMSN = DirectCast(tb_eFormSn, TextBox)
                nAppStockDateS = DirectCast(Sdate, TextBox)
                nAppStockDateE = DirectCast(Edate, TextBox)
                ddl_nAppStockStatus2 = DirectCast(ddl_nAppStockStatus, DropDownList)
                sEFormSn = EFORMSN.Text.Trim()
                sDateS = nAppStockDateS.Text.Trim()
                sDateE = nAppStockDateE.Text.Trim()
                nAppStockStatus = ddl_nAppStockStatus2.SelectedValue

                Me.ViewState.Clear()

                If Not String.IsNullOrEmpty(sEFormSn) Then
                    Me.ViewState.Add("sEFormSn", sEFormSn)
                End If
                If Not String.IsNullOrEmpty(sDateS) Then
                    sDateS = sDateS + " 00:00:00"
                    Me.ViewState.Add("sDateS", sDateS)
                End If
                If Not String.IsNullOrEmpty(sDateE) Then
                    sDateE = sDateE + " 23:59:59"
                    Me.ViewState.Add("sDateE", sDateE)
                End If
                If Not String.IsNullOrEmpty(nAppStockStatus) And nAppStockStatus <> "0" Then
                    If nAppStockStatus = "N" Then
                        nAppStockStatus = "Null"
                    End If
                    Me.ViewState.Add("nAppStockStatus", nAppStockStatus)
                End If
            End If

        Next
        Me.DataRecords_DataRebinding()
    End Sub
    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Try
            Div_grid.Visible = True
            Div_grid.Style("Top") = "55px"
            Div_grid.Style("left") = "375px"
            Calendar1.SelectedDate = Sdate.Text
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
        Try
            Div_grid2.Visible = True
            Div_grid2.Style("Top") = "55px"
            Div_grid2.Style("left") = "595px"
            Calendar2.SelectedDate = Edate.Text
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Try
            Sdate.Text = Calendar1.SelectedDate.Date
            Div_grid.Visible = False

        Catch ex As Exception

        End Try
    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        Try
            Edate.Text = Calendar2.SelectedDate.Date
            Div_grid2.Visible = False

        Catch ex As Exception

        End Try
    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub

    Protected Sub ibtclearall_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtclearall.Click
        Sdate.Text = ""
        Edate.Text = ""
        tb_eFormSn.Text = ""
        ddl_nAppStockStatus.SelectedIndex = "0"

    End Sub
End Class