'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Created at 2010/08/26
'   Description : First Version for Development
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Imports WebUtilities.Pages

Partial Class Source_01_MOA01011
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

            If LoginCheck.LoginCheck(user_id, "MOA01011") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA01011.aspx")
                Response.End()
            End If

        End If

        If Not Me.IsPostBack Then

            Select Case Session.Item("Role")

                Case "1"
                    sdsOrganization.SelectCommand = sdsOrganization.SelectCommand.Replace("(1 = 2)", "(1 = 1)")

                Case Nothing

                Case Else

                    If Me.IsAuthorizedAdministrator(MOA01008.PublicServantSalariesAdministrator) Then
                        sdsOrganization.SelectCommand = sdsOrganization.SelectCommand.Replace("(1 = 2)", "(1 = 1)")
                    End If

            End Select

            sdsOrganization.SelectCommand += " Order by ORG_NAME"

        End If

        Me.ImageButtons_ClientScriptsPreparing()

    End Sub

    Private Sub ImageButtons_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_hour_limit.ClientID) _
               .Append("numeric_validation: [{ message: '加班總時數限制請輸入數字' }] } ") _
               .Append("]});")

        ibSearch.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

        scripts.Append("var error_message = '';")

        scripts.Append("function TextBoxes_Validation(container) { ") _
               .Append("var table = document.getElementById(container);") _
               .Append("if (table) { RecursiveControlsToValid(table); }}")

        scripts.Append("function RecursiveControlsToValid(control) { ") _
               .Append("if (typeof(control.tagName) != 'string') return; ") _
               .Append("if (control.tagName != 'INPUT' && control.tagName != 'TEXTAREA' && control.tagName != 'SELECT') {") _
               .Append("for (var i = 0; i < control.childNodes.length; ++i) { RecursiveControlsToValid(control.childNodes[i]); }") _
               .Append("return; } else { var acc_id = control.parentNode.parentNode.cells[0].innerText; switch (control.tagName) { ") _
               .Append("case 'INPUT': if (control.type == 'text') { if (control.value.Trim() == '') error_message += '- ' + acc_id + ' ") _
               .Append("加班總時數限制未填寫 !\n'; else { if ((/^[0-9]{1,}(\.[0-9]+)?$/.test(control.value.Trim()))) { if ((") _
               .Append("control.value < 20) || (control.value > 70)) error_message += '- ' + acc_id + ' ") _
               .Append("加班總時數限制請填入 20 - 70 之間的數字 !\n'; } else error_message += '- ' + ") _
               .Append("acc_id + ' 加班總時數限制請填寫數字 !\n'; }} break; } acc_id = null; }}")

        ClientScript.RegisterClientScriptBlock(Me.GetType(), "TextBoxes_Validator", scripts.ToString(), True)
        scripts.Remove(0, scripts.Length)

        scripts.Append("javascript:error_message = '';") _
               .AppendFormat("TextBoxes_Validation('{0}');", gvDataRecords.ClientID) _
               .Append("if (error_message != '') alert(error_message);") _
               .Append("return (error_message == '');")

        ibApply.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

    End Sub

    Private Function IsAuthorizedAdministrator(ByRef group_uid As String) As Boolean

        Dim command As New SqlCommand : With command
            .Connection = New SqlConnection(ConfigurationManager.ConnectionStrings("connectionString").ConnectionString)
            .CommandType = CommandType.Text
        End With

        scripts.Append("Select Count(employee_id) From ROLEGROUPITEM Where ") _
               .AppendFormat("Group_Uid = '{0}' ", group_uid) _
               .Append("And employee_id = @employee_id;")

        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)

        command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = user_id

        Try
            command.Connection.Open()

            Return CType(command.ExecuteScalar(), Boolean)

        Catch ex As Exception

            scripts.Append(ex.Message.Replace("'"c, """"c).Replace(Environment.NewLine, "\n")) _
                   .Insert(0, "alert('") _
                   .Append("');")

            ClientScript.RegisterStartupScript(Me.GetType(), "AuthorizedAdministrator", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

            Return False

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

            command.Dispose()
            command = Nothing

        End Try

    End Function

    Private Sub DataRecords_DataRebinding()

        Me.sdsDataRecords.SelectParameters.Clear()
        Me.SqlDataSourceControllers_Init(sdsDataRecords, EventArgs.Empty)

        Dim where_conditions As String = String.Empty

        If Not String.IsNullOrEmpty(Me.ViewState(ddl_user_picker.ID)) Then

            where_conditions += " And E.emp_chinese_name Like (@emp_chinese_name + '%')"
            sdsDataRecords.SelectParameters.Add(New Parameter("emp_chinese_name", DbType.String, Me.ViewState(ddl_user_picker.ID)))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState(tb_employee_id.ID)) Then

            where_conditions += " And E.employee_id = @employee_id"
            sdsDataRecords.SelectParameters.Add(New Parameter("employee_id", DbType.String, Me.ViewState(tb_employee_id.ID)))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState(tb_hour_limit.ID)) Then

            where_conditions += " And P.HourLimit = @HourLimit"
            sdsDataRecords.SelectParameters.Add(New Parameter("HourLimit", DbType.Int32, Me.ViewState(tb_hour_limit.ID)))

        End If

        sdsDataRecords.SelectCommand = sdsDataRecords.SelectCommand.Replace _
                                                                    ( _
                                                                          " And (1 = 1)", _
                                                                            where_conditions _
                                                                    )

    End Sub

    Protected Sub DropDownListControllers_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_org_picker.DataBound, ddl_user_picker.DataBound

        DirectCast(sender, DropDownList).Items.Insert(0, New ListItem("請選擇", String.Empty))

    End Sub

    Protected Sub ddl_org_picker_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_org_picker.SelectedIndexChanged
        ddl_user_picker.Items.Clear()
    End Sub

    Protected Sub SqlDataSourceControllers_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles sdsUsersList.Init, sdsDataRecords.Init

        Dim sdsController As SqlDataSource = DirectCast(sender, SqlDataSource)

        If IsNothing(sdsController.SelectParameters.Item("AD_TITLE")) Then
            sdsController.SelectParameters.Add(New Parameter("AD_TITLE", DbType.String))
            sdsController.SelectParameters.Item("AD_TITLE").Size = 50
        End If

        sdsController.SelectParameters("AD_TITLE").DefaultValue = MOA01009.Special_AD_Title
        sdsController = Nothing

    End Sub

    Protected Sub sdsDataRecords_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles sdsDataRecords.Selecting

        Dim param_organization As String = "''"

        If String.IsNullOrEmpty(Me.ViewState(ddl_org_picker.ID)) Then

            Select Case Session.Item("Role")

                Case "1"
                    param_organization = "ORG_UID"

                Case Nothing

                Case Else

                    If Me.IsAuthorizedAdministrator(MOA01008.PublicServantSalariesAdministrator) Then
                        param_organization = "ORG_UID"
                    End If

            End Select

        Else
            param_organization = String.Format("'{0}'", Me.ViewState(ddl_org_picker.ID).ToString())
        End If

        e.Command.CommandText = e.Command.CommandText.Replace _
                                                      ( _
                                                             "ORG_UID In ('')", _
                                                              String.Format _
                                                                     ( _
                                                                           "ORG_UID In ({0})", _
                                                                            param_organization _
                                                                     ) _
                                                      )

    End Sub

    Protected Sub ibApply_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibApply.Click

        If gvDataRecords.Rows.Count.Equals(0) Then
            Return
        End If

        Dim command As New SqlCommand : With command
            .CommandType = CommandType.Text
            .Connection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("ConnectionString").ConnectionString)
        End With

        scripts.Append("If ((Select Count(*) From P_0103 Where employee_id = @employee_id) = 0) ") _
               .Append("Insert Into P_0103(employee_id, HourLimit) Values(@employee_id, @HourLimit); ") _
               .Append("Else Update P_0103 Set HourLimit = @HourLimit Where employee_id = @employee_id;")

        command.CommandText = scripts.ToString()

        scripts.Remove(0, scripts.Length)

        command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = DBNull.Value
        command.Parameters.Add(New SqlParameter("HourLimit", SqlDbType.Int)).Value = DBNull.Value

        Try
            command.Connection.Open()
            command.Transaction = command.Connection.BeginTransaction(IsolationLevel.ReadCommitted)

            For Each data_row As GridViewRow In gvDataRecords.Rows

                command.Parameters("employee_id").Value = data_row.Cells(1).Text.Trim()
                command.Parameters("HourLimit").Value = DirectCast(data_row.Cells(2).FindControl("gv_tb_hour_limit"), TextBox).Text.Trim()

                command.ExecuteNonQuery()

            Next

            command.Transaction.Commit()
            scripts.Append("加班資料異動成功")

            Me.DataRecords_DataRebinding()

        Catch ex As Exception

            If TypeOf command.Transaction Is SqlTransaction Then
                command.Transaction.Rollback()
            End If

            If (TypeOf ex Is SqlException) Then

                scripts.AppendFormat("{0}: ", CType(ex, SqlException).Number) _
                       .Append(ex.Message.Replace("'"c, """"c).Replace(Environment.NewLine, String.Empty))

            Else
                scripts.Append(ex.Message.Replace("'"c, """"c).Replace(Environment.NewLine, String.Empty))
            End If

        Finally

            If TypeOf command.Transaction Is SqlTransaction Then
                command.Transaction.Dispose()
                command.Transaction = Nothing
            End If

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

            If Not IsNothing(command) Then
                command.Dispose()
                command = Nothing
            End If

            If Not scripts.Length.Equals(0) Then

                scripts.Insert(0, "alert('").Append(" !\n');")

                ClientScript.RegisterStartupScript(Me.GetType(), "Processed_Message", scripts.ToString(), True)
                scripts.Remove(0, scripts.Length)

            End If

        End Try

    End Sub

    Protected Sub ibSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibSearch.Click

        Dim tbControl As TextBox = Nothing

        For Each element As Control In plConditionsContainer.Controls

            If TypeOf element Is TextBox Then

                tbControl = DirectCast(element, TextBox)
                tbControl.Text = tbControl.Text.Trim()

                Me.ViewState.Remove(element.ID)

                If Not String.IsNullOrEmpty(tbControl.Text) Then
                    Me.ViewState.Item(element.ID) = tbControl.Text
                End If

            End If

        Next

        tbControl = Nothing

        Me.ViewState.Remove(ddl_org_picker.ID)

        If Not String.IsNullOrEmpty(ddl_org_picker.SelectedValue) Then
            Me.ViewState(ddl_org_picker.ID) = ddl_org_picker.SelectedValue
        End If

        Me.ViewState.Remove(ddl_user_picker.ID)

        If Not String.IsNullOrEmpty(ddl_user_picker.SelectedValue) Then
            Me.ViewState(ddl_user_picker.ID) = ddl_user_picker.SelectedItem.Text.Trim()
        End If

        Me.DataRecords_DataRebinding()

    End Sub

End Class