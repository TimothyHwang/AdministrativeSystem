'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Finished at 2010/06/28
'   Description : First Version for Development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/07/12
'   Description : 新增 "單月累積時數" 欄位的相關處理及驗證等相關功能。
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Imports WebUtilities.Pages

Partial Class Source_04_MOA01008
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder

    Dim P_0102 As DataTable = Nothing

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

            If LoginCheck.LoginCheck(user_id, "MOA01008") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA01008.aspx")
                Response.End()
            End If

        End If

    End Sub

    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        ImageButtonApply_ClientScriptsPreparing()
    End Sub

    Private Sub DropDownListOfMoneyKind_Initialization(ByRef ddlControl As DropDownList)

        For numeric As Integer = 1 To 10
            ddlControl.Items.Add("年功俸" & numeric.ToString())
            ddlControl.Items.Insert((numeric - 1), ("本俸" & numeric.ToString()))
        Next

        ddlControl.Items.Insert(0, String.Empty)

    End Sub

    Private Sub ImageButtonApply_ClientScriptsPreparing()

        scripts.Append("var error_message = '';")

        scripts.Append("function DropDownLists_Validation(container) { ") _
               .Append("var table = document.getElementById(container);") _
               .Append("if (table) { RecursiveControlsToValid(table); }}")

        scripts.Append("function RecursiveControlsToValid(control) { ") _
               .Append("if (typeof(control.tagName) != 'string') return; ") _
               .Append("if (control.tagName != 'INPUT' && control.tagName != 'TEXTAREA' && control.tagName != 'SELECT') {") _
               .Append("for (var i = 0; i < control.childNodes.length; ++i) { RecursiveControlsToValid(control.childNodes[i]); }") _
               .Append("return; } else { var acc_id = control.parentNode.parentNode.cells[0].innerText; switch (control.tagName) { ") _
               .Append("case 'SELECT': if (control.value.Trim() == '') { if (control.id.indexOf('gv_ddl_p_money1') != -1) ") _
               .Append("error_message += '- ' + acc_id + ' 的俸額未選擇 !\n'; else if (control.id.indexOf('gv_ddl_p_money2') != -1) ") _
               .Append("error_message += '- ' + acc_id + ' 的專業加給未選擇 !\n'; else if (control.id.indexOf('gv_ddl_p_money3') != -1) ") _
               .Append("error_message += '- ' + acc_id + ' 的主管加給未選擇 !\n'; } break; case 'INPUT': if (control.type == 'text') { ") _
               .Append("if (control.value.Trim() == '') error_message += '- ' + acc_id + ' 單月累積時數未填寫 !\n'; else { ") _
               .Append("if ((/^[0-9]{1,}(\.[0-9]+)?$/.test(control.value.Trim()))) { if ((control.value < 20) || (control.value > 70)) ") _
               .Append("error_message += '- ' + acc_id + ' 單月累積時數請填入 20 - 70 之間的數字 !\n'; } else error_message += '- ' + ") _
               .Append("acc_id + ' 單月累積時數請填寫數字 !\n'; }} break; } acc_id = null; }}")

        ClientScript.RegisterClientScriptBlock(Me.GetType(), "DropDownList_Validator", scripts.ToString(), True)
        scripts.Remove(0, scripts.Length)

        scripts.Append("javascript:error_message = '';") _
               .AppendFormat("DropDownLists_Validation('{0}');", GridView1.ClientID) _
               .Append("if (error_message != '') alert(error_message);") _
               .Append("return (error_message == '');")

        btnImgOK.Attributes("OnClick") = scripts.ToString()
        scripts.Remove(0, scripts.Length)

    End Sub

    Protected Sub SqlDataSource1_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles SqlDataSource1.Init
        SqlDataSource1.SelectParameters("AD_TITLE").DefaultValue = MOA01009.Special_AD_Title
    End Sub

    Protected Sub SqlDataSource1_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles SqlDataSource1.Selecting

        Dim c_public As C_Public = Nothing
        Dim param_organization As String = Nothing

        If Not IsNothing(Session("Role")) Then

            If Session.Item("Role").Equals("1") Then

                '系統管理員
                param_organization = "ORG_UID"

            Else

                c_public = New C_Public

                If Me.IsAuthorizedAdministrator(MOA01008.UnitSalariesAdministrator) Then

                    '單位薪俸管理員
                    param_organization = c_public.getchildorg(c_public.getUporg(org_uid, 1))

                Else

                    '一般使用者
                    param_organization = c_public.getchildorg(org_uid)

                    e.Command.CommandText = e.Command.CommandText.Replace _
                                                                  ( _
                                                                        "(1 = 1)", _
                                                                          String.Format _
                                                                                 ( _
                                                                                       "employee_id = '{0}'", _
                                                                                        user_id _
                                                                                 ) _
                                                                  )

                End If

            End If

        End If

        c_public = Nothing

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

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then

            Dim ddlControl As DropDownList = Nothing
            Dim selected_item As ListItem = Nothing

            If TypeOf P_0102 Is DataTable Then

                For index As Integer = 1 To 3

                    ddlControl = e.Row.Cells(index).FindControl(String.Format("gv_ddl_p_money{0}", index))

                    If Not ddlControl Is Nothing Then

                        ddlControl.DataSource = P_0102
                        ddlControl.DataValueField = "P_Money"
                        ddlControl.DataTextField = ddlControl.DataValueField

                        CType(ddlControl.DataSource, DataTable).DefaultView.RowFilter = String.Format("P_KIND = {0}", index)

                        ddlControl.DataBind()
                        ddlControl.Items.Insert(0, New ListItem(String.Empty, String.Empty))

                        selected_item = ddlControl.Items.FindByValue(CType(e.Row.DataItem, DataRowView).Row(String.Format("P_Money{0}", index)).ToString())

                        If Not selected_item Is Nothing Then
                            selected_item.Selected = True
                        End If

                    End If

                Next

                selected_item = Nothing
                ddlControl = Nothing

            End If

            ddlControl = e.Row.Cells(4).FindControl("gv_ddl_money_kind")

            If Not IsNothing(ddlControl) Then

                Me.DropDownListOfMoneyKind_Initialization(ddlControl)

                selected_item = ddlControl.Items.FindByText(DirectCast(e.Row.DataItem, DataRowView).Row("MoneyKind").ToString())

                If Not IsNothing(selected_item) Then
                    selected_item.Selected = True
                    selected_item = Nothing
                End If

                ddlControl = Nothing

            End If

        End If

    End Sub

    Protected Sub GridView1_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBinding

        Dim adapter As New SqlDataAdapter( _
                                             "Select P_Money, P_KIND From P_0102 Group by P_Money, P_KIND Order by P_KIND", _
                                             (New C_SQLFUN()).G_conn_string _
                                         )

        Try
            If P_0102 Is Nothing Then
                P_0102 = New DataTable()
            Else
                P_0102.Clear()
            End If

            adapter.Fill(P_0102)

        Catch ex As Exception

            If (TypeOf P_0102 Is DataTable) Then

                P_0102.Dispose()
                P_0102 = Nothing

            End If

            If (TypeOf ex Is SqlException) Then

                scripts.AppendFormat("{0}: ", CType(ex, SqlException).Number) _
                       .Append(ex.Message.Replace(Environment.NewLine, String.Empty))

            Else
                scripts.Append(ex.Message.Replace(Environment.NewLine, String.Empty))
            End If

            scripts.Insert(0, "alert('").Append(" !');")

            ClientScript.RegisterStartupScript(Me.GetType(), "ErrorMessage", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

        Finally

            If adapter.SelectCommand.Connection.State.Equals(ConnectionState.Open) Then
                adapter.SelectCommand.Connection.Close()
            End If

            adapter.Dispose()

        End Try

    End Sub

    Protected Sub btnImgOK_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgOK.Click

        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        command.CommandType = Data.CommandType.Text

        command.CommandText = "If ((Select Count(*) From P_0103 Where employee_id = @employee_id) = 0) " & _
                              "Insert Into P_0103(employee_id, P_Money1, P_Money2, P_Money3, HourLimit, OverTime, MoneyKind) " & _
                              "Values(@employee_id, @P_Money1, @P_Money2, @P_Money3, @HourLimit, " & _
                              "Ceiling((@P_Money1 + @P_Money2 + @P_Money3) / 240.0), @MoneyKind) Else " & _
                              "Update P_0103 Set P_Money1 = @P_Money1, P_Money2 = @P_Money2, P_Money3 = @P_Money3, " & _
                              "HourLimit = @HourLimit, OverTime = Ceiling((@P_Money1 + @P_Money2 + @P_Money3) / 240.0), " & _
                              "MoneyKind = @MoneyKind Where employee_id = @employee_id;"

        'Add for new version at 2010/08/25 : Shifted the column "HourLimit" to MOA01011.aspx
        command.CommandText = command.CommandText.Replace(" HourLimit,", String.Empty) _
                                                 .Replace(" @HourLimit,", String.Empty) _
                                                 .Replace(" HourLimit =", String.Empty)

        command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = DBNull.Value
        command.Parameters.Add(New SqlParameter("P_Money1", SqlDbType.Int)).Value = DBNull.Value
        command.Parameters.Add(New SqlParameter("P_Money2", SqlDbType.Int)).Value = DBNull.Value
        command.Parameters.Add(New SqlParameter("P_Money3", SqlDbType.Int)).Value = DBNull.Value
        command.Parameters.Add(New SqlParameter("MoneyKind", SqlDbType.NVarChar, 5)).Value = DBNull.Value

        command.Transaction = Nothing

        Try
            command.Connection.Open()
            command.Transaction = command.Connection.BeginTransaction(IsolationLevel.ReadCommitted)

            For Each row As GridViewRow In GridView1.Rows

                command.Parameters("employee_id").Value = DirectCast(row.Cells(0).FindControl("hf_employee_id"), HiddenField).Value.Trim()
                command.Parameters("P_Money1").Value = CType(row.Cells(1).FindControl("gv_ddl_p_money1"), DropDownList).SelectedValue.Trim()
                command.Parameters("P_Money2").Value = CType(row.Cells(2).FindControl("gv_ddl_p_money2"), DropDownList).SelectedValue.Trim()
                command.Parameters("P_Money3").Value = CType(row.Cells(3).FindControl("gv_ddl_p_money3"), DropDownList).SelectedValue.Trim()
                command.Parameters("MoneyKind").Value = DirectCast(row.Cells(4).FindControl("gv_ddl_money_kind"), DropDownList).SelectedItem.Text.Trim()

                If String.IsNullOrEmpty(command.Parameters("MoneyKind").Value) Then
                    command.Parameters("MoneyKind").Value = DBNull.Value
                End If

                command.ExecuteNonQuery()

            Next

            command.Transaction.Commit()

            GridView1.DataBind()

            scripts.Append("加班資料異動成功")

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

            If Not scripts.Length.Equals(0) Then

                scripts.Insert(0, "alert('").Append(" !\n');")

                ClientScript.RegisterStartupScript(Me.GetType(), "ErrorMessage", scripts.ToString(), True)
                scripts.Remove(0, scripts.Length)

            End If

        End Try

    End Sub

End Class