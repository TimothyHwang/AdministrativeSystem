'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Create at 2010/06/17
'   Description : First version of development
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_04_MOA04011_BK
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'Session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA04010") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04011.aspx")
                Response.End()
            End If

        End If

        ImageButtonAdd_ClientScriptsPreparing()

    End Sub

    Private Sub ImageButtonAdd_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_bd_code.ClientID) _
               .Append("message: '請輸入預算代碼' }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_fl_code.ClientID) _
               .Append("message: '請輸入樓層代碼' }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_rnum_code.ClientID) _
               .Append("message: '請輸入房間地理位置代碼' }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_wa_code.ClientID) _
               .Append("message: '請輸入房間規格代碼' }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_bg_code.ClientID) _
               .Append("message: '請輸入建物代碼' }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_it_code.ClientID) _
               .Append("message: '請輸入物料分類代碼' }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_element_no.ClientID) _
               .Append("message: '請輸入設備編碼流水號' }, ") _
               .AppendFormat("{{ client_id: '{0}', ", liEmployee.ClientID) _
               .Append("message: '請選取人員帳號' } ") _
               .Append("]});")

        scripts.Append("javascript: alert(document.getElementById('liEmployee').value);")

        ibtnAdd.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        Dim sql_statement As String

        sql_statement = SqlDataSource1.SelectCommand
        sql_statement = sql_statement.Substring(0, sql_statement.ToUpper().IndexOf("AND 1 = 2")).Trim()

        sql_statement += " Where EMPLOYEE.emp_chinese_name like '" & tb_operator.Text.Trim() & "%'" & _
                         " Order by EMPLOYEE.emp_chinese_name"

        SqlDataSource1.SelectCommand = sql_statement

    End Sub

    Protected Sub ibtnAdd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        command.CommandType = Data.CommandType.Text

        command.CommandText = "Insert Into P_0405(bd_code, fl_code, rnum_code, wa_code, bg_code, it_code, " & _
                              "element_no, element_code, insertime, operator) Values(@bd_code, @fl_code, " & _
                              "@rnum_code, @wa_code, @bg_code, @it_code, @element_no, @element_code, " & _
                              "GETDATE(), @operator)"

        command.Parameters.Add(New SqlParameter("bd_code", SqlDbType.VarChar, 1)).Value = tb_bd_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("fl_code", SqlDbType.VarChar, 2)).Value = tb_fl_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("rnum_code", SqlDbType.VarChar, 5)).Value = tb_rnum_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("wa_code", SqlDbType.VarChar, 1)).Value = tb_wa_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("bg_code", SqlDbType.VarChar, 2)).Value = tb_bg_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("it_code", SqlDbType.VarChar, 6)).Value = tb_it_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("element_no", SqlDbType.VarChar, 3)).Value = tb_element_no.Text.Trim()

        command.Parameters.Add(New SqlParameter("element_code", SqlDbType.VarChar, 20)) _
                          .Value = String.Format( _
                                                 "{0}{1}{2}{3}{4}{5}{6}", _
                                                   command.Parameters("bd_code").Value, _
                                                   command.Parameters("fl_code").Value, _
                                                   command.Parameters("rnum_code").Value, _
                                                   command.Parameters("wa_code").Value, _
                                                   command.Parameters("bg_code").Value, _
                                                   command.Parameters("it_code").Value, _
                                                   command.Parameters("element_no").Value _
                                                )

        command.Parameters.Add(New SqlParameter("operator", SqlDbType.VarChar, 10)).Value = liEmployee.SelectedValue.Trim()

        Try
            command.Connection.Open()

            If (command.ExecuteNonQuery() > 0) Then

                scripts.AppendFormat( _
                                     "alert('所輸入的設備編碼代碼為: {0}\n');" & _
                                     "location.href = 'MOA04010.aspx';", _
                                      command.Parameters("element_code").Value _
                                    )

                ClientScript.RegisterStartupScript(Me.GetType(), "MOA04011_RecordInserted", scripts.ToString(), True)
                scripts.Remove(0, scripts.Length)

            End If

        Catch ex As Exception

            If (TypeOf ex Is SqlException) Then

                If (CType(ex, SqlException).Number.Equals(2627)) Then
                    ClientScript.RegisterStartupScript(Me.GetType(), "SqlErrorMsg", "alert('該筆資料已存在 !');", True)
                End If

            End If

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

        End Try

    End Sub

    Protected Sub ibtnPrevious_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnPrevious.Click

        Server.Transfer("MOA04010.aspx")

    End Sub

End Class