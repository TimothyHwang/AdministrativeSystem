'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Create at 2010/06/17
'   Description : First version of development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/18
'   Description : (1) 修正流水號控制規則。
'                 (2) 將頁面改為新增與修改時共用。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/21
'   Description : 修正 Javascript 的 StringValidationAfterTrim 之使用方法。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/24
'   Description : 修正執行查詢後，GridView 操作異動作業時，呈現錯誤問題。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/30
'   Description : 將相關輸入項目改為 DropDownList 控制項以供選擇；並為其加入相關篩選機制。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/07/01
'   Description : 修正修改模式下，將預設值帶入各 DropDownList 控制項的機制。
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Imports WebUtilities.Functions

Partial Class Source_04_MOA04011
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

        If Page.IsPostBack Then

            If Not String.IsNullOrEmpty(Request.Form("hid_it_code")) Then

                Me.ClientScript.RegisterHiddenField("hid_it_code", Request.Form("hid_it_code"))
                Me.tb_it_code.Text = Request.Form("hid_it_code").Trim(",")

            End If

        Else

            Me.DropDownListControls_Databind()

            If String.IsNullOrEmpty(Request.Form("element_no")) Then

                ibtnModify.Enabled = False
                ibtnModify.Visible = False

                Label1.Text = "設備識別編號新增"

                mvMultiInputParameters.ActiveViewIndex = 0

            Else

                lbl_element_no.Text = Request.Form("element_no")

                Dim selected_item As ListItem = Nothing

                selected_item = ddl_bd_code.Items.FindByValue(Request.Form("bd_code"))
                CType(IIf(IsNothing(selected_item), ddl_bd_code.Items(0), selected_item), ListItem).Selected = True

                selected_item = ddl_fl_code.Items.FindByValue(Request.Form("fl_code"))
                CType(IIf(IsNothing(selected_item), ddl_fl_code.Items(0), selected_item), ListItem).Selected = True

                Me.DropDownLists_RnumCode_Filter_SelectedIndexChanged(Nothing, Nothing)

                selected_item = ddl_rnum_code.Items.FindByValue(Request.Form("rnum_code"))
                CType(IIf(IsNothing(selected_item), ddl_rnum_code.Items(0), selected_item), ListItem).Selected = True

                selected_item = ddl_wa_code.Items.FindByValue(Request.Form("wa_code"))
                CType(IIf(IsNothing(selected_item), ddl_wa_code.Items(0), selected_item), ListItem).Selected = True

                selected_item = ddl_bg_code.Items.FindByValue(Request.Form("bg_code"))
                CType(IIf(IsNothing(selected_item), ddl_bg_code.Items(0), selected_item), ListItem).Selected = True

                ClientScript.RegisterHiddenField("hid_it_code", Request.Form("it_code"))

                tb_it_code.Text = Request.Form("it_code").Trim(New Char() {" ", ",", "?"})

                'Setting DropDownLists Filter SelectedValue.
                If Not String.IsNullOrEmpty(tb_it_code.Text) Then

                    Dim it_code As String = tb_it_code.Text.PadRight(6, " ")

                    selected_item = ddl_it_code_L1.Items.FindByValue(it_code.Substring(0, 1).PadRight(6, "0"))

                    DirectCast(IIf(IsNothing(selected_item), ddl_it_code_L1.Items(0), selected_item), ListItem).Selected = True

                    If Not String.IsNullOrEmpty(ddl_it_code_L1.SelectedValue) Then
                        Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L1, Nothing)
                    End If

                    selected_item = ddl_it_code_L2.Items.FindByValue(it_code.Substring(0, 2).PadRight(6, "0"))

                    DirectCast(IIf(IsNothing(selected_item), ddl_it_code_L2.Items(0), selected_item), ListItem).Selected = True

                    If Not String.IsNullOrEmpty(ddl_it_code_L2.SelectedValue) Then
                        Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L2, Nothing)
                    End If

                    selected_item = ddl_it_code_L3.Items.FindByValue(it_code.Substring(0, 3).PadRight(6, "0"))

                    DirectCast(IIf(IsNothing(selected_item), ddl_it_code_L3.Items(0), selected_item), ListItem).Selected = True

                    If Not String.IsNullOrEmpty(ddl_it_code_L2.SelectedValue) Then

                        Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L3, Nothing)

                        selected_item = list_it_code.Items.FindByValue(it_code)

                        If Not IsNothing(selected_item) Then
                            selected_item.Selected = True
                        End If

                    End If

                    it_code = Nothing

                End If

                selected_item = Nothing

                Me.ViewState("element_code") = Request.Form("element_code")

                ibtnAdd.Enabled = False
                ibtnAdd.Visible = False

                Label1.Text = "設備識別編號修改"

                mvMultiInputParameters.ActiveViewIndex = 1

            End If

            Page.Header.Title = Label1.Text

            If Not String.IsNullOrEmpty(Request.Form("query_element_no")) Then
                Me.ViewState("query_element_no") = Request.Form("query_element_no").Trim()
            End If

            If Not String.IsNullOrEmpty(Request.Form("query_element_code")) Then
                Me.ViewState("query_element_code") = Request.Form("query_element_code").Trim()
            End If

        End If

        Me.ImageButtons_ClientScriptsPreparing()
        Me.WebControllers_ClientEventsBinding()

    End Sub

    Private Sub DropDownListControls_Databind()

        Dim adapter As New SqlDataAdapter(String.Empty, (New C_SQLFUN).G_conn_string)

        Try
            adapter.SelectCommand.Connection.Open()

        Catch ex As Exception
            scripts.Append("- 資料庫連線初化始失敗 !\n")
        End Try

        If adapter.SelectCommand.Connection.State.Equals(ConnectionState.Open) Then

            Dim data_source As New DataTable()

            '建物代碼資料繫結
            adapter.SelectCommand.CommandText = "Select bd_name, bd_code From P_0404 Order by no"

            Try
                adapter.Fill(data_source)

            Catch ex As Exception

                scripts.Append("- 建物項目初化始失敗 !\n")
                data_source.Clear()

            End Try

            ddl_bd_code.DataSource = data_source
            ddl_bd_code.DataBind()

            ddl_bd_code.Items.Insert(0, New ListItem())
            data_source.Clear()

            '樓層代碼資料繫結
            adapter.SelectCommand.CommandText = "Select fl_name, fl_code From P_0406 Order by fl_no"

            Try
                adapter.Fill(data_source)

            Catch ex As Exception

                scripts.Append("- 樓層項目初始化失敗 !\n")
                data_source.Clear()

            End Try

            ddl_fl_code.DataSource = data_source
            ddl_fl_code.DataBind()

            ddl_fl_code.Items.Insert(0, New ListItem())
            data_source.Clear()

            '房間地理位置代碼資料繫結
            adapter.SelectCommand.CommandText = "Select rnum_name, rnum_code From P_0410 Order by rnum_code"

            Try
                adapter.Fill(data_source)

            Catch ex As Exception

                scripts.Append("- 房間(區域)項目初始化失敗 !\n")
                data_source.Clear()

            End Try

            ddl_rnum_code.DataSource = data_source
            ddl_rnum_code.DataBind()

            data_source.Clear()

            For Each item As ListItem In ddl_rnum_code.Items
                item.Text = String.Format("{0} - {1}", item.Value, item.Text)
            Next

            ddl_rnum_code.Items.Insert(0, New ListItem())

            '房間規格代碼資料繫結
            adapter.SelectCommand.CommandText = "Select wa_name, wa_code From P_0413 Order by no"

            Try
                adapter.Fill(data_source)

            Catch ex As Exception

                scripts.Append("- 牆柱項目初始化失敗 !\n")
                data_source.Clear()

            End Try

            ddl_wa_code.DataSource = data_source
            ddl_wa_code.DataBind()

            ddl_wa_code.Items.Insert(0, New ListItem())
            data_source.Clear()

            '預算代碼資料繫結
            adapter.SelectCommand.CommandText = "Select bg_name, bg_code From P_0403 Order by no"

            Try
                adapter.Fill(data_source)

            Catch ex As Exception

                scripts.Append("- 預算項目初始化失敗 !\n")
                data_source.Clear()

            End Try

            ddl_bg_code.DataSource = data_source
            ddl_bg_code.DataBind()

            ddl_bg_code.Items.Insert(0, New ListItem())
            data_source.Clear()

            '物料代碼資料繫結
            adapter.SelectCommand.CommandText = "Select it_name = Case IsNull(it_spec, '') When '' Then it_name " & _
                                                "Else (it_name + ' - ' + it_spec) End, it_code From " & _
                                                "P_0407 Where it_code Like '%00000' Order by it_name;"

            Try
                adapter.Fill(data_source)

            Catch ex As Exception

                scripts.Append("- 物料編號項目初始化失敗 !\n")
                data_source.Clear()

            End Try

            ddl_it_code_L1.DataSource = data_source
            ddl_it_code_L1.DataBind()

            data_source.Clear()

            If Not data_source Is Nothing Then

                data_source.Dispose()
                data_source = Nothing

            End If

            ddl_it_code_L2.DataBind()
            ddl_it_code_L3.DataBind()

        End If

        If adapter.SelectCommand.Connection.State.Equals(ConnectionState.Open) Then
            adapter.SelectCommand.Connection.Close()
        End If

        If Not adapter Is Nothing Then

            adapter.Dispose()
            adapter = Nothing

        End If

        If Not scripts.Length.Equals(0) Then

            scripts.Insert(0, "alert('").Append("');")

            ClientScript.RegisterStartupScript(Me.GetType(), "DropDownLists_DataBind", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

        End If

    End Sub

    Private Function ElementNumberGeneration() As String

        ElementNumberGeneration = String.Empty

        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        command.CommandType = Data.CommandType.Text

        command.CommandText = "Select Top 1 (Cast(element_no As Int) + 1) As element_no From (Select element_no, " & _
                              "element_code From P_0405 Where element_code Like (@element_code + '%') Union " & _
                              "Select '0' As element_no, Null As element_code ) As [TEMP] " & _
                              "Order by element_code DESC"

        command.Parameters.Add(New SqlParameter("element_code", SqlDbType.VarChar, 20)).Value = _
                String.Format( _
                                    "{0}{1}{2}{3}{4}{5}", _
                                      ddl_bd_code.SelectedValue.Trim(), _
                                      ddl_fl_code.SelectedValue.Trim(), _
                                      ddl_rnum_code.SelectedValue.Trim(), _
                                      ddl_wa_code.SelectedValue.Trim(), _
                                      ddl_bg_code.SelectedValue.Trim(), _
                                      Me.Request.Form("hid_it_code").Trim(New Char() {" ", ","}) _
                             )

        Try
            command.Connection.Open()

            ElementNumberGeneration = CType(command.ExecuteScalar(), Int32).ToString("000")

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

    End Function

    Private Sub ImageButtons_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ")

        If mvMultiInputParameters.ActiveViewIndex.Equals(0) Then

            scripts.AppendFormat("{{ client_id: '{0}', ", tb_records_count.ClientID) _
                   .Append("blank_validation: [{ message: '請輸入資料新增筆數' }], ") _
                   .Append("numeric_validation: [{ message: '資料新增筆數請輸入數字' }] }, ")

        End If

        scripts.AppendFormat("{{ client_id: '{0}', ", ddl_bd_code.ClientID) _
               .Append("blank_validation: [{ message: '請選擇建物編號' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", ddl_fl_code.ClientID) _
               .Append("blank_validation: [{ message: '請選擇樓層編號' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", ddl_rnum_code.ClientID) _
               .Append("blank_validation: [{ message: '請選擇房間(區域)編號' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", ddl_wa_code.ClientID) _
               .Append("blank_validation: [{ message: '請選擇牆柱編號' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", ddl_bg_code.ClientID) _
               .Append("blank_validation: [{ message: '請選擇預算編號' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_it_code.ClientID) _
               .Append("blank_validation: [{ message: '請選擇物料編號' }] } ") _
               .Append("]});")

        ibtnAdd.Attributes.Add("OnClick", scripts.ToString())
        ibtnModify.Attributes.Add("OnClick", scripts.ToString())

        scripts.Remove(0, scripts.Length)

    End Sub

    Private Sub WebControllers_ClientEventsBinding()

        list_it_code.Attributes("OnClick") = String.Format _
                                                    ( _
                                                          "javascript:" & _
                                                          "document.getElementById('hid_it_code').value = event.srcElement.value;" & _
                                                          "document.getElementById('{0}').value = event.srcElement.value;", _
                                                           Me.tb_it_code.ClientID _
                                                    )

    End Sub

    Protected Sub DropDownLists_RnumCode_Filter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
                Handles ddl_bd_code.SelectedIndexChanged, ddl_fl_code.SelectedIndexChanged

        Dim adapter As New SqlDataAdapter( _
                                                "Select rnum_name, rnum_code From P_0410 Where bd_code = IsNull(@bd_code, bd_code) " & _
                                                "And fl_code = IsNull(@fl_code, fl_code) Order by rnum_code ", _
                                                (New C_SQLFUN).G_conn_string _
                                         )

        adapter.SelectCommand.Parameters.Clear()
        adapter.SelectCommand.Parameters.Add(New SqlParameter("bd_code", SqlDbType.VarChar, 1)).Value = ddl_bd_code.SelectedValue.Trim()
        adapter.SelectCommand.Parameters.Add(New SqlParameter("fl_code", SqlDbType.VarChar, 2)).Value = ddl_fl_code.SelectedValue.Trim()

        For Each param As SqlParameter In adapter.SelectCommand.Parameters
            If String.IsNullOrEmpty(param.Value) Then
                param.Value = DBNull.Value
            End If
        Next

        Dim data_source As New DataTable()

        Try
            adapter.SelectCommand.Connection.Open()
            adapter.Fill(data_source)

        Catch ex As Exception

            scripts.Append("- 房間(區域)項目初始化失敗 !\n")

            data_source.Clear()
            data_source.Dispose()

            data_source = Nothing

        Finally

            If Not data_source Is Nothing Then

                ddl_rnum_code.DataSource = data_source
                ddl_rnum_code.DataBind()

                data_source.Dispose()
                data_source = Nothing

                For Each item As ListItem In ddl_rnum_code.Items
                    item.Text = String.Format("{0} - {1}", item.Value, item.Text)
                Next

                ddl_rnum_code.Items.Insert(0, New ListItem())

            End If

            If adapter.SelectCommand.Connection.State.Equals(ConnectionState.Open) Then
                adapter.SelectCommand.Connection.Close()
            End If

        End Try

        If Not adapter Is Nothing Then

            adapter.Dispose()
            adapter = Nothing

        End If

        If Not scripts.Length.Equals(0) Then

            scripts.Insert(0, "alert('").Append("');")

            ClientScript.RegisterStartupScript(Me.GetType(), "DropDownList_RnumCode_DataBind", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

        End If

    End Sub

    Protected Sub DropDownLists_ItCode_Filter_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_it_code_L1.DataBound, ddl_it_code_L2.DataBound, ddl_it_code_L3.DataBound

        DirectCast(sender, DropDownList).Items.Insert(0, New ListItem("請選擇", String.Empty))

    End Sub

    Protected Sub DropDownLists_ItCode_Filter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_it_code_L1.SelectedIndexChanged, ddl_it_code_L2.SelectedIndexChanged, ddl_it_code_L3.SelectedIndexChanged

        Dim ddlControl As DropDownList = DirectCast(sender, DropDownList)

        If TypeOf ddlControl Is DropDownList Then

            If String.IsNullOrEmpty(ddlControl.SelectedValue) Then
                Me.tb_it_code.Text = String.Empty
                Me.ClientScript.RegisterHiddenField("hid_it_code", String.Empty)
            End If

            Dim length As Integer = Nothing
            Dim filter_expression As String = String.Empty

            If Integer.TryParse(ddlControl.ClientID(ddlControl.ClientID.Length - 1).ToString(), length) Then

                If Not String.IsNullOrEmpty(ddlControl.SelectedValue) Then

                    filter_expression = ddlControl.SelectedValue.Substring(0, length)
                    filter_expression += "%"
                    filter_expression = IIf(length.Equals(3), filter_expression, filter_expression.PadRight(6, "0"))
                    filter_expression = filter_expression.Insert(0, "'") & "'"

                End If

            End If

            Dim data_source As DataSet = Me.RetrieveItCodeFilter_DataSource(filter_expression)

            Select Case ddlControl.ClientID

                Case "ddl_it_code_L1"

                    ddl_it_code_L2.Items.Clear()
                    ddl_it_code_L3.Items.Clear()

                    list_it_code.Items.Clear()

                    ddl_it_code_L2.DataSource = data_source
                    ddl_it_code_L2.DataBind()
                    ddl_it_code_L3.DataBind()

                Case "ddl_it_code_L2"

                    ddl_it_code_L3.Items.Clear()

                    list_it_code.Items.Clear()

                    ddl_it_code_L3.DataSource = data_source
                    ddl_it_code_L3.DataBind()

                Case "ddl_it_code_L3"

                    list_it_code.Items.Clear()

                    list_it_code.DataSource = data_source
                    list_it_code.DataBind()

            End Select

        End If

    End Sub

    Private Function RetrieveItCodeFilter_DataSource(ByVal expression As String) As DataSet

        RetrieveItCodeFilter_DataSource = New DataSet

        Dim adapter As New SqlDataAdapter(String.Empty, ConfigurationManager.ConnectionStrings.Item("ConnectionString").ConnectionString)

        adapter.SelectCommand.CommandType = CommandType.Text

        adapter.SelectCommand.CommandText = "Select it_name = Case IsNull(it_spec, '') When '' Then it_name Else (it_name + ' - ' + it_spec) End, " & _
                                            "it_code From P_0407 Where it_code Like " & expression & " And it_code <> '" & _
                                             expression.Replace("%", "0").Trim("'"c).PadRight(6, "0") & "' " & _
                                            "Order by it_name;"

        Try
            adapter.SelectCommand.Connection.Open()
            adapter.Fill(RetrieveItCodeFilter_DataSource)

        Catch ex As Exception
            RetrieveItCodeFilter_DataSource = Nothing

        Finally

            If adapter.SelectCommand.Connection.State.Equals(ConnectionState.Open) Then
                adapter.SelectCommand.Connection.Close()
            End If

            adapter.Dispose()
            adapter = Nothing

        End Try

    End Function

    Protected Sub ImageButtons_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim initial_number As Integer

        If Integer.TryParse(Me.ElementNumberGeneration(), initial_number) Then

            Dim sql_function As New C_SQLFUN
            Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

            command.CommandType = CommandType.Text

            scripts.Append("Insert Into P_0405(bd_code, fl_code, rnum_code, wa_code, bg_code, it_code, ") _
                   .Append("element_no, element_code, insertime, operator) Values(@bd_code, @fl_code, ") _
                   .Append("@rnum_code, @wa_code, @bg_code, @it_code, @element_no, @element_code, ") _
                   .Append("GETDATE(), @operator);")

            command.CommandText = scripts.ToString()
            scripts.Remove(0, scripts.Length)

            command.Parameters.Add(New SqlParameter("bd_code", SqlDbType.VarChar, 1)).Value = ddl_bd_code.SelectedValue.Trim()
            command.Parameters.Add(New SqlParameter("fl_code", SqlDbType.VarChar, 2)).Value = ddl_fl_code.SelectedValue.Trim()
            command.Parameters.Add(New SqlParameter("rnum_code", SqlDbType.VarChar, 5)).Value = ddl_rnum_code.SelectedValue.Trim()
            command.Parameters.Add(New SqlParameter("wa_code", SqlDbType.VarChar, 1)).Value = ddl_wa_code.SelectedValue.Trim()
            command.Parameters.Add(New SqlParameter("bg_code", SqlDbType.VarChar, 2)).Value = ddl_bg_code.SelectedValue.Trim()
            command.Parameters.Add(New SqlParameter("it_code", SqlDbType.VarChar, 6)).Value = Me.Request.Form("hid_it_code").Trim(New Char() {" ", ","})
            command.Parameters.Add(New SqlParameter("element_no", SqlDbType.VarChar, 3)).Value = DBNull.Value
            command.Parameters.Add(New SqlParameter("operator", SqlDbType.VarChar, 10)).Value = user_id

            command.Parameters.Add(New SqlParameter("element_code", SqlDbType.VarChar, 20)) _
                              .Value = String.Format _
                                              ( _
                                                   "{0}{1}{2}{3}{4}{5}{6}", _
                                                     command.Parameters("bd_code").Value, _
                                                     command.Parameters("fl_code").Value, _
                                                     command.Parameters("rnum_code").Value, _
                                                     command.Parameters("wa_code").Value, _
                                                     command.Parameters("bg_code").Value, _
                                                     command.Parameters("it_code").Value, _
                                                     command.Parameters("element_no").Value _
                                              )

            command.Transaction = Nothing

            Dim element_code_partial As String = DirectCast(command.Parameters("element_code").Value, String)

            Try

                command.Connection.Open()
                command.Transaction = command.Connection.BeginTransaction(IsolationLevel.ReadCommitted)

                For element_code As Integer = initial_number To ((initial_number + CType(tb_records_count.Text, Integer)) - 1)

                    command.Parameters("element_no").Value = element_code.ToString("000")
                    command.Parameters("element_code").Value = element_code_partial & DirectCast(command.Parameters("element_no").Value, String)

                    command.ExecuteNonQuery()

                Next

                command.Transaction.Commit()

            Catch ex As Exception

                command.Transaction.Rollback()
                scripts.Append(ex.Message.Replace("'"c, """"c).Replace(Environment.NewLine, "\n"))

            Finally

                If TypeOf command.Transaction Is SqlTransaction Then
                    command.Transaction.Dispose()
                    command.Transaction = Nothing
                End If

                If command.Connection.State.Equals(ConnectionState.Open) Then
                    command.Connection.Close()
                End If

                element_code_partial = Nothing

            End Try

        Else
            scripts.Append("設備識別編號產生失敗 !!\n無法完成資料新增作業。")
        End If

        If scripts.Length.Equals(0) Then
            ibtnPrevious_Click(Me.ibtnPrevious, Nothing)
        Else

            scripts.Insert(0, "alert('").Append("');")

            ClientScript.RegisterStartupScript(Me.GetType(), "RecordsInserted_Message", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

        End If

    End Sub

    Protected Sub ibtnPrevious_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnPrevious.Click

        Dim parameters As New SortedList()

        parameters.Add("Action", "MOA04010.aspx")

        If Not String.IsNullOrEmpty(Me.ViewState("query_element_no")) Then
            parameters.Add("query_element_no", Me.ViewState("query_element_no"))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("query_element_code")) Then
            parameters.Add("query_element_code", Me.ViewState("query_element_code"))
        End If

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()

    End Sub

End Class