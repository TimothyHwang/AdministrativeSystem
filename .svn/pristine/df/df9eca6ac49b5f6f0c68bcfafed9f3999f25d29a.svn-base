Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO

Imports WebUtilities.Functions

Partial Class M_Source_04_MOA04026
    Inherits System.Web.UI.Page
    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder
    Public sPicPath As String = ConfigurationManager.AppSettings("PicPathUrl")

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

            Me.DropDownListControls_Databind()

            If String.IsNullOrEmpty(Request.Form("shcode")) Then


                mvMultiInputParameters.ActiveViewIndex = 0

                lblTitle.Text = "廢料入庫資料新增"

                tb_shcode_middle.Text = Date.Today.ToString("yyyyMMdd")

            Else

                ibtnAdd.Enabled = False
                ibtnAdd.Visible = False

                mvMultiInputParameters.ActiveViewIndex = 1

                lblTitle.Text = "廢料入庫資料修改"

                Dim selected_item As ListItem = Nothing
                Dim shcode_partial As String = Left(Request.Form("shcode").PadLeft(18, " "c), 6)

                '物料代碼篩選器第一層之 DropDownList 預設值設定
                selected_item = ddl_it_code_L1.Items.FindByValue(Left(shcode_partial, 1).Trim().PadRight(6, "0"c))
                selected_item = IIf(IsNothing(selected_item), ddl_it_code_L1.Items(0), selected_item)

                selected_item.Selected = True

                Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L1, Nothing)

                '物料代碼篩選器第二層之 DropDownList 預設值設定
                selected_item = ddl_it_code_L2.Items.FindByValue(Left(shcode_partial, 2).Trim().PadRight(6, "0"c))
                selected_item = IIf(IsNothing(selected_item), ddl_it_code_L2.Items(0), selected_item)

                selected_item.Selected = True

                Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L2, Nothing)

                '物料代碼篩選器第三層之 DropDownList 預設值設定
                selected_item = ddl_it_code_L3.Items.FindByValue(Left(shcode_partial, 3).Trim().PadRight(6, "0"c))
                selected_item = IIf(IsNothing(selected_item), ddl_it_code_L3.Items(0), selected_item)

                selected_item.Selected = True

                Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L3, Nothing)

                '物料代碼篩選器最末層之 ListBox 預設值設定
                selected_item = list_it_code.Items.FindByValue(shcode_partial.Trim())

                If Not IsNothing(selected_item) Then
                    selected_item.Selected = True
                    selected_item = Nothing
                End If

                '物品編號的 TextBox 預設值設定
                shcode_partial = Request.Form("shcode").PadLeft(18, " "c)

                tb_shcode_first.Text = Left(shcode_partial, 6)
                tb_shcode_middle.Text = Right(Left(shcode_partial, 14), 8)
                tb_shcode_last.Text = Right(shcode_partial, 4)

                '其他輸入項的 TextBox 預設值設定
                'tb_it_spec.Text = Request.Form("Buy_Num")
                tb_memo.Text = Request.Form("memo")

                Me.ViewState("shcode_original") = Request.Form("shcode")

            End If

            'Common Configuration with Creation/Modification Mode.
            Page.Header.Title = Label1.Text

            If Not String.IsNullOrEmpty(Request.Form("query_shcode")) Then
                Me.ViewState("query_shcode") = Request.Form("query_shcode").Trim()
            End If

            If Not String.IsNullOrEmpty(Request.Form("query_company")) Then
                Me.ViewState("query_company") = Request.Form("query_company").Trim()
            End If

        End If

        Me.ImageButtons_ClientScriptsPreparing()

    End Sub

    Private Sub DropDownListControls_Databind()

        sdsItCodeFilter.SelectParameters("it_code").DefaultValue = "%".PadRight(6, "0"c)
        sdsItCodeFilter.SelectParameters("it_code_ignore").DefaultValue = String.Empty

        Dim data_source As DataTable = DirectCast(sdsItCodeFilter.Select(New DataSourceSelectArguments), DataView).Table

        If Not IsNothing(data_source) Then

            ddl_it_code_L1.DataSource = data_source
            data_source = Nothing

        End If

        ddl_it_code_L1.DataBind()
        ddl_it_code_L2.DataBind()
        ddl_it_code_L3.DataBind()

    End Sub

    Private Sub FileUploadControls_Configuration(ByRef multi_viewer As MultiView, ByVal file_name As String)

        Dim prefix As String = String.Format _
                                      ( _
                                           "{0}{1}{2}_", _
                                             Me.tb_shcode_first.Text.Trim(), _
                                             Me.tb_shcode_middle.Text.Trim(), _
                                             Right(multi_viewer.ID, 1).ToUpper() _
                                      )

        If String.IsNullOrEmpty(file_name) Then

            For Each file_path As String In Directory.GetFiles(Me.Server.MapPath("~/M_Source/99"), prefix & "*.*")

                file_name = file_path.Substring(file_path.LastIndexOf("\"c) + 1).Trim()
                file_name = file_name.Substring(prefix.Length)

                Exit For

            Next

        End If

        If String.IsNullOrEmpty(file_name) Then
            multi_viewer.ActiveViewIndex = 0
        Else

            multi_viewer.ActiveViewIndex = 1

            DirectCast(Me.FindControl("lkb_" & multi_viewer.ID.Substring(3)), LinkButton).Text = file_name

        End If

    End Sub

    Protected Sub sdsItCodeFilter_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles sdsItCodeFilter.Selecting

        For Each element As DbParameter In e.Command.Parameters
            If String.IsNullOrEmpty(element.Value) Then
                element.Value = DBNull.Value
            End If
        Next

    End Sub

    Private Sub ImageButtons_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ")

        If mvMultiInputParameters.ActiveViewIndex.Equals(0) Then

            scripts.AppendFormat("{{ client_id: '{0}', ", tb_records_count.ClientID) _
                   .Append("blank_validation: [{ message: '請輸入資料新增筆數' }], ") _
                   .Append("numeric_validation: [{ message: '資料新增筆數請輸入數字' }] }, ")

        End If

        scripts.AppendFormat("{{ client_id: '{0}', ", tb_shcode_first.ClientID) _
               .Append("blank_validation: [{ message: '請輸入物品編號前六碼' }] }, ") _
                .AppendFormat("{{ client_id: '{0}', ", tb_it_spec.ClientID) _
               .Append("blank_validation: [{ message: '請輸入物料規格' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_shcode_middle.ClientID) _
               .Append("blank_validation: [{ message: '請輸入物品編號中間八碼' }], ") _
               .Append("numeric_validation: [{ message: '物品編號中間八碼請輸入數字' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_shcode_last.ClientID) _
               .Append("blank_validation: [{ message: '請輸入物品編號後四碼' }], ") _
               .Append("numeric_validation: [{ message: '物品編號後四碼請輸入數字' }] }") _
               .Append("]});")

        ibtnAdd.Attributes.Add("OnClick", scripts.ToString())

        scripts.Remove(0, scripts.Length)

    End Sub

    Private Function SHCode_LastNumeric_Generation() As String
        SHCode_LastNumeric_Generation = String.Empty

        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        scripts.Append("Select Top 1 (Cast(Right(shcode, 4) As Int) + 1) As shcode From (select shcode ") _
               .Append("From P_0414 Where shcode Like (@shcode +'%')) As [TMP] ") _
               .Append("Order by shcode DESC;")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)


        Dim sitem As String = tb_shcode_first.Text & tb_shcode_middle.Text
        command.Parameters.Add(New SqlParameter("shcode", SqlDbType.NVarChar, 14)).Value = sitem

        Try
            command.Connection.Open()
            Dim exobject As New Object
            exobject = command.ExecuteScalar()
            If exobject = Nothing Then
                SHCode_LastNumeric_Generation = "0001"
            Else
                SHCode_LastNumeric_Generation = DirectCast(command.ExecuteScalar(), Integer).ToString("0000")
            End If
        Catch ex As Exception

            SHCode_LastNumeric_Generation = Nothing

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

            command.Dispose()
            command = Nothing

        End Try

    End Function
    Protected Sub DropDownLists_ItCode_Filter_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_it_code_L1.DataBound, ddl_it_code_L2.DataBound, ddl_it_code_L3.DataBound

        DirectCast(sender, DropDownList).Items.Insert(0, New ListItem("請選擇", String.Empty))

    End Sub

    Protected Sub DropDownLists_ItCode_Filter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_it_code_L1.SelectedIndexChanged, ddl_it_code_L2.SelectedIndexChanged, ddl_it_code_L3.SelectedIndexChanged

        Dim ddlControl As DropDownList = DirectCast(sender, DropDownList)

        If TypeOf ddlControl Is DropDownList Then

            If String.IsNullOrEmpty(ddlControl.SelectedValue) Then
                tb_shcode_first.Text = String.Empty
                tb_shcode_last.Text = String.Empty
            End If

            Dim length As Integer = Nothing

            If Integer.TryParse(ddlControl.ClientID.Chars(ddlControl.ClientID.Length - 1).ToString(), length) Then

                Dim filter_expression As String = Left(ddlControl.SelectedValue.PadRight(6, " "c), length).Trim()

                filter_expression += IIf(String.IsNullOrEmpty(filter_expression), String.Empty, "%")
                filter_expression = IIf(String.IsNullOrEmpty(filter_expression), String.Empty, filter_expression.PadRight(6, "0"c))

                Select Case ddlControl.ClientID

                    Case "ddl_it_code_L1", "ddl_it_code_L2"
                        sdsItCodeFilter.SelectParameters.Item("it_code").DefaultValue = filter_expression

                    Case "ddl_it_code_L3"
                        sdsItCodeFilter.SelectParameters.Item("it_code").DefaultValue = Regex.Replace(filter_expression, "[%][0]*", "%")

                End Select

                sdsItCodeFilter.SelectParameters.Item("it_code_ignore").DefaultValue = filter_expression.Replace("%", "0")

                Select Case ddlControl.ClientID

                    Case "ddl_it_code_L1"
                        ddl_it_code_L2.Items.Clear()
                        ddl_it_code_L3.Items.Clear()
                        list_it_code.Items.Clear()
                        ddl_it_code_L2.DataSource = sdsItCodeFilter.Select(New DataSourceSelectArguments)
                        ddl_it_code_L2.DataBind()
                        ddl_it_code_L3.DataBind()

                    Case "ddl_it_code_L2"
                        ddl_it_code_L3.Items.Clear()
                        list_it_code.Items.Clear()
                        ddl_it_code_L3.DataSource = sdsItCodeFilter.Select(New DataSourceSelectArguments)
                        ddl_it_code_L3.DataBind()

                    Case "ddl_it_code_L3"
                        list_it_code.Items.Clear()
                        Dim itemdataview As New DataView
                        itemdataview = sdsItCodeFilter.Select(New DataSourceSelectArguments)
                        itemdataview.RowFilter = "it_sort in ('零料件')"
                        list_it_code.DataSource = itemdataview
                        list_it_code.DataBind()

                End Select

            End If

            length = Nothing

        End If

        ddlControl = Nothing

    End Sub

    Protected Sub list_it_code_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_it_code.SelectedIndexChanged
        Dim s_ItemCodeValue As String = (DirectCast(sender, ListBox).SelectedValue.Trim())
        Dim itemarray() As String = Split(s_ItemCodeValue, ",")
        Me.tb_shcode_first.Text = itemarray(0)
        Me.tb_it_spec.Text = itemarray(5)
        Dim sImagePath As String = sPicPath
        If itemarray(3).Length = 0 Then
            lb_image1.Visible = True
            it_image1.Visible = False
        Else
            Dim image_file_path1 As String = String.Format _
                                      ( _
                                           "{0}{1}{2}_{3}", _
                                             sImagePath, _
                                             Me.tb_shcode_first.Text, _
                                             "A", _
                                             itemarray(3) _
                                      )
            lb_image1.Visible = False
            it_image1.Visible = True
            it_image1.ImageUrl = image_file_path1
        End If
        If itemarray(4).Length = 0 Then
            lb_image2.Visible = True
            it_image2.Visible = False
        Else
            Dim image_file_path2 As String = String.Format _
                                        ( _
                                             "{0}{1}{2}_{3}", _
                                               sImagePath, _
                                               Me.tb_shcode_first.Text, _
                                               "B", _
                                               itemarray(4) _
                                        )
            lb_image2.Visible = False
            it_image2.Visible = True
            it_image2.ImageUrl = image_file_path2
        End If
        If mvMultiInputParameters.ActiveViewIndex.Equals(0) Then
            Me.tb_shcode_last.Text = Me.SHCode_LastNumeric_Generation()
        Else

            If (tb_shcode_first.Text.Trim() & tb_shcode_middle.Text.Trim()).Equals(Left(Me.ViewState("shcode_original").ToString().PadRight(18, " "c), 14).Trim()) Then
                Me.tb_shcode_last.Text = Right(Me.ViewState("shcode_original").ToString().PadRight(18, " "c), 4).Trim()
            Else
                Me.tb_shcode_last.Text = Me.SHCode_LastNumeric_Generation()
            End If
        End If

        'reload數量編號範圍
        Showschcode_lastest()
    End Sub

    Protected Sub ImageButtons_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim command As New SqlCommand _
                           ( _
                                    String.Empty, _
                                    New SqlConnection _
                                        ( _
                                                    ConfigurationManager.ConnectionStrings("ConnectionString") _
                                                                        .ConnectionString _
                                        ) _
                           )

        scripts.Append("Insert Into P_0414 (shcode, it_spec, Seat_Num, memo, shtype,") _
                       .Append("UseCheck,StockKeyIn, StockDate) Values (@shcode, @it_spec, @Seat_Num, @memo, ") _
                       .Append("@shtype, @UseCheck, @StockKeyIn, getdate());")


        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("shcode", SqlDbType.VarChar, 18)).Value = String.Empty
        command.Parameters.Add(New SqlParameter("it_spec", SqlDbType.VarChar, 255)).Value = tb_it_spec.Text.Trim()
        command.Parameters.Add(New SqlParameter("Seat_Num", SqlDbType.VarChar, 20)).Value = ddl_seat_num.SelectedValue
        command.Parameters.Add(New SqlParameter("memo", SqlDbType.NVarChar, 255)).Value = tb_memo.Text.Trim()
        command.Parameters.Add(New SqlParameter("shtype", SqlDbType.VarChar, 1)).Value = "1"
        command.Parameters.Add(New SqlParameter("UseCheck", SqlDbType.VarChar, 1)).Value = "0"
        command.Parameters.Add(New SqlParameter("StockKeyIn", SqlDbType.NVarChar, 10)).Value = user_id

        For Each element As SqlParameter In command.Parameters
            If String.IsNullOrEmpty(element.Value) Then
                element.Value = DBNull.Value
            End If
        Next

        Dim processed_message As String = Nothing
        processed_message = Me.DatabaseRecords_InsertionProcess(command)

        If Not IsNothing(command) Then

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
                command.Connection.Dispose()
            End If

            command.Dispose()
            command = Nothing

        End If

        If String.IsNullOrEmpty(processed_message) Then
            ibtnPrevious_Click(Me.ibtnPrevious, Nothing)
        Else
            scripts.Append("alert('").Append(processed_message).Append("');")

            ClientScript.RegisterStartupScript(Me.GetType(), "RecordsModification_Message", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

        End If

    End Sub

    Private Function DatabaseRecords_InsertionProcess(ByRef command As SqlCommand) As String

        DatabaseRecords_InsertionProcess = String.Empty

        Dim initial_number As Integer = Nothing
 
        If Integer.TryParse(tb_shcode_last.Text.Trim(), initial_number) Then

            Try
                command.Connection.Open()
                command.Transaction = command.Connection.BeginTransaction(IsolationLevel.ReadCommitted)

                For shcode_serial_no As Integer = initial_number To ((initial_number + Integer.Parse(tb_records_count.Text.Trim())) - 1)

                    command.Parameters.Item("shcode").Value = tb_shcode_first.Text.Trim() & _
                                                              tb_shcode_middle.Text.Trim() & _
                                                              shcode_serial_no.ToString("0000")

                    command.ExecuteNonQuery()

                Next

                command.Transaction.Commit()

            Catch ex As Exception

                command.Transaction.Rollback()

                DatabaseRecords_InsertionProcess = ex.Message _
                                                     .Replace("'"c, """"c) _
                                                     .Replace(Environment.NewLine, "\n")

            Finally

                If Not IsNothing(command.Transaction) Then
                    command.Transaction.Dispose()
                    command.Transaction = Nothing
                End If

            End Try

        End If

    End Function

    Private Function DatabaseRecord_ModificationProcess(ByRef command As SqlCommand) As String

        command.Parameters.Item("shcode").Value = tb_shcode_first.Text.Trim() & _
                                                          tb_shcode_middle.Text.Trim() & _
                                                          tb_shcode_last.Text.Trim()

        Try
            command.Connection.Open()
            command.ExecuteNonQuery()

            DatabaseRecord_ModificationProcess = String.Empty

        Catch ex As Exception

            DatabaseRecord_ModificationProcess = ex.Message _
                                                   .Replace("'"c, """"c) _
                                                   .Replace(Environment.NewLine, "\n")

        End Try

    End Function

    Protected Sub ibtnPrevious_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnPrevious.Click

        Dim parameters As New SortedList : With parameters
            .Add("Action", "MOA04025.aspx")
        End With

        If Not String.IsNullOrEmpty(Me.ViewState("query_shcode")) Then
            parameters.Add("query_shcode", Me.ViewState.Item("query_shcode"))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("query_company")) Then
            parameters.Add("query_company", Me.ViewState.Item("query_company"))
        End If

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()

    End Sub

    Public Sub Showschcode_lastest()
        Dim iInsertlastCnt As Integer = Integer.Parse(tb_records_count.Text.Trim())
        If Integer.Parse(tb_records_count.Text.Trim()) > 1 Then
            tb_schcode_lastest.Visible = True
            tb_schcode_lastest.Text = "~" + (SHCode_LastNumeric_Generation() + (Integer.Parse(tb_records_count.Text.Trim() - 1))).ToString("0000")
        Else
            tb_schcode_lastest.Visible = False
        End If
    End Sub
End Class
