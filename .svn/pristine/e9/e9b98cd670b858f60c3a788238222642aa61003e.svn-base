'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Create at 2010/06/17
'   Description : First version of development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Finished at 2010/06/21
'   Description : 完成開發並修正 SQL 執行錯誤發生的訊息定義。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/25
'   Description : 修正執行查詢後，GridView 操作異動作業時，呈現錯誤問題。
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Imports WebUtilities.Functions

Partial Class Source_04_MOA04014
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder
    Public FileNamedt As DataTable
    Public sfile_a As String = String.Empty
    Public sfile_b As String = String.Empty
    Public sfile_a_path As String = String.Empty
    Public sfile_b_path As String = String.Empty
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

            If LoginCheck.LoginCheck(user_id, "MOA04013") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04014.aspx")
                Response.End()
            End If

        End If

        If Not Page.IsPostBack Then

            Me.DropDownListControls_Databind()

            If String.IsNullOrEmpty(Request.Form("it_code")) Then

                ibtnModify.Enabled = False
                ibtnModify.Visible = False

                Label1.Text = "設備物料分類資料新增"

            Else

                '物料代碼篩選器第一層之 DropDownList 預設值設定
                Dim it_code_front As String = Left(Request.Form("it_code").PadRight(6, " "c), 3)

                it_code_front = Left(it_code_front, 1).Trim()
                it_code_front = IIf(String.IsNullOrEmpty(it_code_front), String.Empty, it_code_front.PadRight(6, "0"))

                Dim selected_item As ListItem = Nothing

                selected_item = ddl_it_code_L1.Items.FindByValue(it_code_front)
                selected_item = IIf(IsNothing(selected_item), ddl_it_code_L1.Items(0), selected_item)

                selected_item.Selected = True

                Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L1, Nothing)

                '物料代碼篩選器第二層之 DropDownList 預設值設定
                it_code_front = Left(Request.Form("it_code").PadRight(6, " "c), 3)

                it_code_front = Left(it_code_front, 2).Trim()
                it_code_front = IIf(String.IsNullOrEmpty(it_code_front), String.Empty, it_code_front.PadRight(6, "0"))

                selected_item = ddl_it_code_L2.Items.FindByValue(it_code_front)
                selected_item = IIf(IsNothing(selected_item), ddl_it_code_L2.Items(0), selected_item)

                selected_item.Selected = True

                Me.DropDownLists_ItCode_Filter_SelectedIndexChanged(ddl_it_code_L2, Nothing)

                '物料代碼篩選器第三層之 DropDownList 預設值設定
                it_code_front = Left(Request.Form("it_code").PadRight(6, " "c), 3)

                it_code_front = Left(it_code_front, 3).Trim()
                it_code_front = IIf(String.IsNullOrEmpty(it_code_front), String.Empty, it_code_front.PadRight(6, "0"))

                selected_item = ddl_it_code_L3.Items.FindByValue(it_code_front)
                selected_item = IIf(IsNothing(selected_item), ddl_it_code_L3.Items(0), selected_item)

                selected_item.Selected = True

                it_code_front = Nothing
                selected_item = Nothing

                '物料代碼前三碼與後三碼 TextBox 預設值設定
                tb_it_code_front.Text = Left(Request.Form("it_code").PadRight(6, " "c), 3).Trim()
                tb_it_code_back.Text = Right(Request.Form("it_code").PadRight(6, " "c), 3).Trim()

                '其他輸入項的 TextBox 預設值設定
                tb_it_name.Text = Request.Form("it_name")
                tb_it_spec.Text = Request.Form("it_spec")
                tb_it_cost.Text = Request.Form("it_cost")
                tb_manufacturer.Text = Request.Form("manufacturer")
                tb_memo.Text = Request.Form("memo")
                tb_snum.Text = Request.Form("snum")
                tb_expired_y.Text = Request.Form("expired_y")
                tb_it_sort.Text = Request.Form("it_sort")
                tb_it_unit.Text = Request.Form("it_unit")

                FileNamedt = GetfileName(Request.Form("it_code"))
                If Not FileNamedt Is Nothing And FileNamedt.Rows.Count > 0 Then
                    sfile_a = FileNamedt.Rows(0)("file_a").ToString()
                    sfile_b = FileNamedt.Rows(0)("file_b").ToString()
                    If sfile_a.Length > 0 Then
                        img_file_a.Visible = True
                        img_file_a.ImageUrl = sPicPath + Request.Form("it_code") + "A_" + sfile_a
                        sfile_a_path = sPicPath + Request.Form("it_code") + "A_" + sfile_a
                    End If
                    If sfile_b.Length > 0 Then
                        img_file_b.Visible = True
                        img_file_b.ImageUrl = sPicPath + Request.Form("it_code") + "B_" + sfile_b
                        sfile_b_path = sPicPath + Request.Form("it_code") + "B_" + sfile_b
                    End If
                Else

                End If

                ibtnAdd.Enabled = False
                ibtnAdd.Visible = False

                Me.ViewState("it_code_original") = Request.Form("it_code")
                Me.Label1.Text = "設備物料分類修改"

            End If

            Page.Header.Title = Label1.Text

            If Not String.IsNullOrEmpty(Request.Form("query_it_code")) Then
                Me.ViewState("query_it_code") = Request.Form("query_it_code").Trim()
            End If

            If Not String.IsNullOrEmpty(Request.Form("query_it_name")) Then
                Me.ViewState("query_it_name") = Request.Form("query_it_name").Trim()
            End If

            '其他輸入項的 FileUpload 預設值設定
            Me.FileUploadControls_Configuration(mv_file_a)
            Me.FileUploadControls_Configuration(mv_file_b)

        End If

        ImageButtons_ClientScriptsPreparing()

    End Sub

    Private Sub DropDownListControls_Databind()

        If TypeOf sdsItCodeFilter.SelectParameters("it_code") Is Parameter Then

            sdsItCodeFilter.SelectParameters("it_code").DefaultValue = "%".PadRight(6, "0")
            ddl_it_code_L1.DataSource = DirectCast(sdsItCodeFilter.Select(New DataSourceSelectArguments), DataView).Table

        End If

        ddl_it_code_L1.DataBind()
        ddl_it_code_L2.DataBind()
        ddl_it_code_L3.DataBind()

    End Sub

    Private Sub FileUploadControls_Configuration(ByRef multi_viewer As MultiView)

        Dim key As String = multi_viewer.ID.Substring(3)
        Dim file_name As String = IIf(IsNothing(Request.Form(key)), String.Empty, Request.Form(key)).ToString().Trim()

        If String.IsNullOrEmpty(file_name) Then
            multi_viewer.ActiveViewIndex = 0
        Else

            multi_viewer.ActiveViewIndex = 1

            DirectCast(Me.FindControl("lkb_" & key), LinkButton).Text = file_name
            DirectCast(Me.FindControl("btn_del_" & key), Button).CommandArgument = file_name

        End If

        file_name = Nothing

    End Sub

    Private Sub ImageButtons_ClientScriptsPreparing()
        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_it_code_front.ClientID) _
               .Append("blank_validation: [{ message: '請輸入物料代碼前三碼' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_it_code_back.ClientID) _
               .Append("blank_validation: [{ message: '請輸入物料代碼後三碼' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_it_name.ClientID) _
               .Append("blank_validation: [{ message: '請輸入物料名稱' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_snum.ClientID) _
               .Append("blank_validation: [{ message: '請輸入安全數量' }], ") _
               .Append("numeric_validation: [{ message: '安全數量請輸入數字' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_expired_y.ClientID) _
               .Append("numeric_validation: [{ message: '有效期請輸入數字' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_it_name.ClientID) _
               .Append("RequestFormError_validation: [{ message: '物料名稱請勿輸入非法字元' }] }, ") _
                .AppendFormat("{{ client_id: '{0}', ", tb_it_spec.ClientID) _
                .Append("RequestFormError_validation: [{ message: '物料規格請勿輸入非法字元' }] }, ") _
                .AppendFormat("{{ client_id: '{0}', ", tb_it_sort.ClientID) _
                .Append("RequestFormError_validation: [{ message: '物料類別請勿輸入非法字元' }] }, ") _
                .AppendFormat("{{ client_id: '{0}', ", tb_it_unit.ClientID) _
                .Append("RequestFormError_validation: [{ message: '物料單位請勿輸入非法字元' }] }, ") _
                .AppendFormat("{{ client_id: '{0}', ", tb_memo.ClientID) _
                .Append("RequestFormError_validation: [{ message: '備註請勿輸入非法字元' }] } ") _
                .Append("]});")

        ibtnAdd.Attributes.Add("OnClick", scripts.ToString())
        ibtnModify.Attributes.Add("OnClick", scripts.ToString())

        scripts.Remove(0, scripts.Length)
    End Sub

    Protected Sub DropDownLists_ItCode_Filter_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_it_code_L1.DataBound, ddl_it_code_L2.DataBound, ddl_it_code_L3.DataBound

        DirectCast(sender, DropDownList).Items.Insert(0, New ListItem("請選擇", String.Empty))

    End Sub

    Protected Sub DropDownLists_ItCode_Filter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_it_code_L1.SelectedIndexChanged, ddl_it_code_L2.SelectedIndexChanged, ddl_it_code_L3.SelectedIndexChanged

        Dim ddlControl As DropDownList = DirectCast(sender, DropDownList)

        If TypeOf ddlControl Is DropDownList Then

            Dim length As Integer = Nothing
            Dim filter_expression As String = "".PadRight(0, " "c)

            If Integer.TryParse(ddlControl.ClientID(ddlControl.ClientID.Length - 1).ToString(), length) Then

                If Not String.IsNullOrEmpty(ddlControl.SelectedValue) Then

                    filter_expression = ddlControl.SelectedValue.Substring(0, length)
                    filter_expression += "%"
                    filter_expression = filter_expression.PadRight(6, "0")

                End If

            End If

            sdsItCodeFilter.SelectParameters("it_code").DefaultValue = filter_expression

            If TypeOf sdsItCodeFilter.SelectParameters.Item("it_code_ignore") Is Parameter Then
                sdsItCodeFilter.SelectParameters.Remove(sdsItCodeFilter.SelectParameters("it_code_ignore"))
            End If

            Select Case ddlControl.ClientID

                Case "ddl_it_code_L1", _
                     "ddl_it_code_L2"

                    sdsItCodeFilter.SelectCommand = sdsItCodeFilter.SelectCommand _
                                                                   .Replace("(1 = 1)", "it_code <> @it_code_ignore")

                    sdsItCodeFilter.SelectParameters.Add(New Parameter("it_code_ignore", DbType.String, filter_expression.Replace("%"c, "0"c)))

                    Me.tb_it_code_front.Text = String.Empty

            End Select

            Select Case ddlControl.ClientID

                Case "ddl_it_code_L1"

                    ddl_it_code_L2.Items.Clear()
                    ddl_it_code_L3.Items.Clear()

                    ddl_it_code_L2.DataSource = sdsItCodeFilter.Select(New DataSourceSelectArguments)
                    ddl_it_code_L2.DataBind()
                    ddl_it_code_L3.DataBind()

                Case "ddl_it_code_L2"

                    ddl_it_code_L3.Items.Clear()

                    ddl_it_code_L3.DataSource = sdsItCodeFilter.Select(New DataSourceSelectArguments)
                    ddl_it_code_L3.DataBind()

                Case "ddl_it_code_L3"

                    Me.tb_it_code_front.Text = ddlControl.SelectedValue.PadRight(6, " "c).Substring(0, 3).Trim()

            End Select

        End If

    End Sub

    Protected Sub ImageButtons_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click, ibtnModify.Click

        If Me.UploadFileModification_Processes(mv_file_a) And Me.UploadFileModification_Processes(mv_file_b) Then

            Dim sql_function As New C_SQLFUN
            Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

            command.CommandType = Data.CommandType.Text

            Select Case (CType(sender, ImageButton).CommandName)

                Case "Add"
                    command.CommandText = "Insert Into P_0407(it_code, it_name, it_spec, it_cost, manufacturer, memo, " & _
                                          "file_a, file_b, snum, expired_y, it_sort, it_unit, operator) Values(@it_code, " & _
                                          "@it_name, @it_spec, @it_cost, @manufacturer, @memo, @file_a, @file_b, @snum, " & _
                                          "@expired_y, @it_sort, @it_unit, @operator)"

                    command.Parameters.Add(New SqlParameter("operator", SqlDbType.VarChar, 10)).Value = user_id

                Case "Modify"
                    command.CommandText = "Update P_0407 Set it_code = @it_code, it_name = @it_name, it_spec = @it_spec, " & _
                                          "it_cost = @it_cost, manufacturer = @manufacturer, memo = @memo, file_a = @file_a, " & _
                                          "file_b = @file_b, snum = @snum, expired_y = @expired_y, it_sort = @it_sort, " & _
                                          "it_unit = @it_unit Where it_code = @it_code_original;"

                    command.Parameters.Add(New SqlParameter("it_code_original", SqlDbType.VarChar, 6)).Value = Me.ViewState("it_code_original")

            End Select

            command.Parameters.Add(New SqlParameter("it_name", SqlDbType.NVarChar, 255)).Value = tb_it_name.Text.Trim()
            command.Parameters.Add(New SqlParameter("it_spec", SqlDbType.NVarChar, 255)).Value = tb_it_spec.Text.Trim()
            command.Parameters.Add(New SqlParameter("it_cost", SqlDbType.Int)).Value = tb_it_cost.Text.Trim()
            command.Parameters.Add(New SqlParameter("manufacturer", SqlDbType.NVarChar, 255)).Value = tb_manufacturer.Text.Trim()
            command.Parameters.Add(New SqlParameter("memo", SqlDbType.NVarChar, 255)).Value = tb_memo.Text.Trim()
            command.Parameters.Add(New SqlParameter("snum", SqlDbType.Int)).Value = tb_snum.Text.Trim()
            command.Parameters.Add(New SqlParameter("expired_y", SqlDbType.Int)).Value = tb_expired_y.Text.Trim()
            command.Parameters.Add(New SqlParameter("it_sort", SqlDbType.NVarChar, 10)).Value = tb_it_sort.Text.Trim()
            command.Parameters.Add(New SqlParameter("it_unit", SqlDbType.NVarChar, 10)).Value = tb_it_unit.Text.Trim()

            command.Parameters.Add(New SqlParameter("it_code", SqlDbType.VarChar, 6)).Value = tb_it_code_front.Text.Trim().PadRight(3, "0"c) & _
                                                                                              tb_it_code_back.Text.Trim().PadRight(3, "0"c)

            If Not Me.ViewState("it_code_original") Is Nothing Then
                FileNamedt = GetfileName(Me.ViewState("it_code_original"))
                If Not FileNamedt Is Nothing And FileNamedt.Rows.Count > 0 Then
                    sfile_a = FileNamedt.Rows(0)("file_a").ToString()
                    sfile_b = FileNamedt.Rows(0)("file_b").ToString()

                    If sfile_a <> fu_file_a.FileName.Trim() And fu_file_a.FileName.Trim() <> "" Then
                        sfile_a = fu_file_a.FileName.Trim()
                    End If
                    If sfile_b <> fu_file_b.FileName.Trim() And fu_file_b.FileName.Trim() <> "" Then
                        sfile_b = fu_file_b.FileName.Trim()
                    End If
                End If
            End If

            
            command.Parameters.Add(New SqlParameter("file_a", SqlDbType.VarChar, 50)).Value = sfile_a
            command.Parameters.Add(New SqlParameter("file_b", SqlDbType.VarChar, 50)).Value = sfile_b

            For Each element As SqlParameter In command.Parameters
                If String.IsNullOrEmpty(element.Value) Then
                    element.Value = DBNull.Value
                End If
            Next

            'Database Modification Process
            Try
                command.Connection.Open()
                command.ExecuteNonQuery()

                ibtnPrevious_Click(Me.ibtnPrevious, Nothing)

            Catch ex As Exception

                If (TypeOf ex Is SqlException) Then

                    Select Case (CType(ex, SqlException).Number)

                        Case 2627
                            scripts.Append("該筆資料已存在")

                        Case 8152
                            scripts.Append("欄位不允許輸入中文")

                        Case Else
                            scripts.AppendFormat("{0}: ", CType(ex, SqlException).Number) _
                                   .Append(ex.Message.Replace(Environment.NewLine, String.Empty))

                    End Select
                End If

                If (scripts.Length.Equals(0)) Then
                    scripts.Append(ex.Message.Replace(Environment.NewLine, String.Empty))
                End If

                scripts.Insert(0, "alert('").Append(" !');")

                ClientScript.RegisterStartupScript(Me.GetType(), "ErrorMessage", scripts.ToString(), True)
                scripts.Remove(0, scripts.Length)

            Finally

                If command.Connection.State.Equals(ConnectionState.Open) Then
                    command.Connection.Close()
                End If

            End Try

        End If

        fu_file_a.PostedFile.InputStream.Close()
        fu_file_b.PostedFile.InputStream.Close()

    End Sub

    Private Function UploadFileModification_Processes(ByRef multi_viewer As MultiView) As Boolean

        Dim file_path As String = String.Format _
                                         ( _
                                              "{0}\{{0}}{{1}}{1}_{{2}}", _
                                                Me.Server.MapPath("~/M_Source/99"), _
                                                Right(multi_viewer.ID, 1).ToUpper() _
                                         )

        Select Case multi_viewer.ActiveViewIndex

            Case 0

                Dim file_uploader As FileUpload = multi_viewer.Views(multi_viewer.ActiveViewIndex).FindControl(multi_viewer.ID.Substring(3).Insert(0, "fu_"))

                If Not String.IsNullOrEmpty(file_uploader.FileName.Trim()) Then
                    Dim FileExtension As String = Path.GetExtension(file_uploader.FileName)
                    If FileExtension = ".jpg" Or FileExtension = ".JPG" Or FileExtension = ".JPEG" Or FileExtension = ".GIF" Or FileExtension = ".gif" Then
                        Using file_writer As FileStream = File.OpenWrite _
                                                               ( _
                                                                         String.Format _
                                                                                ( _
                                                                                       file_path, _
                                                                                       tb_it_code_front.Text.Trim(), _
                                                                                       tb_it_code_back.Text.Trim(), _
                                                                                       file_uploader.FileName.Trim() _
                                                                                ) _
                                                               )

                            Dim file_buffer(1024) As Byte
                            Dim read_length As Integer = 0

                            Try
                                Do
                                    read_length = file_uploader.PostedFile.InputStream.Read(file_buffer, 0, file_buffer.Length)
                                    file_writer.Write(file_buffer, 0, read_length)

                                Loop Until read_length.Equals(0)

                            Catch ex As Exception

                                scripts.Append("alert('檔案上傳處理失敗 !!\t');")

                                ClientScript.RegisterStartupScript(Me.GetType(), "FileUpload_ProcessesError", scripts.ToString(), True)
                                scripts.Remove(0, scripts.Length)

                                UploadFileModification_Processes = False

                            Finally

                                file_writer.Close()

                                Array.Clear(file_buffer, 0, file_buffer.Length)
                                file_buffer = Nothing

                            End Try

                        End Using
                    Else
                        scripts.Append("alert('上傳圖檔錯誤，請上傳.jpg或.gif副檔名之圖片檔案 !!\t');")

                        ClientScript.RegisterStartupScript(Me.GetType(), "FileUpload_ProcessesError", scripts.ToString(), True)
                        scripts.Remove(0, scripts.Length)

                        Return False
                    End If
                End If

            Case 1

                    Dim it_code_original As String = _
                        IIf(IsNothing(Me.ViewState("it_code_original")), String.Empty, Me.ViewState("it_code_original")).ToString().Trim().PadLeft(6, " "c)

                    If Not (tb_it_code_front.Text & tb_it_code_back.Text).Trim().Equals(it_code_original) Then

                        Dim current_file_name As String = DirectCast(multi_viewer.Views(multi_viewer.ActiveViewIndex). _
                                                          FindControl(multi_viewer.ID.Substring(3).Insert(0, "lkb_")), LinkButton).Text.Trim()
                        Try
                            File.Copy(String.Format(file_path, Left(it_code_original, 3), Right(it_code_original, 3), current_file_name), _
                                      String.Format(file_path, tb_it_code_front.Text.Trim(), tb_it_code_back.Text.Trim(), current_file_name))

                            File.Delete(String.Format(file_path, Left(it_code_original, 3), Right(it_code_original, 3), current_file_name))

                        Catch ex As Exception

                            scripts.Append("alert('檔案異動處理失敗 !!\t');")

                            ClientScript.RegisterStartupScript(Me.GetType(), "FileModification_ProcessesError", scripts.ToString(), True)
                            scripts.Remove(0, scripts.Length)

                            UploadFileModification_Processes = False

                        End Try

                        current_file_name = Nothing

                    End If

                    it_code_original = Nothing

        End Select

        UploadFileModification_Processes = True

    End Function

    Protected Sub FileUploadDeletionButtons_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_del_file_a.Click, btn_del_file_b.Click

        Dim it_code_original As String = _
            IIf(IsNothing(Me.ViewState("it_code_original")), String.Empty, Me.ViewState("it_code_original")).ToString().Trim().PadLeft(6, " "c)

        Dim file_path As String = String.Format _
                                         ( _
                                              "{0}\{1}{2}{3}_{4}", _
                                                Me.Server.MapPath("~/M_Source/99"), _
                                                Left(it_code_original, 3), _
                                                Right(it_code_original, 3), _
                                                Right(DirectCast(sender, Button).ID, 1).ToUpper(), _
                                                DirectCast(sender, Button).CommandArgument _
                                         )

        it_code_original = Nothing

        If File.Exists(file_path) Then

            Dim has_deleted As Boolean = False

            Try
                File.Delete(file_path)
                has_deleted = True

            Catch ex As Exception

                scripts.Append("alert('檔案刪除失敗 !!\t');")

                ClientScript.RegisterStartupScript(Me.GetType(), "FileUploadDeletionError", scripts.ToString(), True)
                scripts.Remove(0, scripts.Length)

            End Try

            If has_deleted Then

                Dim command As New SqlCommand _
                                   ( _
                                            String.Empty, _
                                            New SqlConnection _
                                                ( _
                                                            ConfigurationManager.ConnectionStrings("ConnectionString") _
                                                                                .ConnectionString _
                                                ) _
                                   )

                command.CommandText = "Update P_0407 Set {0} = Null Where it_code = @it_code_original;"
                command.CommandText = String.Format(command.CommandText, DirectCast(sender, Button).ID.Substring(8))

                command.Parameters.Add(New SqlParameter("it_code_original", SqlDbType.VarChar, 6)).Value = _
                        IIf(IsNothing(Me.ViewState("it_code_original")), String.Empty, Me.ViewState("it_code_original")).ToString().Trim().PadLeft(6, " "c)

                Try
                    command.Connection.Open()
                    command.ExecuteNonQuery()

                    has_deleted = True

                Catch ex As Exception

                    scripts.Append(ex.Message.Replace("'"c, """"c).Replace(Environment.NewLine, "\n"))

                    ClientScript.RegisterStartupScript(Me.GetType(), "FileUploadDeletionDbError", scripts.ToString(), True)
                    scripts.Remove(0, scripts.Length)

                    has_deleted = False

                Finally

                    If command.Connection.State.Equals(ConnectionState.Open) Then
                        command.Connection.Close()
                    End If

                    command.Dispose()
                    command = Nothing

                End Try

                If has_deleted Then
                    DirectCast(DirectCast(sender, Button).Parent.Parent, MultiView).ActiveViewIndex = 0
                End If

            End If

        End If

    End Sub

    Protected Sub ibtnPrevious_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnPrevious.Click
        scripts = Nothing
        Dim parameters As New SortedList()

        parameters.Add("Action", "MOA04013.aspx")

        If Not String.IsNullOrEmpty(Me.ViewState("query_it_code")) Then
            parameters.Add("query_it_code", Me.ViewState("query_it_code"))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("query_it_name")) Then
            parameters.Add("query_it_name", Me.ViewState("query_it_name"))
        End If

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()

    End Sub

    Private Function GetfileName(ByVal it_code As String) As DataTable
        GetfileName = New DataTable("filenameList")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        scripts.Append("select isnull(file_a,'') as file_a,isnull(file_b,'') as file_b from P_0407 with(nolock) where it_code=@it_code")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("it_code", SqlDbType.NVarChar, 6)).Value = it_code
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            GetfileName.Load(dr)

        Catch ex As Exception
            GetfileName = Nothing

        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function
End Class