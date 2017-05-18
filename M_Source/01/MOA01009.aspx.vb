'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Created at 2010/07/12
'   Description : First Version for Development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/07/16
'   Description : 修改表單資料擷取方式 (以 Data Server 處理大部分運算，而不再依靠 Web Server 處理)。
'==================================================================================================================================================================================
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.Serialization.Formatters.Binary

Imports WebUtilities.Pages

Partial Class Source_04_MOA01009
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder

    Protected ReadOnly Property Over_Time_Baseline() As TimeSpan

        Get

            Dim time_span As TimeSpan = Nothing

            If Not TimeSpan.TryParse(ddl_overtime.SelectedValue.Trim(), time_span) Then
                time_span = TimeSpan.Parse("17:00:00")
            End If

            Return time_span

        End Get

    End Property

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

            If LoginCheck.LoginCheck(user_id, "MOA01009") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA01009.aspx")
                Response.End()
            End If

        End If

        If Not Me.IsPostBack Then

            Me.PageNonPostBack_Preparation()

            If IsNothing(Me.ViewState.Item("Normal_User")) Then
                Me.ImgSearch_Click(Me.ImgSearch, Nothing)
            End If

        End If

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

    Private Sub PageNonPostBack_Preparation()

        Dim organization As New C_Public
        Dim sql_statement As String = sdsOrganization.SelectCommand.Replace("(1 = 2)", String.Empty)

        If Not IsNothing(Session("Role")) Then

            If Session.Item("Role").Equals("1") Then

                '系統管理員
                sdsOrganization.SelectCommand = sql_statement & "1 = 1"

            Else

                If Me.IsAuthorizedAdministrator(MOA01008.UnitSalariesAdministrator) Then

                    '單位薪俸管理員
                    sdsOrganization.SelectCommand = sql_statement & _
                                                    String.Format("ORG_UID In ({0})", organization.getchildorg(organization.getUporg(org_uid, 1)))

                Else

                    '一般使用者
                    sdsOrganization.SelectCommand = sql_statement & _
                                                    String.Format("ORG_UID In ({0})", organization.getchildorg(org_uid))

                    Me.ViewState.Add("Normal_User", String.Empty)

                End If

            End If

        End If

        organization = Nothing

        sdsOrganization.SelectCommand += " Order by ORG_NAME"

        '判斷是否開啟中午加班簽入/簽出功能
        sdsUsersList.SelectParameters.Item("ORG_UID").DefaultValue = org_uid

        sdsUsersList.SelectCommand = sdsUsersList.SelectCommand _
                                                 .Replace _
                                                  ( _
                                                        "(1 = 1)", _
                                                          String.Format _
                                                                 ( _
                                                                       "employee_id = '{0}'", _
                                                                        user_id _
                                                                 ) _
                                                  )

        Dim result As DataTable = DirectCast(sdsUsersList.Select(New DataSourceSelectArguments), DataView).Table

        If IsNothing(result) Or result.Rows.Count.Equals(0) Then

            btnNoonSignIn.Enabled = False
            btnNoonSignIn.Visible = False

            btnNoonSignOut.Enabled = False
            btnNoonSignOut.Visible = False

        End If

        If Not IsNothing(result) Then

            result.Clear()
            result.Dispose()

            result = Nothing

        End If

        sdsUsersList.SelectParameters.Item("ORG_UID").DefaultValue = String.Empty

        '設定日期條件預設值
        Dim current_date As Date = DateTime.Today

        If tb_begin_date.Text.Trim().Equals(String.Empty) Then
            tb_begin_date.Text = current_date.AddDays(-7).Date
        End If

        If tb_end_date.Text.Trim().Equals(String.Empty) Then
            tb_end_date.Text = current_date.Date
        End If

        current_date = Nothing

        'DropDownList Controllers Preparation
        ddl_overtime.DataBind()
        ddl_sign_in.DataBind()
        ddl_sign_out.DataBind()

    End Sub

    Protected Sub DropDownListControllers_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles ddl_org_picker.DataBound, ddl_user_picker.DataBound, ddl_overtime.DataBound, _
                    ddl_sign_in.DataBound, ddl_sign_out.DataBound

        CType(sender, DropDownList).Items.Insert(0, New ListItem("請選擇", String.Empty))

        If Not IsNothing(Me.ViewState.Item("Normal_User")) Then

            Dim ddlControl As DropDownList = DirectCast(sender, DropDownList)
            Dim selected_item As ListItem = ddlControl.Items(0)

            Select Case ddlControl.ID

                Case ddl_org_picker.ID
                    selected_item = ddlControl.Items.FindByValue(org_uid)

                Case ddl_user_picker.ID
                    selected_item = ddlControl.Items.FindByValue(user_id)

            End Select

            selected_item = IIf(IsNothing(selected_item), ddlControl.Items(0), selected_item)

            selected_item.Selected = True
            selected_item = Nothing

            If Not Me.IsPostBack Then

                If ddlControl.ID.Equals(ddl_user_picker.ID) Then
                    Me.ImgSearch_Click(Me.ImgSearch, Nothing)
                End If

            End If

            ddlControl = Nothing

        End If

    End Sub

    Protected Sub ddl_org_picker_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_org_picker.SelectedIndexChanged
        ddl_user_picker.Items.Clear()
    End Sub

    Protected Sub sdsUsersList_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles sdsUsersList.Selecting

        If Not IsNothing(Me.ViewState.Item("Normal_User")) Then

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

    End Sub

    Protected Sub ibtn_begin_date_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtn_begin_date.Click

        div_grid1.Visible = True

        div_grid1.Style("Top") = "90px"
        div_grid1.Style("Left") = "58px"

        Calendar1.SelectedDate = tb_begin_date.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        tb_begin_date.Text = Calendar1.SelectedDate.Date
        div_grid1.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click
        div_grid1.Visible = False
    End Sub

    Protected Sub ibtn_end_date_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtn_end_date.Click

        div_grid2.Visible = True

        div_grid2.Style("Top") = "90px"
        div_grid2.Style("Left") = "218px"

        Calendar2.SelectedDate = tb_end_date.Text

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        tb_end_date.Text = Calendar2.SelectedDate.Date
        div_grid2.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click
        div_grid2.Visible = False
    End Sub

    Protected Sub SqlDataSourceControllers_Init(ByVal sender As Object, ByVal e As System.EventArgs) _
            Handles sdsUsersList.Init, sdsSignRecords.Init

        CType(sender, SqlDataSource).SelectParameters("AD_TITLE").DefaultValue = MOA01009.Special_AD_Title

    End Sub

    Protected Sub SqlDataSourceControllers_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) _
            Handles sdsSignRecords.Selecting

        e.Command.Parameters("@employee_id").Value = IIf(IsNothing(Me.ViewState(ddl_user_picker.ID)), String.Empty, Me.ViewState.Item(ddl_user_picker.ID))

        If String.IsNullOrEmpty(e.Command.Parameters("@employee_id").Value) Then
            e.Command.Parameters("@employee_id").Value = DBNull.Value
        End If

        e.Command.Parameters("@BEGIN_DATE").Value = Me.tb_begin_date.Text
        e.Command.Parameters("@END_DATE").Value = Me.tb_end_date.Text

        For Each parameter As System.Data.Common.DbParameter In e.Command.Parameters

            sdsSignRecords.SelectParameters _
                          .Item(parameter.ParameterName.TrimStart("@"c)) _
                          .DefaultValue = _
                           IIf(parameter.Value.Equals(DBNull.Value), String.Empty, parameter.Value.ToString())

        Next

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then

            If e.Row.RowState.Equals(DataControlRowState.Normal) Or e.Row.RowState.Equals(DataControlRowState.Alternate) Then

                Dim result As List(Of Integer) = Nothing
                Dim drv_record As DataRowView = CType(e.Row.DataItem, DataRowView)

                MOA01009.HoursAndSubTotal_Calculations _
                         ( _
                                 drv_record("Out_Time_nvc"), _
                                 drv_record("HourLimit"), _
                                 drv_record("OverTime"), _
                                 drv_record("NoonHour"), _
                                 Me.Over_Time_Baseline, _
                                 result _
                         )

                If Not result.Item(0).Equals(Integer.MinValue) Then
                    CType(e.Row.Cells(4).FindControl("lbl_hours_difference"), Label).Text = result.Item(0).ToString()
                End If

                If Not result.Item(1).Equals(Integer.MinValue) Then
                    CType(e.Row.Cells(5).FindControl("lbl_subtotal"), Label).Text = result.Item(1).ToString("$#,###")
                End If

                result.Clear()
                result = Nothing

                drv_record = Nothing

            End If

        End If

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging

        Me.GridView1.PageIndex = e.NewPageIndex

        Me.SearchSqlParameters_Preparation()
        Me.SearchSqlStatementWhereConditions_Preparation()

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Me.SearchSqlParameters_Preparation()
        Me.SearchSqlStatementWhereConditions_Preparation()

    End Sub

    Protected Sub btnNoonSignIn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNoonSignIn.Click

        Dim current_date_time As DateTime = DateTime.Now

        If current_date_time.CompareTo(current_date_time.Date.Add(TimeSpan.Parse("11:00:00"))) > -1 _
           And _
           current_date_time.CompareTo(current_date_time.Date.Add(TimeSpan.Parse("13:00:00"))) < 1 Then

            If Me.IsOverTimeAtNoon_SignedIn(current_date_time.Date) Then
                scripts.Append("今日已完成中午加班簽到手續 !!\n")
            Else

                Dim command As New SqlCommand(String.Empty, New SqlConnection(New C_SQLFUN().G_conn_string))

                scripts.Append("Insert Into P_0104 (employee_id, Noon, NoonIn) ") _
                       .Append("Values (@employee_id, @Noon, @NoonIn)")

                command.CommandType = CommandType.Text
                command.CommandText = scripts.ToString()

                scripts.Remove(0, scripts.Length)

                command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = user_id
                command.Parameters.Add(New SqlParameter("Noon", SqlDbType.DateTime)).Value = current_date_time.Date
                command.Parameters.Add(New SqlParameter("NoonIn", SqlDbType.DateTime)).Value = current_date_time

                Try
                    command.Connection.Open()
                    command.ExecuteNonQuery()

                    scripts.Append("中午加班簽到手續完成 !!\n") _
                           .AppendFormat("您的簽到時間為：{0}。\n", current_date_time.ToString("yyyy/M/d HH:mm:ss"))

                Catch ex As Exception

                    scripts.Append("系統執行錯誤 !!\n") _
                           .Append("中午加班簽到手續未完成 !!\n")

                Finally

                    If command.Connection.State.Equals(ConnectionState.Open) Then
                        command.Connection.Close()
                    End If

                    command.Dispose()
                    command = Nothing

                End Try

            End If

        Else

            scripts.Append("注意：中午加班簽到時段為 11:00 ~ 13:00  !!\n") _
                   .Append("因未在中午加班時段，故無法完成簽到手續 !!\n")

        End If

        If Not scripts.Length.Equals(0) Then

            scripts.Insert(0, "function ShowNoonSignIn_Message() { alert('") _
                   .Append("'); }")

            ClientScript.RegisterClientScriptBlock(Me.GetType(), "NoonSignIn_Message", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

            ClientScript.RegisterStartupScript(Me.GetType(), "NoonSignIn_Display", "ShowNoonSignIn_Message();", True)

        End If

    End Sub

    Protected Sub btnNoonSignOut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNoonSignOut.Click

        Dim current_date_time As DateTime = DateTime.Now

        If Me.IsOverTimeAtNoon_SignedIn(current_date_time.Date) Then

            Dim command As New SqlCommand(String.Empty, New SqlConnection(New C_SQLFUN().G_conn_string))

            scripts.Append("Update P_0104 Set NoonOut = @NoonOut Where employee_id = @employee_id ") _
                   .Append("And Noon = @Noon")

            command.CommandType = CommandType.Text
            command.CommandText = scripts.ToString()

            scripts.Remove(0, scripts.Length)

            command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = user_id
            command.Parameters.Add(New SqlParameter("Noon", SqlDbType.DateTime)).Value = current_date_time.Date
            command.Parameters.Add(New SqlParameter("NoonOut", SqlDbType.DateTime)).Value = current_date_time

            Try
                command.Connection.Open()
                command.ExecuteNonQuery()

                scripts.Append("中午加班簽退手續完成 !!\n") _
                       .AppendFormat("您的簽退時間為：{0}。\n", current_date_time.ToString("yyyy/M/d HH:mm:ss"))

            Catch ex As Exception

                scripts.Append("系統執行錯誤 !!\n") _
                       .Append("中午加班簽退手續未完成 !!\n")

            Finally

                If command.Connection.State.Equals(ConnectionState.Open) Then
                    command.Connection.Close()
                End If

                command.Dispose()
                command = Nothing

            End Try

        Else

            scripts.Append("注意：未完成中午加班簽到手續 !!\n") _
                   .Append("因此無法執行中午加班簽退手續 !!\n")

        End If

        If Not scripts.Length.Equals(0) Then

            scripts.Insert(0, "function ShowNoonSignOut_Message() { alert('") _
                   .Append("'); }")

            ClientScript.RegisterClientScriptBlock(Me.GetType(), "NoonSignOut_Message", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

            ClientScript.RegisterStartupScript(Me.GetType(), "NoonSignOut_Display", "ShowNoonSignOut_Message();", True)

        End If

    End Sub

    Private Function IsOverTimeAtNoon_SignedIn(ByRef noon_date As DateTime) As Boolean

        Dim command As New SqlCommand(String.Empty, New SqlConnection(New C_SQLFUN().G_conn_string))

        scripts.Append("Select Count(*) From P_0104 Where employee_id = @employee_id ") _
               .Append("And Noon = @Noon")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()

        scripts.Remove(0, scripts.Length)

        command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = user_id
        command.Parameters.Add(New SqlParameter("Noon", SqlDbType.DateTime)).Value = noon_date

        Try
            command.Connection.Open()

            Return (Not command.ExecuteScalar().Equals(0))

        Catch ex As Exception

            Throw ex

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

            command.Dispose()
            command = Nothing

        End Try

    End Function

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Me.ViewState(ddl_org_picker.ID) = ddl_org_picker.SelectedValue.Trim()
        Me.ViewState(ddl_user_picker.ID) = ddl_user_picker.SelectedValue.Trim()

        If Not IsNothing(Me.ViewState.Item("Normal_User")) Then
            Me.ViewState(ddl_org_picker.ID) = IIf(String.IsNullOrEmpty(Me.ViewState(ddl_org_picker.ID)), org_uid.Trim(), Me.ViewState(ddl_org_picker.ID))
            Me.ViewState(ddl_user_picker.ID) = IIf(String.IsNullOrEmpty(Me.ViewState(ddl_user_picker.ID)), user_id.Trim(), Me.ViewState(ddl_user_picker.ID))
        End If

        Me.SearchSqlParameters_Preparation()
        Me.SearchSqlStatementWhereConditions_Preparation()

    End Sub

    Private Sub SearchSqlParameters_Preparation()

        Dim param_organization As String = IIf(IsNothing(Me.ViewState(ddl_org_picker.ID)), String.Empty, Me.ViewState.Item(ddl_org_picker.ID))

        param_organization = String.Format("'{0}'", param_organization)

        If String.IsNullOrEmpty(param_organization) Then

            Dim c_public As C_Public = Nothing

            If Not IsNothing(Session("Role")) Then

                If Me.Session.Item("Role").Equals("1") Then

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

                    End If

                End If

            End If

            c_public = Nothing

        End If

        sdsSignRecords.SelectCommand = sdsSignRecords.SelectCommand _
                                                     .Replace _
                                                      ( _
                                                             "ORG_UID In ('')", _
                                                              String.Format _
                                                                     ( _
                                                                           "ORG_UID In ({0})", _
                                                                            param_organization _
                                                                     ) _
                                                      )

    End Sub

    Private Sub SearchSqlStatementWhereConditions_Preparation()

        If TypeOf sdsSignRecords.SelectParameters.Item("OT_NVC_BEGIN") Is Parameter Then
            sdsSignRecords.SelectParameters.Remove(sdsSignRecords.SelectParameters.Item("OT_NVC_BEGIN"))
        End If

        If TypeOf sdsSignRecords.SelectParameters.Item("OT_NVC_END") Is Parameter Then
            sdsSignRecords.SelectParameters.Remove(sdsSignRecords.SelectParameters.Item("OT_NVC_END"))
        End If

        Dim where_conditions As String = String.Empty

        If Not String.IsNullOrEmpty(ddl_overtime.SelectedValue) Then

            where_conditions += " And (Out_Time_nvc >= @OT_NVC_BEGIN Or Out_Time_nvc <= @OT_NVC_END)"

            sdsSignRecords.SelectParameters.Add _
                                            ( _
                                                    New Parameter _
                                                        ( _
                                                                 "OT_NVC_BEGIN", DbType.String, _
                                                                  TimeSpan.Parse(ddl_overtime.SelectedValue) _
                                                                          .Add(TimeSpan.FromHours(1)) _
                                                                          .ToString() _
                                                        ) _
                                            )

            sdsSignRecords.SelectParameters.Add(New Parameter("OT_NVC_END", DbType.String, MOA01009.Out_Time_NVC_End))

        End If

        Select Case ddl_sign_in.SelectedValue

            Case "已簽到"
                where_conditions += " And In_Time_nvc Is Not Null"

            Case "未簽到"
                where_conditions += " And In_Time_nvc Is Null"

        End Select

        Select Case ddl_sign_out.SelectedValue

            Case "已簽退"
                where_conditions += " And Out_Time_nvc Is Not Null"

            Case "未簽退"
                where_conditions += " And Out_Time_nvc Is Null"

        End Select

        sdsSignRecords.SelectCommand = sdsSignRecords.SelectCommand.Replace("(1 = 1)", String.Format("(1 = 1){0}", where_conditions))

    End Sub

    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgExport.Click

        Dim data_source As DataTable = Me.DataSourceTable_Generation()

        If TypeOf data_source Is DataTable Then

            Me.DownloadExportedFile_DirectedResponse(data_source)

            data_source.Dispose()
            data_source = Nothing

            Response.End()

        End If

    End Sub

    Private Sub DownloadExportedFile_DirectedResponse(ByRef data_source As DataTable)

        Me.Response.Clear()
        Me.Response.Buffer = True
        Me.Response.Charset = "BIG5"
        Me.Response.ContentType = "application/vnd.ms-excel"
        Me.Response.ContentEncoding = Encoding.GetEncoding(Me.Response.Charset)

        Me.Response.AddHeader _
                    ( _
                             "content-disposition", _
                              String.Format("attachment; filename={0:yyyyMMdd}_01009.csv", Date.Today) _
                    )

        Dim begin_date As Date = Date.Parse(Date.Parse(tb_begin_date.Text).ToString("yyyy/MM/01"))
        Dim end_date As Date = Date.Parse(Date.Parse(tb_end_date.Text).ToString("yyyy/MM/01"))

        Response.Write((begin_date.Year - 1911).ToString("# 年 ") & begin_date.Month.ToString("# 月"))

        If Not begin_date.Subtract(end_date).TotalDays.Equals(0) Then
            Response.Write((end_date.Year - 1911).ToString(" 到 # 年 ") & end_date.Month.ToString("# 月"))
        End If

        Response.Write(" 文職人員加班統計表")
        Response.Write(Environment.NewLine)
        Response.Flush()

        Dim headers(Me.GridView1.Columns.Count - 1) As DataControlField

        Me.GridView1.Columns.CopyTo(headers, 0)

        Response.Write(String.Join(",", Array.ConvertAll(headers, New Converter(Of DataControlField, String)(AddressOf Me.ObjectToString_Converter))))
        Response.Write(Environment.NewLine)
        Response.Flush()

        Array.Clear(headers, 0, headers.Length)
        Array.Resize(Of DataControlField)(headers, 0)

        headers = Nothing

        For Each data_row As DataRow In data_source.Rows

            scripts.Append(String.Join(",", Array.ConvertAll(data_row.ItemArray, New Converter(Of Object, String)(AddressOf Me.ObjectToString_Converter)))) _
                   .Replace(data_row.Item("SIGNDATE_d").ToString(), CType(data_row.Item("SIGNDATE_d"), DateTime).ToString("yyyy/M/d"))

            If Not data_row.Item("SubTotal").Equals(DBNull.Value) Then
                scripts.AppendFormat("|""{0}""", DirectCast(data_row("SubTotal"), Integer).ToString("$#,###"))
            End If

            Response.Write(Regex.Replace(scripts.ToString(), "[,]\d*(\.\d+)?[|]", ","))

            scripts.Remove(0, scripts.Length)

            Response.Write(Environment.NewLine)
            Response.Flush()

        Next

    End Sub

    Private Function ObjectToString_Converter(ByVal element As Object) As String
        Return element.ToString().Trim()
    End Function

    Private Function ObjectToString_Converter(ByVal element As DataControlField) As String
        Return element.HeaderText.Trim()
    End Function

    Private Function DataSourceTable_Generation() As DataTable

        Me.SearchSqlParameters_Preparation()
        Me.SearchSqlStatementWhereConditions_Preparation()

        Dim data_source As DataTable = CType(sdsSignRecords.Select(New DataSourceSelectArguments), DataView).Table

        data_source.Columns.Add(New DataColumn("Hours", GetType(Integer)))

        With data_source.Columns(data_source.Columns.Count - 1)
            .AllowDBNull = True
            .DefaultValue = DBNull.Value
        End With

        data_source.Columns.Add(New DataColumn("SubTotal", GetType(Integer)))

        With data_source.Columns(data_source.Columns.Count - 1)
            .AllowDBNull = True
            .DefaultValue = DBNull.Value
        End With

        Dim result As List(Of Integer) = Nothing

        For Each row As DataRow In data_source.Rows

            MOA01009.HoursAndSubTotal_Calculations _
                     ( _
                             row("Out_Time_nvc"), _
                             row("HourLimit"), _
                             row("OverTime"), _
                             row("NoonHour"), _
                             Me.Over_Time_Baseline, _
                             result _
                     )

            If Not result.Item(0).Equals(Integer.MinValue) Then
                row("Hours") = result.Item(0).ToString()
            End If

            If Not result.Item(1).Equals(Integer.MinValue) Then
                row("SubTotal") = result.Item(1).ToString()
            End If

        Next

        If Not IsNothing(result) Then
            result.Clear()
            result = Nothing
        End If

        With data_source.Columns
            .Remove("EMPLOYEE_ID")
            .Remove("HourLimit")
            .Remove("OverTime")
        End With

        Return data_source

    End Function

    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click

        If String.IsNullOrEmpty(ddl_org_picker.SelectedValue) Or String.IsNullOrEmpty(ddl_user_picker.SelectedValue) Then
            scripts.Append("- 請選擇單位及人員姓名，報表列印僅支援個人加班資料 !!\n")
        End If

        If Not Date.Parse(tb_begin_date.Text).ToString("yyyy/M/1").Equals(Date.Parse(tb_end_date.Text).ToString("yyyy/M/1")) Then
            scripts.Append("- 請確認日期起迄為同一月份，報表列印僅支援單月加班資料 !!\n")
        End If

        If scripts.Length.Equals(0) Then

            Dim btnControl As ImageButton = CType(sender, ImageButton)

            Select Case btnControl.CommandName

                Case "Confirm"

                    scripts.Append("function ShowConfirmMessageBox() { if (confirm('是否要施行擇期補休 ?\n')) ") _
                           .Append("document.getElementById('hidCompensatoryLeave').value = true; else ") _
                           .Append("document.getElementById('hidCompensatoryLeave').value = false; }")

                    ClientScript.RegisterClientScriptBlock(Me.GetType(), "ShowConfirmMessageBox_Codes", scripts.ToString(), True)
                    scripts.Remove(0, scripts.Length)

                    ClientScript.RegisterHiddenField("hidCompensatoryLeave", String.Empty)

                    ClientScript.RegisterStartupScript _
                                 ( _
                                         Me.GetType(), _
                                        "TriggerPrintButton", _
                                         String.Format("document.getElementById('{0}').click();", Me.ImagePrint.ClientID), _
                                         True _
                                 )

                    btnControl.CommandName = "Print"
                    btnControl.OnClientClick = "javascript:ShowConfirmMessageBox();"

                Case "Print"

                    Me.ImgSearch_Click(Me.ImgSearch, Nothing)
                    Me.sdsSignRecords.Select(New DataSourceSelectArguments)

                    Dim htParameters As New Hashtable()
                    Dim current_date_time As DateTime = DateTime.Now

                    With htParameters
                        .Add("Compensatory_Leave", Me.Request.Form("hidCompensatoryLeave").Trim())
                        .Add("Over_Time_Baseline", Me.Over_Time_Baseline.ToString())
                        .Add("Report_BeginDate", Me.tb_begin_date.Text)
                        .Add("SQL_Statements", _
                              sdsSignRecords.SelectCommand _
                                            .Replace("P.HourLimit,", String.Empty) _
                                            .Replace("P.OverTime,", String.Empty) _
                                            .Replace("Select E.EMPLOYEE_ID,", "Select"))
                    End With

                    For Each parameter As Parameter In sdsSignRecords.SelectParameters

                        htParameters.Add _
                                     ( _
                                            String.Format("@{0}", parameter.Name), _
                                            String.Join _
                                                   ( _
                                                           "|", _
                                                            New String() _
                                                            { _
                                                                    parameter.DbType.ToString("D"), _
                                                                    parameter.DefaultValue _
                                                            } _
                                                   ) _
                                     )

                    Next

                    Using file_writer As New System.IO.StreamWriter _
                                                       ( _
                                                             String.Format _
                                                                    ( _
                                                                           Server.MapPath("~/Drs/") & "SERIAL01009{0:yyyyMMddHHmmssffff}.dat", _
                                                                           current_date_time _
                                                                    ), _
                                                             False, _
                                                             Encoding.GetEncoding("BIG5") _
                                                       )

                        Dim formatter As New BinaryFormatter()

                        Try
                            formatter.Serialize(file_writer.BaseStream, htParameters)
                        Catch ex As Exception
                            Throw
                        Finally
                            file_writer.Close()
                        End Try

                    End Using

                    htParameters.Clear()
                    htParameters = Nothing

                    btnControl.CommandName = "Confirm"
                    btnControl.OnClientClick = Nothing

                    scripts.Append("function Open_ReportWindow() { var setting = 'channelmode=no, toolbar=no, location=no, directories=no, status=no, ") _
                           .Append("menubar=no, scrollbars=no, resizable=yes, width=810, height=640, top=' + ((screen.availHeight - 640) / 2).toString() + ', ") _
                           .AppendFormat("left=' + ((screen.availWidth - 810) / 2).toString(); var newWindow = window.open('{0}", Me.Request.ApplicationPath) _
                           .AppendFormat("/M_Source/01/MOA01010.aspx?sn={0:yyyyMMddHHmmssffff}', null, setting); ", current_date_time) _
                           .Append("if (newWindow) newWindow.focus(); setting = null; }")

                    ClientScript.RegisterClientScriptBlock(Me.GetType(), "Open_ReportWindow_Codes", scripts.ToString(), True)
                    scripts.Remove(0, scripts.Length)

                    ClientScript.RegisterStartupScript(Me.GetType(), "Open_ReportWindow_Action", "Open_ReportWindow();", True)
                    current_date_time = Nothing

            End Select

        Else

            scripts.Insert(0, "alert('").Append("');")

            ClientScript.RegisterStartupScript(Me.GetType(), "PrintError_Message", scripts.ToString(), True)
            scripts.Remove(0, scripts.Length)

        End If

    End Sub

End Class