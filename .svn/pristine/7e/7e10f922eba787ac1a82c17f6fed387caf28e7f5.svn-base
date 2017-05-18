Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Partial Class M_Source_04_MOA04025_1
    Inherits System.Web.UI.Page
    Dim scripts As New StringBuilder
    Dim chk As New C_CheckFun
    Dim user_id, org_uid As String
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

        If CType(Request.QueryString("itcode"), String) Is Nothing Or CType(Request.QueryString("usecheck"), String) Is Nothing Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('您的參數已遺失，請重新操作，感謝您！');")
            Response.Write("window.close();")
            Response.Write("</script>")
        Else
            Dim it_code As String = CType(Request.QueryString("itcode"), String)
            Dim usecheck As String = CType(Request.QueryString("usecheck"), String)
            If it_code.Length <> 6 And usecheck.Length <> 1 Then
                Response.Write("<script language='javascript'>")
                Response.Write("alert('您的參數不正確，請重新操作，感謝您！');")
                Response.Write("window.close();")
                Response.Write("</script>")
            End If
            lbl_it_code.Text = it_code
            lbl_usecheck.Text = usecheck
            Dim checkstatus As String
            Select Case usecheck
                Case "0"
                    checkstatus = "庫存"
                Case "1"
                    checkstatus = "待出庫"
                Case "2"
                    checkstatus = "出庫"
            End Select
            lblTitle.Text = "廢料查詢-物料詳細資料(" + checkstatus + ")"

            Dim dt As New DataTable
            dt = GetItemP_0425DetailData(it_code)
            If dt Is Nothing Or dt.Rows.Count < 1 Then
                Response.Write("<script language='javascript'>")
                Response.Write("alert('此物品目前查不到詳細資料，請重新操作，感謝您！');")
                Response.Write("window.close();")
                Response.Write("</script>")
            End If

            RptitemList.DataSource = dt
            RptitemList.DataBind()
            ErrMsg.Text = String.Empty
        End If

    End Sub

    Private Function GetItemP_0425DetailData(ByVal it_code As String) As DataTable
        GetItemP_0425DetailData = New DataTable("Detail")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        scripts.Append("select it_code,it_name,isnull(file_a,'') as file_a,isnull(file_b,'') as file_b from P_0407 with(nolock) where it_code=@it_code")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("it_code", SqlDbType.NVarChar, 6)).Value = it_code
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            GetItemP_0425DetailData.Load(dr)

        Catch ex As Exception
            GetItemP_0425DetailData = Nothing

        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

    
    Public Function showPic(ByVal filename As String) As Boolean
        If filename.Length <= 3 Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function showPicLB(ByVal filename As String) As Boolean
        If filename.Length <= 3 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function showOutLB(ByVal usecheck As String) As Boolean
        If usecheck = "0" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function showeditdelLB(ByVal usecheck As String) As Boolean
        If usecheck = "2" Then
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub GVItemList_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVItemList.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.FindControl("ddl_seat_num") IsNot Nothing Then
                Dim ddl_seat_num As DropDownList = CType(e.Row.FindControl("ddl_seat_num"), DropDownList)
                sqlds_Seat_Num.SelectCommand = "SELECT [seat_num]+'_'+[seat_name] as seatnum,[seat_num], [seat_name] FROM [P_0417] ORDER BY [seat_num]"
                ddl_seat_num.DataSource = sqlds_Seat_Num
                ddl_seat_num.DataTextField = "seatnum"
                ddl_seat_num.DataValueField = "seat_num"
                ddl_seat_num.DataBind()
                ddl_seat_num.SelectedValue = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "seat_num"))
            End If
        End If
    End Sub

    Protected Sub GVItemList_RowUpdating(sender As Object, e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GVItemList.RowUpdating
        Try
            Dim ddl As DropDownList = DirectCast(GVItemList.Rows(e.RowIndex).FindControl("ddl_seat_num"), DropDownList)
            Dim it_spec As String = e.NewValues("it_spec")
            Dim receive As String = e.NewValues("receive")
            chk.CheckDataLen(ddl.SelectedValue.ToString, 20, "修改時：<儲庫編號>", True)
            chk.CheckDataLen(it_spec, 255, "修改時：<物料規格>", True)
            'chk.CheckDataLen(receive, 10, "修改時：<領取人員>", True)
            sqldatalist.UpdateParameters("Seat_num").DefaultValue = ddl.SelectedValue
        Catch ex As Exception
            e.Cancel = True
            ErrMsg.Text = ex.Message
        End Try
    End Sub

    Protected Sub GVItemList_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GVItemList.RowCommand
        Select Case e.CommandName
            Case "UPOut"
                Dim shcode As String = e.CommandArgument.ToString()
                Dim blCheckStatus As Boolean = bl_ChectOutDataExist(shcode)

                If blCheckStatus = True Then
                    Dim sql_function As New C_SQLFUN
                    Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

                    scripts.Append("update P_0414 set UseKeyIn=@UseKeyIn ,UseDate=getdate(),usecheck='2' where shcode=@shcode and usecheck = '0' and shtype='1'")

                    command.CommandType = CommandType.Text
                    command.CommandText = scripts.ToString()
                    scripts.Remove(0, scripts.Length)
                    command.Parameters.Add(New SqlParameter("UseKeyIn", SqlDbType.NVarChar, 10)).Value = user_id
                    command.Parameters.Add(New SqlParameter("shcode", SqlDbType.NVarChar, 18)).Value = shcode

                    Try
                        command.Connection.Open()
                        command.ExecuteNonQuery()

                    Catch ex As Exception
                        ErrMsg.Text = "出庫失敗:" + ex.Message

                    Finally
                        If command.Connection.State.Equals(ConnectionState.Open) Then
                            command.Connection.Close()
                        End If
                        command.Dispose()
                        command = Nothing
                    End Try
                Else
                    ErrMsg.Text = "您的物料規格或領取人未填，請先進行編輯再做出庫操作。"
                End If
                GVItemList.DataBind()
        End Select
    End Sub

    Private Function bl_ChectOutDataExist(ByVal shcode As String) As Boolean
        bl_ChectOutDataExist = False
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        Dim OutDatadt As New DataTable("dtOut")
        scripts.Append("select it_spec,receive from P_0414 with(nolock) where shcode=@shcode and shtype=1")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("shcode", SqlDbType.NVarChar, 18)).Value = shcode
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            OutDatadt.Load(dr)

        Catch ex As Exception
            OutDatadt = Nothing

        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try

        If OutDatadt Is Nothing Then
            bl_ChectOutDataExist = False
        Else
            If (OutDatadt.Rows.Count > 0) Then
                If (OutDatadt.Rows(0)("it_spec").ToString().Length > 0 And OutDatadt.Rows(0)("receive").ToString().Length > 0) Then
                    bl_ChectOutDataExist = True
                End If
            End If
        End If

    End Function


End Class
