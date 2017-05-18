Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Partial Class M_Source_04_MOA04022_1
    Inherits System.Web.UI.Page
    Dim scripts As New StringBuilder
    Dim chk As New C_CheckFun
    Dim user_id, org_uid As String
    Public bl_checkFlag As Boolean = False
    Public EFORMSN As String


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

        If CType(Request.QueryString("EFORMSN"), String) Is Nothing Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('您的參數已遺失，請重新操作，感謝您！');")
            Response.Write("window.close();")
            Response.Write("</script>")
        Else
            Dim EFORMSN As String = CType(Request.QueryString("EFORMSN"), String)
            If EFORMSN.Length <> 16 Then
                Response.Write("<script language='javascript'>")
                Response.Write("alert('您的參數不正確，請重新操作，感謝您！');")
                Response.Write("window.close();")
                Response.Write("</script>")
            End If
            lbl_EFORMSN.Text = EFORMSN
            ErrMsg.Text = String.Empty
        End If
    End Sub

    Private Function bl_GetItem(ByVal shcode As String) As Boolean
        bl_GetItem = False
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        scripts.Append("update P_0414 set usecheck = 2, UseDate = getdate(),UseKeyIn=@UseKeyIn where shcode=@shcode and usecheck = 1")
        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("shcode", SqlDbType.NVarChar, 18)).Value = shcode
        command.Parameters.Add(New SqlParameter("UseKeyIn", SqlDbType.VarChar, 10)).Value = user_id
        Try
            command.Connection.Open()
            command.ExecuteNonQuery()
            bl_GetItem = True
        Catch ex As Exception
            ErrMsg.Text = "勾選領料失敗：" + ex.Message

        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function
    Private Function bl_FinisheForm(ByVal EFORMSN As String) As Boolean
        bl_FinisheForm = False
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        scripts.Append("update P_0415 set nAppStockStatus = 'Y' where EFORMSN=@EFORMSN")
        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("EFORMSN", SqlDbType.NVarChar, 16)).Value = EFORMSN
        Try
            command.Connection.Open()
            command.ExecuteNonQuery()
            bl_FinisheForm = True
        Catch ex As Exception
            ErrMsg.Text = "完結領料作業失敗：" + ex.Message
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

    Public Function showcheckbox(ByVal usecheck As String) As Boolean
        If usecheck = "2" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function enableGetCK(ByVal usecheck As String) As Boolean
        If usecheck = "2" Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function bl_ChectOutDataExist(ByVal Job_Num As String) As Boolean
        bl_ChectOutDataExist = False
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        Dim OutDatadt As New DataTable("dtOut")
        scripts.Append("select 1 from P_0414 with(nolock) where Job_Num=@Job_Num and usecheck=1")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("Job_Num", SqlDbType.NVarChar, 16)).Value = Job_Num
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
                bl_ChectOutDataExist = True
            End If
        End If

    End Function

    Protected Sub btGetItem_Click(sender As Object, e As System.EventArgs) Handles btGetItem.Click
        Dim s_ItemshCodeValue As String = String.Empty
        Dim bl_checkitem As Boolean = False
        For i As Integer = 0 To GVItemList.Rows.Count - 1
            Dim ck As CheckBox = CType(GVItemList.Rows(i).Cells(0).FindControl("ck_GetForm"), CheckBox)
            If ck.Checked = True And ck.Enabled = True Then
                Dim lb_shcode As Label = CType(GVItemList.Rows(i).Cells(1).FindControl("lb_shcode"), Label)
                s_ItemshCodeValue = s_ItemshCodeValue + lb_shcode.Text + ","
                bl_checkitem = True
            End If
        Next
        If bl_checkitem = True Then
            Dim itemarray() As String = Split(s_ItemshCodeValue.Substring(0, s_ItemshCodeValue.Length - 1), ",")
            For j As Integer = 0 To itemarray.Length - 1
                Dim blGetItem As Boolean = bl_GetItem(itemarray(j))
                If blGetItem = False Then
                    ErrMsg.Text = "領取物品:" + itemarray(j) + "失敗。"
                    bl_checkitem = False
                    Exit For
                End If
            Next
        Else
            ErrMsg.Text = "請至少勾選一個您欲領取的物品。"
        End If

        GVItemList.DataBind()

    End Sub

    Protected Sub btFinishEForm_Click(sender As Object, e As System.EventArgs) Handles btFinishEForm.Click
        EFORMSN = lbl_EFORMSN.Text
        Dim blCheckStatus As Boolean = bl_ChectOutDataExist(EFORMSN)

        If blCheckStatus = True Then
            ErrMsg.Text = "您有未領料的物品存在，請先完成所有物品領料再進行完結。"
        Else
            btGetItem.Enabled = False
            Dim blFinish As Boolean = bl_FinisheForm(EFORMSN)
            If blFinish = False Then
                ErrMsg.Text = "完結作業失敗!!"
            Else
                ErrMsg.Text = "完結作業成功!!"
                btFinishEForm.Enabled = False
            End If
        End If
        GVItemList.DataBind()
    End Sub

    Protected Sub GVItemList_DataBound(sender As Object, e As System.EventArgs) Handles GVItemList.DataBound
        If GVItemList.Rows.Count = 0 Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('本報修單無任何未完結的領料資訊！');")
            Response.Write("window.close();")
            Response.Write("</script>")
        End If

        If (bl_checkFlag) Then
            btGetItem.Enabled = True
        End If
    End Sub

    Protected Sub GVItemList_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVItemList.RowDataBound
        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then
            Dim ck_GetForm As CheckBox
            ck_GetForm = e.Row.Cells(0).FindControl("ck_GetForm")
            If (ck_GetForm.Enabled) Then
                bl_checkFlag = True
            End If
        End If
    End Sub
End Class
