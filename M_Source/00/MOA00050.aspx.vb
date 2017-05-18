Imports System.Data.SqlClient
Imports System.Data

Partial Class Source_00_MOA00050
    Inherits System.Web.UI.Page
    Dim connstr As String = ""
    Dim user_id, org_uid As String

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Try
            Dim SelDel As String = ""
            '取得刪除的群組代碼
            SelDel = GridView1.DataKeys(e.RowIndex).Value
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)
            Dim GroupPer As Integer

            '判斷群組是否有人
            db.Open()
            Dim strPer As New SqlCommand("SELECT count(*) as GroupPer FROM ROLEGROUPITEM WHERE Group_Uid = '" & SelDel & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                GroupPer = RdPer("GroupPer")
            End If
            db.Close()
            If GroupPer > 0 Then
                '不刪除群組
                SqlDataSource1.DeleteCommand = "DELETE FROM [ROLEGROUP] WHERE 1=2"
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('此群組內還有人員未移除')")
                Response.Write(" </script>")
            Else
                '刪除群組
                SqlDataSource1.DeleteCommand = "DELETE FROM [ROLEGROUP] WHERE [Group_Uid] = '" & SelDel & "'"
            End If
        Catch ex As Exception
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "刪除時發生錯誤，請重新再試!!"
        End Try
    End Sub

    Protected Sub btnInsert_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If CType(Me.DetailsView1.FindControl("Group_Name_Ins"), TextBox).Text = "" Then
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "請輸入群組名稱"
        ElseIf CType(Me.DetailsView1.FindControl("Group_Order"), TextBox).Text = "" Then
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "請輸入群組順序"
        Else
            '新增資料
            If Me.DetailsView1.CurrentMode = DetailsViewMode.Insert Then
                Dim randstr As New C_Public
                Dim Group_Uid = randstr.randstr(10)
                CType(Me.DetailsView1.FindControl("Group_Uid"), TextBox).Text = Group_Uid
                '判斷群組名稱與群組順序是否重複
                If ChkGroupData(CType(Me.DetailsView1.FindControl("Group_Name_Ins"), TextBox).Text.Trim, "", Group_Uid) Then
                    CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "您欲新增的群組名稱已存在!!"
                ElseIf ChkGroupData("", CType(Me.DetailsView1.FindControl("Group_Order"), TextBox).Text.Trim, Group_Uid) Then
                    CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "您欲新增的群組順序已存在!!"
                Else
                    CType(Me.DetailsView1.FindControl("btninsert"), ImageButton).CommandName = "insert"
                End If
            End If
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            user_id = Session("user_id")
            org_uid = Session("ORG_UID")
            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            Else
                '判斷登入者權限
                Dim LoginCheck As New C_Public
                If LoginCheck.LoginCheck(user_id, "MOA00050") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00050.aspx")
                    Response.End()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub GridView1_RowUpdating(sender As Object, e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim Group_Uid As HiddenField = CType(GridView1.Rows(e.RowIndex).FindControl("HF_Group_Uid"), HiddenField)
        Dim GV_GroupName As TextBox = CType(GridView1.Rows(e.RowIndex).FindControl("GV_GroupName"), TextBox)
        Dim GV_GroupOrder As TextBox = CType(GridView1.Rows(e.RowIndex).FindControl("GV_GroupOrder"), TextBox)
        If GV_GroupName.Text.Trim() = "" Or GV_GroupOrder.Text.Trim() = "" Then
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "請輸入群組名稱"
            e.Cancel = True
        End If
        '判斷群組名稱是否重複
        If ChkGroupData(GV_GroupName.Text.Trim, "", Group_Uid.Value) Then
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "您欲更新的群組名稱已存在!!"
            e.Cancel = True
        End If
        '判斷群組順序是否重複
        If ChkGroupData("", GV_GroupOrder.Text.Trim, Group_Uid.Value) Then
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "您欲更新的群組順序已存在!!"
            e.Cancel = True
        End If
    End Sub

    Public Function ChkGroupData(ByVal GroupName As String, ByVal GroupOrder As String, ByVal Group_Uid As String) As Boolean
        '檢查GroupName or GroupOrder,已存在會傳回True
        Dim bl_ChkLogin As Boolean = False
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand
        Using MyConnection As New SqlConnection(sql_function.G_conn_string)
            Try
                Dim sQuery As String = "select 1 from ROLEGROUP where Group_Uid <> @Group_Uid and "
                command.Parameters.Add("@Group_Uid", SqlDbType.VarChar, 10).Value = Group_Uid
                If GroupName <> "" Then
                    sQuery += "Group_Name=@Group_Name"
                    command.Parameters.Add("@Group_Name", SqlDbType.VarChar, 30).Value = GroupName
                Else
                    sQuery += "Group_Order=@Group_Order"
                    command.Parameters.Add("@Group_Order", SqlDbType.Int).Value = Int16.Parse(GroupOrder)
                End If
                MyConnection.Open()
                command.CommandText = sQuery
                command.Connection = MyConnection
                Dim ob As Object = command.ExecuteScalar()
                If Not ob Is Nothing Then
                    bl_ChkLogin = True
                End If
            Catch ex As Exception
                bl_ChkLogin = False
            Finally
                If command.Connection.State.Equals(ConnectionState.Open) Then
                    command.Connection.Close()
                End If
                command.Dispose()
                command = Nothing
            End Try
        End Using
        Return bl_ChkLogin
    End Function
End Class
