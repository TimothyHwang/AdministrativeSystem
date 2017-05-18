
Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_04_MOA04024
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim seat_name As String = e.NewValues("seat_name")
            chk.CheckDataLen(seat_name, 50, "修改時：<儲庫名稱>", True)
        Catch ex As Exception
            e.Cancel = True
            ErrMsg.Text = ex.Message
        End Try
    End Sub

    Protected Sub ImgInsert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim Seat_Name As TextBox
            Dim Seat_num As TextBox

            Seat_num = Me.DetailsView1.FindControl("tb_seat_num")
            Seat_Name = Me.DetailsView1.FindControl("tb_seat_name")
            chk.CheckDataLen(Seat_Name.Text, 50, "新增時：<儲庫名稱>", True)
            chk.CheckDataLen(Seat_num.Text, 10, "新增時：<儲庫編號>", True)

            Dim scripts As New StringBuilder
            Dim sql_function As New C_SQLFUN
            Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

            scripts.Append("Select 1 seat_num From P_0417 with(nolock) where seat_num = @seat_num ")

            command.CommandType = CommandType.Text
            command.CommandText = scripts.ToString()
            command.Parameters.Add(New SqlParameter("seat_num", SqlDbType.VarChar, 10)).Value = Seat_num.Text

            command.Connection.Open()
            Dim exobject As New Object
            exobject = command.ExecuteScalar()
            If exobject <> Nothing Then
                ErrMsg.Text = "儲庫編號已重複!"
            Else
                Dim btnInsert As ImageButton
                btnInsert = Me.DetailsView1.FindControl("ImgInsert")
                btnInsert.CommandName = "insert"
            End If
        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try
    End Sub

    Protected Sub GridView1_RowDeleting(sender As Object, e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim seat_num As String = GridView1.DataKeys(e.RowIndex).Value

            Dim sql_function As New C_SQLFUN
            Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
            Dim scripts As New StringBuilder
            scripts.Append("Select 1 seat_num From P_0414 with(nolock) where seat_num = @seat_num ")

            command.CommandType = CommandType.Text
            command.CommandText = scripts.ToString()
            command.Parameters.Add(New SqlParameter("seat_num", SqlDbType.VarChar, 10)).Value = seat_num

            command.Connection.Open()
            Dim exobject As New Object
            exobject = command.ExecuteScalar()
            If exobject <> Nothing Then
                e.Cancel = True
                ErrMsg.Text = "此儲庫編號已在使用中，無法刪除!"
                'Else
                '    Dim btnInsert As ImageButton
                '    btnInsert = Me.DetailsView1.FindControl("ImgInsert")
                '    btnInsert.CommandName = "insert"
            End If

        Catch ex As Exception
            e.Cancel = True
            ErrMsg.Text = ex.Message
        End Try
    End Sub
End Class
