Imports System.Data.SqlClient
Partial Class M_Source_00_MOA00111
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        do_sql.G_user_id = Session("user_id")
    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click
        Try

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            If comment.Text <> "" Then

                '新增會議室設備資料
                db.Open()
                Dim insCom As New SqlCommand("INSERT INTO Phrase(employee_id,comment) VALUES ('" & do_sql.G_user_id & "','" & comment.Text & "')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                'Server.Transfer("MOA02006.aspx")
                GridView1.DataBind()
                comment.Text = ""
            End If

        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
    End Sub
End Class
