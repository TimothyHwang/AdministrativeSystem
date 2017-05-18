Imports System.Data.SqlClient
Partial Class Source_02_MOA02006
    Inherits System.Web.UI.Page

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Try

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            If DevName.Text <> "" Then

                '新增會議室設備資料
                db.Open()
                Dim insCom As New SqlCommand("INSERT INTO P_0202(DeviceName) VALUES ('" & DevName.Text & "')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                Server.Transfer("MOA02006.aspx")

            End If

        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try

    End Sub
End Class
