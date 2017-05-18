Imports System.Data
Imports System.Data.sqlclient
Imports System.IO
Partial Class Source_00_MOA00105
    Inherits System.Web.UI.Page

    Dim eformsn As String = ""

    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Try
            '是否有檔案
            If UpFile.HasFile Then

                Dim PostFile As HttpPostedFile = UpFile.PostedFile
                Dim UpPath, FileName As String

                FileName = eformsn & "-" & Path.GetFileName(PostFile.FileName)

                '儲存檔案名稱加上eformsn產生唯一性
                UpPath = Server.MapPath("/MOA/M_Source/99/" & FileName)

                '檔案上傳
                UpFile.SaveAs(UpPath)

                '新增資料
                Dim connstr As String
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '新增UPLOAD基本資料
                db.Open()
                Dim insCom As New SqlCommand("INSERT INTO UPLOAD(eformsn,FileName,FilePath) VALUES ('" & eformsn & "','" & FileName & "','/MOA/M_Source/99/')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                Server.Transfer("MOA00105.aspx")

            End If

        Catch ex As Exception
            ErrLab.Text = ex.Message
        End Try

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'eformsn
        eformsn = Request.QueryString("eformsn")

        '重新讀取SQL
        SqlDataSource1.SelectCommand = ""

        Dim sqlstr As String = ""

        sqlstr = "SELECT * FROM UPLOAD WHERE eformsn='" & eformsn & "'"

        SqlDataSource1.SelectCommand = sqlstr


    End Sub

End Class
