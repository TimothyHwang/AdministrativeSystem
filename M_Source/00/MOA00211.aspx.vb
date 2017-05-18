Imports System.Data.SqlClient
Partial Class M_Source_00_MOA00211
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA00210.aspx")

    End Sub

    Protected Sub ImgOK_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgOK.Click

        Try

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            If txtTitle.Text = "" Or txtName.Text = "" Or txtTel.Text = "" Or EDate.Text = "" Or txtContent.Text = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('全部欄位皆為必填，請全部輸入!');")
                Response.Write(" </script>")

            Else

                '新增公告資料
                db.Open()
                Dim insCom As New SqlCommand("INSERT INTO T_ANNOUNTCEMENT(DATE_d,DEPT_i,TITLE_nvc,CONTENT_nvc,uicode,DATE_e,ANN_Name,ANN_Phone) VALUES ('" & Now().Date & "','" & org_uid & "','" & txtTitle.Text & "','" & txtContent.Text & "','" & DDLUnicode.SelectedValue & "','" & EDate.Text & "','" & txtName.Text & "','" & txtTel.Text & "')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                Server.Transfer("MOA00210.aspx")

            End If

        Catch ex As Exception

        End Try

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

                If LoginCheck.LoginCheck(user_id, "MOA00210") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00211.aspx")
                    Response.End()
                End If

                Dim conn As New C_SQLFUN
                Dim connstr As String = conn.G_conn_string
                Dim db As New SqlConnection(connstr)
                db.Open()
                Dim strSqlComd As New SqlCommand("select emp_chinese_name from EMPLOYEE where employee_id='" + user_id + "'", db)
                Dim dr = strSqlComd.ExecuteReader()
                If dr.read() Then
                    txtName.Text = dr.item("emp_chinese_name")
                End If
                db.Close()


            End If

        Catch ex As Exception

        End Try

    End Sub
End Class
