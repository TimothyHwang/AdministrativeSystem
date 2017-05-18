Imports System.Data.SqlClient
Partial Class Source_00_MOA00109
    Inherits System.Web.UI.Page

    Dim EmpFlag As String
    Dim user_id, org_uid As String
    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Try

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            If CVar.Text <> "" Then

                '基本代碼是否已存在
                db.Open()
                Dim strPer As New SqlCommand("SELECT CONFIG_VAR FROM SYSCONFIG WHERE CONFIG_VAR = '" & CVar.Text & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    EmpFlag = "1"
                End If
                db.Close()

                '判斷人員是否重複
                If EmpFlag = "" Then

                    '新增基本設定資料
                    db.Open()
                    Dim insCom As New SqlCommand("INSERT INTO SYSCONFIG(CONFIG_VAR,CONFIG_DESC,CONFIG_VALUE) VALUES ('" & CVar.Text & "','" & CDesc.Text & "','" & CVal.Text & "')", db)
                    insCom.ExecuteNonQuery()
                    db.Close()

                Else

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('代碼已存在!!!');")
                    Response.Write(" </script>")

                End If

                Server.Transfer("MOA00109.aspx")

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        EmpFlag = ""

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

                If LoginCheck.LoginCheck(user_id, "MOA00109") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00109.aspx")
                    Response.End()
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub
End Class
