Imports System.Data.SqlClient
Partial Class OA_System_OrgRight
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")

        '判斷單位
        org_uid = Request.QueryString("org_uid")


        'session被清空回首頁
        If user_id = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA00030") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00032.aspx")
                Response.End()
            End If

            btnIns.PostBackUrl = "MOA00033.aspx?org_uid=" & org_uid & "&UseFlag=1"
            btnUpd.PostBackUrl = "MOA00033.aspx?org_uid=" & org_uid & "&UseFlag=2"


        End If


    End Sub

    Protected Sub btnDel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnDel.Click

        Try
            Dim DelFlag As String = ""

            '新增連線
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷是否有下屬單位
            db.Open()
            Dim strUnit As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
            Dim RdUnit = strUnit.ExecuteReader()
            If RdUnit.read() Then
                DelFlag = "1"
            End If
            db.Close()

            '判斷單位是否有人員
            db.Open()
            Dim strPer As New SqlCommand("SELECT empuid FROM EMPLOYEE WHERE ORG_UID = '" & org_uid & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                DelFlag = "2"
            End If
            db.Close()

            If DelFlag = "" Then

                '刪除單位資料
                db.Open()
                Dim insCom As New SqlCommand("DELETE FROM ADMINGROUP WHERE (ORG_UID = '" & org_uid & "')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('單位刪除完成');")
                Response.Write(" window.parent.location='MOA00030.aspx';")
                Response.Write(" </script>")

            ElseIf DelFlag = "1" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('還有下屬單位不可刪除');")
                Response.Write(" </script>")

            ElseIf DelFlag = "2" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('單位還有人員不可刪除');")
                Response.Write(" </script>")

            End If

        Catch ex As Exception

        End Try

    End Sub
End Class
