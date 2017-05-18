Imports System.Data.SqlClient
Partial Class OA_above
    Inherits System.Web.UI.Page

    Protected Sub Logout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Logout.Click

        FormsAuthentication.SignOut()

        '清空Session
        Session("user_id") = ""
        Session("ORG_UID") = ""
        Session("PARENT_ORG_UID") = ""
        Session("Role") = ""
        Session("PageUp") = Nothing
        Response.Write("<script language='javascript'>window.parent.location='../../index.aspx'</Script>")

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '新增資料
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '人員是否已加入
        db.Open()
        Dim EmpFlag As String = ""
        Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & Session("user_id") & "'", db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.read() Then
            UserName.Text = RdPer("emp_chinese_name")
        End If
        db.Close()

        '判斷是否由網域整合性驗證登入
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then
            Logout.Visible = False
        End If

    End Sub

    Protected Sub GoFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GoFirst.Click

        Response.Write(" <script language='javascript'>")
        Response.Write(" window.parent.location='http://portalkm.oa.mil.tw/MOAKM/';")
        Response.Write(" </script>")

    End Sub

End Class
