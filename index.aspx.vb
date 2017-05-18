Imports System.Data
Imports System.Data.SqlClient

Partial Class index
    Inherits System.Web.UI.Page
    Dim connstr As String
    Dim conn As New C_SQLFUN    '新增連線

    Protected Sub login_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Login.Click

        Try
            '呼叫檢查帳號密碼程式，成功會傳回True   
            If ChkLogin(UserName.Text, password.Text) Then

                '取得連線字串
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '新增登入者資料
                db.Open()
                Dim insCom As New SqlCommand("INSERT INTO LoginLog(Login_IP,Login_ID) VALUES ('" & Request.ServerVariables("REMOTE_ADDR") & "','" & UserName.Text & "')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                '執行FormsAuthentication.RedirectFromLoginPage，並且將是否記住帳號一併處理
                FormsAuthentication.RedirectFromLoginPage(UserName.Text, True)

            Else
                loginerr.Text = "帳號或密碼錯誤!!!!!!"
                'Throw New Exception("登入失敗")
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try

    End Sub

    Public Function ChkLogin(ByVal UserName As String, ByVal UserPassword As String) As Boolean
        '檢查帳號密碼程式, 成功會傳回True
        Dim bl_ChkLogin As Boolean = False
        Using MyConnection As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            Try
                MyConnection.Open()
                Dim MyCommand1 As New SqlCommand("select employee_id,password from EMPLOYEE where employee_id='" & UserName & " ' and password='" & UserPassword & "' and leave='Y'", MyConnection)
                Dim reader1 As System.Data.SqlClient.SqlDataReader
                reader1 = MyCommand1.ExecuteReader()
                If reader1.Read() Then
                    bl_ChkLogin = True
                Else
                    bl_ChkLogin = False
                End If
            Catch ex As Exception
                'MsgBox(ex.Message)
            Finally
                MyConnection.Close()
            End Try
        End Using 
        Return bl_ChkLogin
    End Function
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '當整合性驗證時直接登入系統
        If Page.User.Identity.Name.ToString <> "" Then

            Dim LoginFlag As String = ""    '判斷是否為STAFF登入

            If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

                Dim LoginAll As String = Page.User.Identity.Name.ToString

                Dim LoginID() As String = Split(LoginAll, "\")

                LoginFlag = LoginID(0)
            End If

            'If UCase(LoginFlag) <> "STAFF" Then
            '    Response.Write("請先登入STAFF網域")
            '    Response.End()
            'End If

            '寫入登入者資料
            Dim LoginAction As New C_Public

            LoginAction.LoginAction(Request.ServerVariables("REMOTE_ADDR"), Page.User.Identity.Name.ToString, "index.aspx")

            '執行FormsAuthentication.RedirectFromLoginPage，並且將是否記住帳號一併處理
            FormsAuthentication.RedirectFromLoginPage(Page.User.Identity.Name.ToString, True)

        End If

    End Sub

End Class
