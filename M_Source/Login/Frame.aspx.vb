Imports System.Data
Imports System.Data.SqlClient
Partial Class OA_Frame
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            '判斷是否由網域整合性驗證登入
            '登入者帳號
            If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

                Dim LoginAll As String = Page.User.Identity.Name.ToString

                Dim LoginID() As String = Split(LoginAll, "\")

                Session("user_id") = LoginID(1)
            Else
                Session("user_id") = Page.User.Identity.Name.ToString
            End If

            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '登入者組織代號
            db.Open()
            Dim strPer As New SqlCommand("SELECT ORG_UID,PARENT_ORG_UID,ORG_NAME,TU_ID FROM V_EmpInfo WHERE employee_id = '" & Session("user_id") & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                Session("ORG_UID") = RdPer("ORG_UID")
                Session("ORG_NAME") = RdPer("ORG_NAME")
                Session("PARENT_ORG_UID") = RdPer("PARENT_ORG_UID")
                Session("TU_ID") = RdPer("TU_ID")
            End If
            db.Close()

            '判斷登入者角色
            Dim RoleDT As DataTable
            db.Open()
            Dim RoleSQL As String = "select ROLEGROUP.Group_Uid from ROLEGROUP,ROLEGROUPITEM where ROLEGROUP.Group_Uid = ROLEGROUPITEM.Group_Uid AND ROLEGROUPITEM.employee_id = '" & Session("user_id") & "'"
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(RoleSQL, db)
            da.Fill(ds)
            RoleDT = ds.Tables(0)
            db.Close()

            Dim i As Integer
            Dim RoleFlag1, RoleFlag2, RoleFlag3 As String
            RoleFlag1 = ""
            RoleFlag2 = ""
            RoleFlag3 = ""
            For i = 0 To RoleDT.Rows.Count - 1
                If RoleDT.Rows(i).Item("Group_Uid") = "GETK539lC0" Then
                    '系統管理員
                    RoleFlag1 = "1"
                ElseIf RoleDT.Rows(i).Item("Group_Uid") = "LN716924CA" Then
                    '單位管理員
                    RoleFlag2 = "1"
                Else
                    '其他權限
                    RoleFlag3 = "1"
                End If

            Next

            If RoleFlag1 = "1" Then
                Session("Role") = "1"
            Else
                If RoleFlag2 = "1" Then
                    Session("Role") = "2"
                Else
                    Session("Role") = "3"
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class
