
Partial Class Source_00_MOA00003
    Inherits System.Web.UI.Page
    Public eformid As String
    Public GetValue
    Public organization_id As String
    Public frm_chinese_name As String
    Public eformrole As String
    Public table_flag As Boolean = False
    Dim user_id, org_uid As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA00001") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00003.aspx")
                Response.End()
            End If

            'eformid = "JH716YMK2U,MOA,會議室申請單,1" '(Request("eformid"))
            eformid = (Request("eformid"))
            GetValue = Split(eformid, ",")
            eformid = GetValue(0)
            organization_id = GetValue(1)
            frm_chinese_name = GetValue(2)
            eformrole = GetValue(3)

        Catch ex As Exception

        End Try

    End Sub
End Class
