
Partial Class Source_00_MOA00004
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN

    Public eformid, frm_chinese_name, organization_id As String
    Dim sSqlcommand, nodecount, leaderuid, leader_name As String
    Dim rsNodes, rsnodecount, rs, objFlowGroup, obj, objUniversal_r As String
    'Dim ReportData()
    'Dim aMemberArray
    Dim aMemberArrayChild(3)
    Dim bArrayIsEmpty

    'variables declaration for FlowChart
    Dim sFlowTable, sCCTable, sEformTable, NodesEndPosX, rsTempCC, cclist, rsGroupName, NextStep, rsStep, minx, maxx As String
    Dim bDefineReportData As String
    Dim startX, endX, minnextstep, maxnextstep As String

    Dim bIsProjectOrganization As String
    Dim object_uid, object_name As String
    Dim sLeaderName As String
    Public eformrole As String
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA00001") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00004.aspx")
                Response.End()
            End If

            'eformid = "4rM2YFP73N,MOA,會議室申請單,1" '(Request("eformid"))
            eformid = Request("eformid")
            organization_id = Request("organization_id")
            frm_chinese_name = Request("frm_chinese_name")
            eformrole = Request("eformrole")

        Catch ex As Exception

        End Try
    End Sub
End Class
