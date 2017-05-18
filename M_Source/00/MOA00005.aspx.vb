
Partial Class Source_00_MOA00005
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN

    Public eformid, frm_chinese_name, org_chinese_name, organization_id As String
    Public bFlowDesignMode, bDesignColumnMode As Boolean
    Public sSqlcommand, leaderuid, leader_name As String
    Public nodecount As String
    Public rsNodes, rsnodecount, rs, objFlowGroup, obj As String
    'variables declaration for FlowChart
    Public sFlowTable, sCCTable, sEformTable, NodesEndPosX, rsTempCC, cclist, rsGroupName, NextStep, rsStep, minx, maxx As String
    Public startX, endX, minnextstep, maxnextstep As String
    Public scrollLeftPixel, scrollTopPixel As Integer
    Public save_flag As String = "False"
    Public eformrole As String
    Public err_msg As String = ""
    Dim user_id, org_uid As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            'eformid = "4rM2YFP73N,MOA,會議室申請單,2" '(Request("eformid"))

            eformid = Request("eformid")
            organization_id = Request("organization_id")
            frm_chinese_name = Request("frm_chinese_name")
            eformrole = Request("eformrole")
            err_msg = Request("err_msg")

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            If eformid = "" Then

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00005.aspx")
                Response.End()

            Else

                'Dim GetValue
                'GetValue = Split(eformid, ",")
                'eformid = GetValue(0)
                'organization_id = GetValue(1)
                'frm_chinese_name = GetValue(2)
                'eformrole = GetValue(3)

                scrollLeftPixel = CInt("0" & Request("scrollLeftPixel"))
                scrollTopPixel = CInt("0" & Request("scrollTopPixel"))
                If Request("save") = "TRUE" Then '可修改
                    save_flag = "True"
                End If
                If Request("FlowDesignMode") = "" Then
                    bFlowDesignMode = True
                Else
                    bFlowDesignMode = False
                End If

                'If Request("table_flag") <> "" Then
                '    table_flag = Request("table_flag")
                '    bFlowDesignMode = False
                'End If
                bDesignColumnMode = False

            End If




        Catch ex As Exception

        End Try

    End Sub
    End Class
