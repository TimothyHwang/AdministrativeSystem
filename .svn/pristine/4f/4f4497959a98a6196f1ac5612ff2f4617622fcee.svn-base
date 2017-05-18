Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_00_MOA00042
    Inherits System.Web.UI.Page

    Public OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid, CancelPer, connstr As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            CancelPer = Request.QueryString("CancelPer")

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00040") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00042.aspx")
                    Response.End()
                End If

                If IsPostBack = False Then

                    '找出上一級單位以下全部單位
                    Dim Org_Down As New C_Public

                    '找出登入者的一級單位
                    Dim strParentOrg As String = ""
                    strParentOrg = Org_Down.getUporg(org_uid, 1)

                    '判斷登入者權限
                    If Session("Role") = "1" Then
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = ""
                    ElseIf Session("Role") = "2" Then
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") ORDER BY ORG_NAME"
                    Else
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID ='" & Session("ORG_UID") & "' ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE employee_id ='" & user_id & "'"
                    End If

                End If



            End If


        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇
            If Session("Role") = "1" Then

                OrgSel.Items.Insert(0, tLItm)
                If OrgSel.Items.Count > 1 Then
                    OrgSel.Items(1).Selected = False
                End If

            End If

            '人員加請選擇
            UserSel.Items.Insert(0, tLItm)
            If UserSel.Items.Count > 1 Then
                UserSel.Items(1).Selected = False
            End If

            OrgChange = ""

        End If

    End Sub

    Protected Sub UserSel_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles UserSel.PreRender

        If OrgChange = "1" Then

            Dim tLItm As New ListItem("請選擇", "")

            '人員加請選擇
            UserSel.Items.Insert(0, tLItm)
            If UserSel.Items.Count > 1 Then
                UserSel.Items(1).Selected = False
            End If

        End If

    End Sub

    Protected Sub OrgSel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OrgSel.SelectedIndexChanged

        '判斷組織是否變更
        OrgChange = "1"

        '清空User重新讀取
        UserSel.Items.Clear()

    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA00043.aspx")

    End Sub

    Protected Sub ImgApply_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgApply.Click

        '重新分派人員
        Dim struser_id As String = ""

        struser_id = UserSel.SelectedValue

        If struser_id = "" Then
            ErrName.Visible = True
        Else
            ErrName.Visible = False

            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '找出移到待派區批核表單
            db.Open()
            Dim StepSQL As String = "SELECT flowsn FROM flowctl WHERE empuid = '" & CancelPer & "' AND hddate IS NULL "
            Dim ds As New DataSet
            Dim Dt As New DataTable
            Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
            da.Fill(ds)
            '設定DataTable
            Dt = ds.Tables(0)

            '表單重新分派
            Dim FC As New C_FlowSend.C_FlowSend

            For x As Integer = 0 To Dt.Rows.Count - 1

                Dim Val_P As String = ""

                Dim SendVal As String = struser_id & "," & Dt.Rows(x).Item(0)

                Val_P = FC.F_Assign(SendVal, connstr)

            Next
            db.Close()

            '將被重新分派的人員移至待派區
            If CancelPer <> "" Then

                '修改人員為待派區
                db.Open()

                Dim updCom As New SqlCommand("UPDATE EMPLOYEE SET ORG_UID='520',Leave='N' WHERE employee_id = '" & CancelPer & "'", db)
                updCom.ExecuteNonQuery()
                db.Close()

                '取消人員群組權限
                db.Open()
                Dim delCom As New SqlCommand("DELETE FROM ROLEGROUPITEM WHERE (employee_id = '" & CancelPer & "') AND (Group_Uid <> 'JROHEST34K')", db)
                delCom.ExecuteNonQuery()
                db.Close()

                '取消人員關卡權限
                db.Open()
                Dim delSTCom As New SqlCommand("DELETE FROM SYSTEMOBJUSE WHERE (employee_id = '" & CancelPer & "')", db)
                delSTCom.ExecuteNonQuery()
                db.Close()

                '刪除人員的影印管理使用者資料，人員有異動皆需做該人員的影印機管理的重設作業
                db.Open()
                Dim delPrinterCom As New SqlCommand("DELETE FROM P_0803 WHERE (employee_id = '" & CancelPer & "')", db)
                delPrinterCom.ExecuteNonQuery()
                db.Close()
            End If

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('重新分派完成!!!');")
            Response.Write(" </script>")

            Server.Transfer("MOA00040.aspx")


        End If


    End Sub
End Class
