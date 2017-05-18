Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_00_MOA00056
    Inherits System.Web.UI.Page

    Public OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lbMsg.Text = String.Empty
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

                If LoginCheck.LoginCheck(user_id, "MOA00056") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00056.aspx")
                    Response.End()
                End If

                If IsPostBack = False Then

                    '找出上一級單位以下全部單位
                    Dim Org_Down As New C_Public

                    '判斷登入者權限
                    If Session("Role") = "1" Then
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = ""
                    ElseIf Session("Role") = "2" Then
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(Session("PARENT_ORG_UID")) & ") ORDER BY ORG_NAME"
                    Else
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID ='" & Session("ORG_UID") & "' ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE employee_id ='" & Session("User_id") & "'"
                    End If

                End If

            End If

        Catch ex As Exception
            lbMsg.Text = "很抱歉! 系統發生錯誤，請連絡系統相關人員處理，謝謝!!"
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

    Protected Sub ImgOK_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgOK.Click

        Try
            Dim struser_id As String = ""

            struser_id = UserSel.SelectedValue

            If struser_id = "" Then
                ErrName.Visible = True
            Else
                ErrName.Visible = False

                '新增資料
                Dim connstr As String
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '先刪除個人權限
                db.Open()
                Dim sOrgSel As String = OrgSel.SelectedValue.ToString()
                Dim delCom As New SqlCommand("DELETE ROLES WHERE Group_Uid = '" & struser_id & "'", db)
                delCom.ExecuteNonQuery()
                db.Close()

                '刪除人員的影印管理使用者資料，人員有異動皆需做該人員的影印機管理的重設作業
                db.Open()
                Dim delPrinterCom As New SqlCommand("DELETE FROM P_0803 WHERE (employee_id = '" & struser_id & "')", db)
                delPrinterCom.ExecuteNonQuery()
                db.Close()

                Dim j As Integer
                '全部權限
                For j = 0 To ChkLimit.Items.Count - 1
                    If ChkLimit.Items(j).Selected = True Then
                        '新增權限資料
                        db.Open()

                        Dim insCom As New SqlCommand("INSERT INTO ROLES(Group_Uid,LinkNum) VALUES ('" & struser_id & "','" & ChkLimit.Items(j).Value & "')", db)
                        insCom.ExecuteNonQuery()

                        db.Close()
                        lbMsg.Text = "個人權限修改成功！"
                    End If
                Next
            End If
        Catch ex As Exception
            lbMsg.Text = "個人權限修改失敗，請重新操作或連絡系統相關人員！"
        End Try

    End Sub

    Protected Sub ChkLimit_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkLimit.PreRender

        '新增資料
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '找出個人權限
        Dim MeetDT As DataTable
        db.Open()

        Dim StepSQL As String = ""
        Dim i, j As Integer

        If UserSel.SelectedValue = "" Then
            '清空全部個人權限
            For j = 0 To ChkLimit.Items.Count - 1
                ChkLimit.Items(j).Selected = False
            Next
        Else
            'StepSQL = "SELECT RGI.Role_Num,RGI.Group_Uid,RGI.employee_id,R.LinkNum FROM [moa].[dbo].[ROLEGROUPITEM] AS RGI"
            'StepSQL = StepSQL + " JOIN ROLES AS R ON RGI.Group_Uid = R.Group_Uid WHERE employee_id = '" + UserSel.SelectedValue + "'"
            StepSQL = "select * from roles where group_uid='" & UserSel.SelectedValue & "' order by linknum desc"
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
            da.Fill(ds)
            MeetDT = ds.Tables(0)
            db.Close()

            '清空全部個人權限
            For j = 0 To ChkLimit.Items.Count - 1
                ChkLimit.Items(j).Selected = False
            Next
            If MeetDT.Rows.Count = 0 Then

            Else
                '判斷個人權限是否有加入
                For i = 0 To MeetDT.Rows.Count - 1
                    '全部個人權限
                    For j = 0 To ChkLimit.Items.Count - 1
                        If MeetDT.Rows(i).Item("LinkNum") = ChkLimit.Items(j).Value Then
                            ChkLimit.Items(j).Selected = True
                        End If
                    Next
                Next
            End If
        End If
    End Sub

    Protected Sub UserSel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles UserSel.SelectedIndexChanged

        '新增資料
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        Dim strGroup_Uid As String = ""

        '判斷是否有個人權限,沒有則新增
        db.Open()
        Dim strPer As New SqlCommand("SELECT Group_Uid FROM ROLEGROUPITEM WHERE Group_uid = '" & UserSel.SelectedValue & "'", db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.read() Then
            strGroup_Uid = RdPer("Group_Uid")
        End If
        db.Close()

        ''將資料庫中該人員已勾選的功能先勾選顯示出來，避免還要全部重選重勾
        'Dim StepSQL As String = "select LinkNum from ROLES where Group_uid = '" & UserSel.SelectedValue & "'"

        'Dim ds As New DataSet
        'Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
        'da.Fill(ds)
        'Dim MeetDT As DataTable = ds.Tables(0)
        'db.Close()

        'If MeetDT.Rows.Count = 0 Then
        '    '清空全部個人權限
        '    For j As Int16 = 0 To ChkLimit.Items.Count - 1
        '        ChkLimit.Items(j).Selected = False
        '    Next
        'Else
        '    '判斷個人權限是否有加入
        '    For i As Int16 = 0 To MeetDT.Rows.Count - 1
        '        '全部個人權限
        '        For j As Int16 = 0 To ChkLimit.Items.Count - 1
        '            If MeetDT.Rows(i).Item("LinkNum") = ChkLimit.Items(j).Value Then
        '                ChkLimit.Items(j).Selected = True
        '            End If
        '        Next
        '    Next
        'End If

    End Sub
End Class
