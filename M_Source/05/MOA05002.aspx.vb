Imports System.Data.SqlClient
Imports System.Data
Imports System.Threading

Partial Class Source_05_MOA05002
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Dim connstr As String = ""
    Dim VistorFlag As String = ""
    Public Shared OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '新增連線
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷登入者是否為會客相關群組人員
            db.Open()
            Dim carCom As New SqlCommand("select * from ROLEGROUPITEM where employee_id = '" & user_id & "' AND (Group_Uid ='2I59557MG6')", db)
            Dim RdvCar = carCom.ExecuteReader()
            If RdvCar.read() Then
                VistorFlag = "1"
            End If
            db.Close()

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            '判斷登入者權限
            If Session("Role") = "1" Or VistorFlag = "1" Then
                '會客室管理者以及系統管理者可以查看全部會客內容
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
            Else
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") ORDER BY ORG_NAME"
            End If

            '先設定起始日期
            Dim dt As Date = Now()
            If (nRECDATE1.Text = "") Then
                nRECDATE1.Text = dt.Date
            End If

            If (nRECDATE2.Text = "") Then
                nRECDATE2.Text = dt.AddDays(7).Date
            End If

        End If


    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Server.Transfer("../00/MOA00020.aspx?x=MOA05001&y=U28r13D6EA&Read_Only=1&EFORMSN=" & GridView1.SelectedValue)
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)
        If (sort) Then
            GridView1.Sort("nRECDATE", SortDirection.Descending)
            Return
        End If

        Dim sql As String = ""
        Dim ParentFlag As String = ""

        sql = "SELECT distinct P_NUM, convert(nvarchar,nRECDATE,111) nRECDATE, nSTARTTIME+'~'+nENDTIME nSTARTTIME "
        sql += ", nREASON, PANAME, P_05.EFORMSN EFORMSN, PENDFLAG, nRECROOM, ADMINGROUP.ORG_NAME AS PAUNIT FROM P_05 "
        sql += "JOIN EMPLOYEE ON P_05.PAIDNO = EMPLOYEE.employee_id "
        sql += "JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID "
        sql += "CROSS JOIN P_0501 where P_05.EFORMSN=P_0501.EFORMSN "

        '會客室管理者以及系統管理者可以查看全部會客內容
        If Session("Role") = "1" Or VistorFlag = "1" Then

            '申請人單位
            If PAUNIT.SelectedItem.Text = "請選擇" Then
                sql += ""
            Else
                If InStr(PAUNIT.SelectedItem.Text, "(") > 0 Then
                    sql += " AND PAUNIT LIKE '%" & Left(PAUNIT.SelectedItem.Text, InStr(PAUNIT.SelectedItem.Text, "(") - 1) & "%'"
                Else
                    sql += " AND PAUNIT = '" & PAUNIT.SelectedItem.Text & "'"
                End If
            End If

            '人員
            If PAIDNO.Text <> "" Then
                sql += " AND PAIDNO = '" & PAIDNO.SelectedValue & "'"
            End If

        Else

            Dim connstr As String = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷是否有下一級單位
            db.Open()
            Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                ParentFlag = "Y"
            End If
            db.Close()

            Dim UserOrgName As String = ""

            '搜尋登入者單位名稱
            db.Open()
            Dim strOrgName As New SqlCommand("SELECT ORG_NAME FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "'", db)
            Dim RdOrgName = strOrgName.ExecuteReader()
            If RdOrgName.read() Then
                UserOrgName = RdOrgName("ORG_NAME")
            End If
            db.Close()

            '主官管
            If ParentFlag = "Y" Then
                '組織搜尋
                If PAUNIT.SelectedItem.Text = "請選擇" Then
                    If InStr(UserOrgName, "(") > 0 Then
                        sql += " AND PAUNIT LIKE '%" & Left(UserOrgName, InStr(UserOrgName, "(") - 1) & "%'"
                    Else
                        sql += " AND PAUNIT = '" & UserOrgName & "'"
                    End If
                Else
                    If InStr(PAUNIT.SelectedItem.Text, "(") > 0 Then
                        sql += " AND PAUNIT LIKE '%" & Left(PAUNIT.SelectedItem.Text, InStr(PAUNIT.SelectedItem.Text, "(") - 1) & "%'"
                    Else
                        sql += " AND PAUNIT = '" & PAUNIT.SelectedItem.Text & "'"
                    End If
                End If

                '人員
                If PAIDNO.Text <> "" Then
                    '人員
                    sql += " AND PAIDNO = '" & PAIDNO.SelectedValue & "'"
                End If
            Else
                '人員
                sql += " AND PAIDNO = '" & user_id & "'"

            End If

        End If


        'If (PAUNIT.Text <> "") Then sql += " and P_05.PAUNIT like '%" & Left(PAUNIT.Text, InStr(PAUNIT.Text, "(") - 1) & "%'" '申請人單位
        'If (PAIDNO.Text <> "") Then sql += " and P_05.PAIDNO like '%" + Trim(PAIDNO.SelectedValue) + "%'" '申請人帳號

        '申請日期搜尋
        If (nRECDATE1.Text <> "") Then sql += " AND (nRECDATE between '" & nRECDATE1.Text & "' AND '" & nRECDATE2.Text & "')" '會客時間
        'If (nRECDATE1.Text <> "") Then sql += " AND (nRECDATE between '" & nRECDATE1.Text & " 00:00:00 ' AND '" & nRECDATE2.Text & " 23:59:59')"
        'If (nRECDATE1.Text <> "") Then sql += " and P_05.nRECDATE>=CONVERT(datetime,'" + nRECDATE1.Text + " 00:00')" '會客時間
        'If (nRECDATE2.Text <> "") Then sql += " and P_05.nRECDATE<=CONVERT(datetime,'" + nRECDATE2.Text + " 23:59')" '會客時間
        If (nService.Text <> "") Then sql += " and P_0501.nService like '%" + Trim(nService.Text) + "%'" '服務單位
        If (nName.Text <> "") Then sql += " and P_0501.nName like '%" + Trim(nName.Text) + "%'" '姓名
        If (PANAME.Text <> "") Then sql += " and P_05.PANAME like '%" + Trim(PANAME.Text) + "%'" '申請人姓名
        If (nREASON.Text <> "") Then sql += " and P_05.nREASON like '%" + Trim(nREASON.Text) + "%'" '事由

        SqlDataSource2.SelectCommand = sql

        OrgChange = ""
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    Protected Sub PAUNIT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PAUNIT.SelectedIndexChanged
       
        '判斷組織是否變更
        OrgChange = "1"

        '清空User重新讀取
        PAIDNO.Items.Clear()

        If PAUNIT.SelectedValue = "" Then
            SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE 1=2"
        Else
            SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & PAUNIT.SelectedValue & "' ORDER BY emp_chinese_name"
        End If


    End Sub

    Protected Sub PAIDNO_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles PAIDNO.PreRender

        If OrgChange = "1" Then
            Dim tLItm As New ListItem("請選擇", "")

            '人員加請選擇
            PAIDNO.Items.Insert(0, tLItm)
            If PAIDNO.Items.Count > 1 Then
                PAIDNO.Items(1).Selected = False
            End If
        End If

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then

            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇
            PAUNIT.Items.Insert(0, tLItm)
            If PAUNIT.Items.Count > 1 Then
                PAUNIT.Items(1).Selected = False
            End If

            '人員加請選擇
            PAIDNO.Items.Insert(0, tLItm)
            If PAIDNO.Items.Count > 1 Then
                PAIDNO.Items(1).Selected = False
            End If

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click

        nRECDATE1.Text = ""
        nRECDATE2.Text = ""
        nService.Text = ""
        PANAME.Text = ""
        nName.Text = ""
        nREASON.Text = ""

    End Sub

    Protected Function FunStatus(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr = Eval(str)

            If tmpStr = "-" Then
                tmpStr = "申請"
            ElseIf tmpStr = "0" Then
                tmpStr = "駁回"
            ElseIf tmpStr = "1" Then
                tmpStr = "送件"
            ElseIf tmpStr = "?" Or tmpStr = "" Then
                tmpStr = "審核中"
            ElseIf tmpStr = "E" Then
                tmpStr = "完成"
            ElseIf tmpStr = "G" Then
                tmpStr = "補登"
            ElseIf tmpStr = "B" Then
                tmpStr = "撤銷"
            ElseIf tmpStr = "R" Then
                tmpStr = "重新分派"
            ElseIf tmpStr = "T" Then
                tmpStr = "呈轉"
            Else
                tmpStr = "未知"
            End If

            FunStatus = tmpStr

        Catch ex As Exception
            FunStatus = ""
        End Try
    End Function

    Protected Sub imbOutSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imbOutSearch.Click
        Try
            Dim strSql = "SELECT A.P_NUM,CONVERT(nvarchar,nRECDATE,111) nRECDATE,nSTARTTIME+'~'+nENDTIME nSTARTTIME,nREASON,PANAME,a.EFORMSN,PENDFLAG,nRECROOM,c.ORG_NAME PAUNIT FROM P_05 A LEFT JOIN EMPLOYEE B ON A.PAIDNO=B.EMPLOYEE_ID LEFT JOIN ADMINGROUP C ON B.ORG_UID = C.ORG_UID WHERE A.PENDFLAG='E' AND A.NCHECKDT IS NULL"

            Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                conn.Open()
                ''取出會客洽公申請單且未成功匯出之資料
                Dim cmd As SqlCommand
                cmd = New SqlCommand(strSql, conn)

                Dim DR As SqlDataReader
                DR = cmd.ExecuteReader
                Dim dt As New DataTable
                dt.Load(DR)
                If (dt.Rows.Count > 0) Then
                    ViewState("HasOutputData") = True
                    imbExport.Visible = True
                End If
            End Using
            SqlDataSource2.SelectCommand = strSql
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub

    Protected Sub imbExport_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imbExport.Click
        Try
            For Each GridViewRow In GridView1.Rows
                Select Case GridViewRow.RowType
                    Case DataControlRowType.DataRow
                        ''insert外部資料庫
                        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                            conn.Open()
                            Dim tran As SqlTransaction
                            tran = conn.BeginTransaction
                            ''會客洽公申請單
                            Dim P_05AData As New P_05A(GridViewRow.Cells(7).Text)							
                            P_05AData.Insert(tran, conn)                            
                            tran.Commit()
                        End Using
                        ''更新門禁欄位已匯出時間
                        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                            conn.Open()
                            Dim tran As SqlTransaction
                            tran = conn.BeginTransaction
                            Dim cmd As New SqlCommand("UPDATE P_05 SET nCheckDT=GetDate() WHERE EFORMSN='" & GridViewRow.Cells(7).Text & "'", conn, tran)
                            Dim iSuccess As Integer = cmd.ExecuteNonQuery
                            tran.Commit()
                        End Using
                End Select
            Next
            imbExport.Visible = False
            MessageBox.Show("資料匯出完成")
            Me.imbOutSearch_Click(sender, e)
        Catch sqlex As SqlException
            If (Not "-1".Equals(sqlex.Message.IndexOf("無法開啟至 SQL Server 的連接").ToString) Or Not "-1".Equals(sqlex.Message.IndexOf("登入失敗").ToString)) Then
                MessageBox.Show("門禁系統連線錯誤，請連繫管理員")
            Else
                MessageBox.Show(sqlex.Message)
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub

    Protected Sub GridView1_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Try
            If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
                '隱藏eformsn,eformid
                e.Row.Cells(7).Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class
