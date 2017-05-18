Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_07_MOA07001
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Public eformsn As String
    Public d_pub As New C_Public
    Dim eformid As String
    Dim str_ORG_UID As String
    Public stmt As String
    Dim p As Integer = 0
    Dim K As Integer = 0
    Dim str_app_id As String
    Public read_only As String = ""
    Dim connstr, user_id, org_uid As String
    Dim AgentEmpuid As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

        '新增資料
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        do_sql.G_errmsg = ""
        do_sql.G_user_id = Session("user_id") '"tempu180"
        'user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        eformsn = Request("eformsn") '"VJ" 'd_pub.randstr(2) '
        eformid = Request("eformid")
        read_only = Request("read_only")

        'session被清空回首頁
        If user_id = "" And org_uid = "" And eformsn = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            'do_sql.G_user_id = "tempu178"
            'eformsn = "R5Y26NOFA5123456" 'd_pub.randstr(16) '
            'eformid = "R5Y26NOFA5"
            'read_only = "2"
            If read_only = "2" Then
                but_exe.Text = "核准"
                If Session("user_id") = "" Then
                    do_sql.G_user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                End If
            End If

            '找出表單審核者
            If read_only = "2" Then

                '開啟連線
                Dim db As New SqlConnection(connstr)

                db.Open()
                Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' and hddate is null", db)
                Dim RdrAgentCheck = strAgentCheck.ExecuteReader()
                If RdrAgentCheck.read() Then
                    AgentEmpuid = RdrAgentCheck.item("empuid")
                End If
                db.Close()

            End If

            Lab_eformsn.Text = eformsn

            If IsPostBack Then
                Exit Sub
            End If
            'eformsn = d_pub.randstr(2) 'Request("eformsn") '

            If read_only = "1" Or read_only = "2" Then
                If select_data() = False Then
                    Exit Sub
                End If
                Exit Sub
            End If

            Label8.Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
            If do_sql.select_urname(do_sql.G_user_id) = False Then
                Lab_PWUNIT.Text = ""
                Exit Sub
            End If
            If do_sql.G_usr_table.Rows.Count > 0 Then
                Lab_PWUNIT.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
                Lab_PWNAME.Text = do_sql.G_usr_table.Rows(0).Item("emp_chinese_name").ToString.Trim
                Lab_PWTITLE.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
                str_ORG_UID = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
            End If
            If str_ORG_UID <> "" Then
                'stmt = "select * from Employee where ORG_UID='" + str_ORG_UID + "' order by emp_chinese_name"
                stmt = "select * from Employee where ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") order by emp_chinese_name"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    n_table = do_sql.G_table
                    p = 0
                    K = 0
                    DrDown_PANAME.Items.Clear()

                    For Each dr In n_table.Rows
                        DrDown_PANAME.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_PANAME.Items(p).Value = Trim(dr("employee_id").ToString)
                        If UCase(do_sql.G_user_id) = UCase(Trim(dr("employee_id").ToString)) Then
                            K = p
                        End If
                        p += 1
                    Next
                    If p > 0 Then
                        DrDown_PANAME.SelectedIndex = K
                        Call DrDown_PANAME_SelectedIndexChanged(sender, e)
                    End If
                End If
            End If

            Call Add_DropDownList()

            GridView1.DataBind()

        End If

    End Sub

    Protected Sub DrDown_PANAME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_PANAME.SelectedIndexChanged
        str_app_id = ""
        If do_sql.select_urname(DrDown_PANAME.Items(DrDown_PANAME.SelectedIndex).Value) = False Then
            DrDown_PANAME.Text = ""
            Exit Sub
        End If

        If do_sql.G_usr_table.Rows.Count > 0 Then
            Dim stmt As String = ""
            Dim CHINE_NAME As String = ""
            p = 0
            Lab_PAUNIT.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
            Lab_PATITLE.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
            str_ORG_UID = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
            str_app_id = do_sql.G_usr_table.Rows(0).Item("employee_id").ToString.Trim
        End If
    End Sub
    Sub Add_DropDownList()
        '事由

        stmt = "select Kind_Name,State_Name from SYSKIND where Kind_Num in(7,8,9,10)"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        n_table = do_sql.G_table
        DrDown_nREASON.Items.Clear()
        p = 0
        DrDown_nREASON.Items.Add("--請選擇--")
        DrDown_nREASON.Items(p).Value = ""
        p += 1
        For Each dr In n_table.Rows
            DrDown_nREASON.Items.Add(Trim(dr("Kind_Name").ToString) + "--" + Trim(dr("State_Name").ToString))
            DrDown_nREASON.Items(p).Value = Trim(dr("State_Name").ToString)
            p += 1
        Next
        '標籤
        stmt = "select Kind_Name,State_Name from SYSKIND where Kind_Num =11"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        n_table = do_sql.G_table
        DrDown_nLabel.Items.Clear()
        p = 0
        DrDown_nLabel.Items.Add("--請選擇--")
        DrDown_nLabel.Items(p).Value = ""
        p += 1
        For Each dr In n_table.Rows
            DrDown_nLabel.Items.Add(Trim(dr("State_Name").ToString))
            DrDown_nLabel.Items(p).Value = Trim(dr("State_Name").ToString)
            p += 1
        Next


        '數量 
        'p = 0
        'DrDown_nAmount.Items.Add("--請選擇--")
        'DrDown_nAmount.Items(p).Value = ""

        'For p = 1 To 10
        '    DrDown_nAmount.Items.Add(p.ToString.Trim)
        '    DrDown_nAmount.Items(p).Value = p.ToString.Trim
        'Next

    End Sub

    Protected Sub But_ins_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_ins.Click
        If ins_p701("1") = False Then
            Exit Sub
        End If
    End Sub
    Function ins_p701(ByVal df As String) As Boolean
        ins_p701 = False

        If Txt_nAssetNum.Text = "" Then
            If df = "2" Then
                stmt = "select * from p_0701 "
                stmt += "where EFORMSN ='" + eformsn + "'"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Function
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    ins_p701 = True
                    Exit Function
                End If
                'Exit Function
            End If
            do_sql.G_errmsg = "財產編號必須輸入"
            Exit Function
        End If
        If Txt_nBios.Text = "" Then
            do_sql.G_errmsg = "Bios密碼必須輸入"
            Exit Function
        End If
        If Txt_nAssetName.Text = "" Then
            do_sql.G_errmsg = "財產名稱必須輸入"
            Exit Function
        End If
        If DrDown_nLabel.SelectedItem.Value = "" Then
            do_sql.G_errmsg = "標籤必須輸入"
            Exit Function
        End If
        If Txt_nLabelNum.Text = "" Then
            do_sql.G_errmsg = "標籤號碼必須輸入"
            Exit Function
        End If
        If DrDown_nREASON.SelectedItem.Value = "" Then
            do_sql.G_errmsg = "問題類別必須輸入"
            Exit Function
        End If
        If TXT_nContent.Text = "" Then
            do_sql.G_errmsg = "說明要必須輸入"
            Exit Function
        End If



        stmt = "insert into P_0701(EFORMSN," '表單序號
        stmt += "nAssetNum,nAssetName,nBios," '財產編號,財產名稱,Bios密碼,
        stmt += "nLabel,nLabelNum,"  '標籤,標籤號碼
        stmt += "nREASON,nAmount,nContent)" '問題類別,數量,說明 "        
        stmt += " values('" + eformsn + "','" + Txt_nAssetNum.Text + "','" + Txt_nAssetName.Text + "','" + Txt_nBios.Text + "','"
        stmt += DrDown_nLabel.SelectedItem.Value + "','" + Txt_nLabelNum.Text + "','"
        stmt += DrDown_nREASON.SelectedItem.Value + "',1,'" + TXT_nContent.Text + "')"

        If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        '新增Grid資料:
        'stmt = "select * from p_0601 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        ''stmt += " order by ncreatedate"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'GridView1.DataSource = do_sql.G_table

        GridView1.DataBind()
        DrDown_nREASON.SelectedIndex = 0
        Txt_nAssetNum.Text = ""
        Txt_nAssetName.Text = ""
        Txt_nBios.Text = ""
        Txt_nLabelNum.Text = ""
        TXT_nContent.Text = ""
        DrDown_nLabel.SelectedIndex = 0
        'End If
        ins_p701 = True
    End Function

    Protected Sub But_exe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_exe.Click

        Dim FC As New C_FlowSend.C_FlowSend

        Dim SendVal As String = ""

        '判斷是否為代理人批核的表單
        If AgentEmpuid = "" Then
            SendVal = eformid & "," & user_id & "," & eformsn & "," & "1" & ","
        Else
            SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1" & ","
        End If

        Dim NextPer As Integer = 0

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '關卡為上一級主管有多少人
        NextPer = FC.F_NextStep(SendVal, connstr)

        Dim Chknextstep As Integer

        '判斷表單關卡
        db.Open()
        Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & eformsn & "' and empuid = '" & user_id & "' and hddate is null", db)
        Dim RdrCheck = strComCheck.ExecuteReader()
        If RdrCheck.read() Then
            Chknextstep = RdrCheck.item("nextstep")
        End If
        db.Close()

        Dim strgroup_id As String = ""

        If read_only <> "" Then

            '判斷上一級主管
            db.Open()
            Dim strUpComCheck As New SqlCommand("SELECT group_id FROM flow WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ", db)
            Dim RdrUpCheck = strUpComCheck.ExecuteReader()
            If RdrUpCheck.read() Then
                strgroup_id = RdrUpCheck.item("group_id")
            End If
            db.Close()

        End If

        If NextPer = 0 And read_only = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('無上一級主管');")
            Response.Write(" </script>")

            Exit Sub

        End If

        do_sql.G_errmsg = ""
        If read_only <> "2" Then
            If check_text() = False Then
                Exit Sub
            End If
            stmt = "insert into P_07(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
            stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
            stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
            stmt += "nAPPTIME,nTel,nSeat) " '申請時間,軍線電話,儲位
            stmt += " values('" + eformsn + "','" + Lab_PWUNIT.Text + "','" + Lab_PWTITLE.Text + "',N'"
            stmt += Lab_PWNAME.Text + "','" + do_sql.G_user_id + "','" + Lab_PAUNIT.Text + "',N'"
            stmt += DrDown_PANAME.SelectedItem.Text + "','" + Lab_PATITLE.Text + "','" + DrDown_PANAME.SelectedItem.Value + "','"
            stmt += Label8.Text + "','" + TXT_nTel.Text.Replace("'", "''") + "','" + Txt_nSeat.Text.Replace("'", "''") + "') "
            If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
        End If

        If read_only = "2" Then

            Dim strAgentName As String = ""

            '判斷是否為代理人批核
            If UCase(user_id) = UCase(AgentEmpuid) Then

                '增加批核意見
                insComment(txtcomment.Text, eformsn, user_id)

            Else

                Dim strComment As String = ""

                '找尋批核者姓名
                db.Open()
                Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    strAgentName = RdPer("emp_chinese_name")
                End If
                db.Close()

                strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)"

                '增加批核意見
                insComment(strComment, eformsn, AgentEmpuid)

            End If

        End If

        Dim Val_P As String

        Val_P = ""

        '判斷下一關為上一級主管時人數是否超過一人
        If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
            Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=1")
        Else
            '表單審核
            Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)
            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P)
            do_sql.G_errmsg = "新增成功"
        End If

    End Sub
    Function check_text() As Boolean
        check_text = False
        If TXT_nTel.Text = "" Then
            do_sql.G_errmsg = "軍線必須輸入"
            Exit Function
        End If
        If Txt_nSeat.Text = "" Then
            do_sql.G_errmsg = "儲位必須輸入"
            Exit Function
        End If
        check_text = True
    End Function
    Function select_data() As Boolean
        select_data = False
        Call Add_DropDownList()
        stmt = "select * from p_07 "
        stmt += "where EFORMSN ='" + eformsn + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        If do_sql.G_table.Rows.Count = 0 Then
            Lab_PWUNIT.Text = ""
            Lab_PWNAME.Text = ""
            Lab_PWTITLE.Text = ""
            Lab_PAUNIT.Text = ""
            DrDown_PANAME.Items.Clear()
            DrDown_PANAME.Items.Add("")
            DrDown_PANAME.Items(0).Value = ""

            Lab_PATITLE.Text = ""
            Label8.Text = ""
            TXT_nTel.Text = ""
            Txt_nSeat.Text = ""


            DrDown_nREASON.SelectedIndex = 0
            Txt_nAssetNum.Text = ""
            Txt_nAssetName.Text = ""
            Txt_nBios.Text = ""
            Txt_nLabelNum.Text = ""
            TXT_nContent.Text = ""
            DrDown_nLabel.SelectedIndex = 0
        Else
            Lab_PWUNIT.Text = do_sql.G_table.Rows(0).Item("PWUNIT").ToString
            Lab_PWNAME.Text = do_sql.G_table.Rows(0).Item("PWNAME").ToString
            Lab_PWTITLE.Text = do_sql.G_table.Rows(0).Item("PWTITLE").ToString
            Lab_PAUNIT.Text = do_sql.G_table.Rows(0).Item("PAUNIT").ToString
            DrDown_PANAME.Items.Clear()
            DrDown_PANAME.Items.Add(Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString))
            DrDown_PANAME.Items(0).Value = Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString)

            Lab_PATITLE.Text = do_sql.G_table.Rows(0).Item("PATITLE").ToString
            Label8.Text = do_sql.G_table.Rows(0).Item("nAPPTIME").ToString
            TXT_nTel.Text = do_sql.G_table.Rows(0).Item("nTel").ToString
            Txt_nSeat.Text = do_sql.G_table.Rows(0).Item("nSeat").ToString

        End If
        'stmt = "select * from p_0601 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        ''stmt += " order by ncreatedate"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'GridView1.DataSource = do_sql.G_table

        GridView1.DataBind()
        Call disabel_field()
        select_data = True
    End Function
    Sub disabel_field()
        Txt_nAssetNum.Enabled = False
        Txt_nAssetName.Enabled = False
        DrDown_nLabel.Enabled = False
        DrDown_nREASON.Enabled = False
        TXT_nContent.Enabled = False
        Txt_nBios.Enabled = False
        Txt_nLabelNum.Enabled = False

        GridView1.Enabled = False
        If read_only = "1" Then
            But_exe.Visible = False
            But_ins.Visible = False
        ElseIf read_only = "2" Then
            But_ins.Visible = False
        End If
    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '單純讀取表單不可送件
        If read_only = "1" Then
            But_exe.Visible = False
            backBtn.Visible = False
            tranBtn.Visible = False

            '讀取表單可送件
        ElseIf read_only = "2" Then

            Dim PerCount As Integer = 0

            '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
            Dim ParentFlag As String = ""
            Dim ParentVal = user_id & "," & eformsn

            Dim FC As New C_FlowSend.C_FlowSend
            ParentFlag = FC.F_NextChief(ParentVal, connstr)

            If ParentFlag = "1" Then
                'but_exe.Visible = False

                '上一級主管多少人
                db.Open()
                Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                Dim PerRdv = PerCountCom.ExecuteReader()
                If PerRdv.read() Then
                    PerCount = PerRdv("PerCount")
                End If
                db.Close()

                '上一級沒人則呈現送件按鈕
                If PerCount = 0 Then
                    But_exe.Visible = True
                    tranBtn.Visible = False
                Else
                    But_exe.Visible = True
                    'But_exe.Visible = False
                End If

            End If

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            '沒有批核權限不可執行動作
            If (Org_Down.ApproveAuth(eformsn, user_id)) = "" Then
                But_exe.Visible = False
                backBtn.Visible = False
                tranBtn.Visible = False
            End If

        Else
            backBtn.Visible = False
            tranBtn.Visible = False
        End If

    End Sub

    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles backBtn.Click

        Try

            If read_only = "2" Then
                '增加批核意見
                insComment(txtcomment.Text, eformsn, user_id)
            End If

            Dim Val_P As String
            Val_P = ""

            '表單駁回
            Dim FC As New C_FlowSend.C_FlowSend

            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & eformsn & "," & "1"
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1"
            End If

            Val_P = FC.F_Back(SendVal, connstr)

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單已駁回給申請人');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub tranBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tranBtn.Click
        Try
            If GetSuperiors(user_id).Rows.Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                Server.Transfer("../00/MOA00014.aspx?eformsn=" + eformsn)
            Else
                If read_only = "2" Then
                    '增加批核意見
                    insComment(txtcomment.Text, eformsn, user_id)
                End If

                Dim Val_P As String
                Val_P = ""

                '表單呈轉
                Dim FC As New C_FlowSend.C_FlowSend

                Dim SendVal As String = ""

                '判斷是否為代理人批核的表單
                If AgentEmpuid = "" Then
                    SendVal = eformid & "," & user_id & "," & eformsn & "," & "1"
                Else
                    SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1"
                End If

                Val_P = FC.F_Transfer(SendVal, connstr)

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('表單已呈轉給上一級主管');")
                '重新整理頁面
                Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
                Response.Write(" window.close();")
                Response.Write(" </script>")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Function GetSuperiors(ByVal employee_id As String) As DataTable
        Dim db As New SqlConnection(connstr)
        Dim ds As New DataSet()

        db.Open()
        Dim comm As New SqlCommand("select employee_id,emp_chinese_name from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id = @employee_id))", db)
        comm.Parameters.Add("@employee_id", Data.SqlDbType.VarChar, 10).Value = employee_id.Trim()
        Dim da As New SqlDataAdapter(comm)
        da.Fill(ds)
        db.Close()
        Return ds.Tables(0)
    End Function

    Public Function insComment(ByVal strComment As String, ByVal strEformsn As String, ByVal strid As String)

        '新增批核意見
        Try

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '新增批何意見
            db.Open()
            Dim insCom As New SqlCommand("UPDATE flowctl SET comment='" & strComment & "' WHERE hddate IS NULL AND eformsn='" & strEformsn & "' AND empuid='" & strid & "'", db)

            insCom.ExecuteNonQuery()
            db.Close()

        Catch ex As Exception
        End Try

        insComment = ""

    End Function
End Class
