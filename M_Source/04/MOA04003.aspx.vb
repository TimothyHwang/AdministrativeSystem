Imports System.Data.SqlClient
Partial Class Source_04_MOA04003
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
    Dim str_nConSume As String = ""
    Dim str_nEffect As String = ""
    Dim Str_nHandleDate As String
    Dim Str_nCheckDate As String
    Dim Str_nDoleDate As String
    Dim Str_nFinishDate As String
    Dim connstr, user_id, org_uid As String
    Dim AgentEmpuid As String = ""
    Dim AppUID As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

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

            '新增連線
            connstr = do_sql.G_conn_string

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            If read_only = "2" Then
                But_exe.Text = "核准"
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

            Call drow_txt()

            If add_DrDown_nFIXITEM() = False Then
                Exit Sub
            End If

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

    
    Function add_DrDown_nFIXITEM() As Boolean
        add_DrDown_nFIXITEM = False
        stmt = "select EFORMSN,PANAME,nAPPTIME,nPHONE,nPLACE,nFIXITEM from p_04 "
        stmt += " where not exists(select nEFORMSN from P_0401 where p_04.EFORMSN=P_0401.nEFORMSN and p_0401.PENDFLAG <> '0') and p_04.PENDFLAG='E' and nAPPTIME >= '2009/1/1' order by nAPPTIME"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            Dim st_value As String = ""
            n_table = do_sql.G_table
            p = 0
            K = 0
            DrDown_nFIXITEM.Items.Clear()

            For Each dr In n_table.Rows
                stmt = CDate(Trim(dr("nAPPTIME").ToString)).ToString("yyyy/MM/dd") + "--" + Trim(dr("PANAME").ToString)
                stmt += "--" + Trim(dr("nPHONE").ToString)
                stmt += "--" + Trim(dr("nPLACE").ToString)
                'stmt += "--" + Trim(dr("nFIXITEM").ToString)
                DrDown_nFIXITEM.Items.Add(stmt)
                DrDown_nFIXITEM.Items(p).Value = Trim(dr("EFORMSN").ToString)
                If p = 0 Then
                    st_value = Trim(dr("nFIXITEM").ToString)
                End If

                p += 1
            Next
            txtnFIXITEM.Text = st_value
        End If
        add_DrDown_nFIXITEM = True
    End Function
    Sub drow_txt()
        DrDown_nHandle_HOUR.Items.Clear()
        DrDown_nCheck_HOUR.Items.Clear()
        DrDown_nDole_HOUR.Items.Clear()
        DrDown_nFinish_HOUR.Items.Clear()
        For p = 8 To 17
            DrDown_nHandle_HOUR.Items.Add(p.ToString("00"))
            DrDown_nHandle_HOUR.Items(p - 8).Value = p.ToString("00")
            '
            DrDown_nCheck_HOUR.Items.Add(p.ToString("00"))
            DrDown_nCheck_HOUR.Items(p - 8).Value = p.ToString("00")
            '
            DrDown_nDole_HOUR.Items.Add(p.ToString("00"))
            DrDown_nDole_HOUR.Items(p - 8).Value = p.ToString("00")
            '
            DrDown_nFinish_HOUR.Items.Add(p.ToString("00"))
            DrDown_nFinish_HOUR.Items(p - 8).Value = p.ToString("00")
        Next
    End Sub

    Protected Sub But_ins_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_ins.Click
        If ins_p402() = False Then
            Exit Sub
        End If
    End Sub

    Protected Sub But_exe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_exe.Click

        Dim FC As New C_FlowSend.C_FlowSend

        Dim SendVal As String = ""

        '判斷是否為代理人批核的表單
        If AgentEmpuid = "" Then
            'SendVal = eformid & "," & user_id & "," & eformsn & "," & "1" & ","
            SendVal = eformid & "," & DrDown_PANAME.SelectedValue & "," & eformsn & "," & "1" & ","
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

            stmt = "insert into P_0401(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
            stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
            stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人帳號 "
            stmt += "nAPPTIME,nEFORMSN, nConSume, " '申請時間,房屋水電修繕單序號,材料,
            stmt += "nEffect,nHandleDate,nCheckDate, " '效果,交辦日期,勘查日期        
            stmt += "nDoleDate,nFinishDate) " '發料日期,完工日期        
            stmt += " values('" + eformsn + "','" + Lab_PWUNIT.Text + "','" + Lab_PWTITLE.Text + "',N'"
            stmt += Lab_PWNAME.Text + "','" + do_sql.G_user_id + "','" + Lab_PAUNIT.Text + "',N'"
            stmt += DrDown_PANAME.SelectedItem.Text + "','" + Lab_PATITLE.Text + "','" + DrDown_PANAME.SelectedItem.Value + "','"
            stmt += Label8.Text + "','" + DrDown_nFIXITEM.SelectedItem.Value + "'," + str_nConSume + ","
            stmt += str_nEffect + ",'" + Str_nHandleDate + "','" + Str_nCheckDate + "','"
            stmt += Str_nDoleDate + "','" + Str_nFinishDate + "')"
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

        Dim Val_P As String = ""
        Val_P = ""

        '判斷下一關為上一級主管時人數是否超過一人
        If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
            Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=1" & "&AppUID=" & DrDown_PANAME.SelectedValue)
        Else

            '表單審核
            Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)
            Dim PageUp As String = ""

            If read_only = "" Then
                PageUp = "New"
            End If

            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
            do_sql.G_errmsg = "新增成功"

        End If



    End Sub
    Function ins_p402() As Boolean
        ins_p402 = False

        If Txt_Block_Name.Text = "" Then
            do_sql.G_errmsg = "品名必須輸入"
            Exit Function
        End If
        If Txt_Block_Unit.Text = "" Then
            do_sql.G_errmsg = "單位必須輸入"
            Exit Function
        End If
        If Txt_Block_Amount.Text = "" Then
            do_sql.G_errmsg = "數量必須輸入"
            Exit Function
        End If


        'stmt = "select * from p_0501 "
        'stmt += "where EFORMSN ='" + eformsn + "' and nID='" + Txt_nID.Text + "'"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'If do_sql.G_table.Rows.Count = 0 Then
        stmt = "insert into P_0402(EFORMSN," '完工表單序號
        stmt += "Block_Name,Block_Unit,Block_Amount)" '品名,單位,數量,
        stmt += " values('" + eformsn + "','" + Txt_Block_Name.Text + "','" + Txt_Block_Unit.Text + "','"
        stmt += Txt_Block_Amount.Text + "')"

        If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        ''新增Grid資料:
        GridView1.DataBind()
        Txt_Block_Name.Text = ""
        Txt_Block_Unit.Text = ""
        Txt_Block_Amount.Text = ""
        'End If
        ins_p402 = True
    End Function
    Function check_text() As Boolean
        check_text = False
        If Lab_PAUNIT.Text = "" Then
            do_sql.G_errmsg = "申請人姓名必須選取"
            Exit Function
        End If
        If DrDown_nFIXITEM.SelectedItem.Value = "" Then
            do_sql.G_errmsg = "申請單必須點選"
            Exit Function
        End If
        If Rad_nConSume1.Checked = True Then
            str_nConSume = "0"
        Else
            str_nConSume = "1"
        End If
        If Rad_nEffect1.Checked = True Then
            str_nEffect = "0"
        Else
            str_nEffect = "1"
        End If
        On Error GoTo err_check
        If Txt_nHandleDate.Text = "" Then
            do_sql.G_errmsg = "交辦日期必須輸入"
            Exit Function
        Else
            do_sql.G_errmsg = "交辦日期必須輸入"
            Txt_nHandleDate.Text = CDate(Txt_nHandleDate.Text).ToString("yyyy/MM/dd")
        End If
        If Txt_nCheckDate.Text = "" Then
            do_sql.G_errmsg = "勘查日期必須輸入"
            Exit Function
        Else
            do_sql.G_errmsg = "勘查日期必須輸入"
            Txt_nCheckDate.Text = CDate(Txt_nCheckDate.Text).ToString("yyyy/MM/dd")
        End If
        If Txt_nDoleDate.Text = "" Then
            do_sql.G_errmsg = "發料日期必須選取"
            Exit Function
        Else
            do_sql.G_errmsg = "發料日期必須輸入"
            Txt_nDoleDate.Text = CDate(Txt_nDoleDate.Text).ToString("yyyy/MM/dd")
        End If
        If Txt_nFinishDate.Text = "" Then
            do_sql.G_errmsg = "完工日期必須選取"
            Exit Function
        Else
            do_sql.G_errmsg = "完工日期必須輸入"
            Txt_nFinishDate.Text = CDate(Txt_nFinishDate.Text).ToString("yyyy/MM/dd")
            do_sql.G_errmsg = ""
            On Error GoTo err_check
        End If
        
        Str_nHandleDate = Txt_nHandleDate.Text + " " + DrDown_nHandle_HOUR.SelectedItem.Value + ":00:00"
        Str_nCheckDate = Txt_nCheckDate.Text + " " + DrDown_nCheck_HOUR.SelectedItem.Value + ":00:00"
        Str_nDoleDate = Txt_nDoleDate.Text + " " + DrDown_nDole_HOUR.SelectedItem.Value + ":00:00"
        Str_nFinishDate = Txt_nFinishDate.Text + " " + DrDown_nFinish_HOUR.SelectedItem.Value + ":00:00"

        stmt = "select * from p_0402 "
        stmt += "where EFORMSN ='" + eformsn + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            do_sql.G_errmsg = "SQL錯誤"
            Exit Function
        End If
        If do_sql.G_table.Rows.Count = 0 Then
            do_sql.G_errmsg = "請輸入物品名稱,單位,數量"
            Exit Function
        End If

        check_text = True
        Exit Function
err_check:
        Exit Function
    End Function
    Function select_data() As Boolean
        select_data = False
        Call drow_txt()
        stmt = "select * from p_0401 "
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

            DrDown_nFIXITEM.SelectedItem.Text = ""
            Txt_nHandleDate.Text = ""
            DrDown_nHandle_HOUR.SelectedIndex = 0
            DrDown_nCheck_HOUR.SelectedIndex = 0
            DrDown_nDole_HOUR.SelectedIndex = 0
            DrDown_nFinish_HOUR.SelectedIndex = 0
            Txt_nCheckDate.Text = ""
            Txt_nDoleDate.Text = ""
            Txt_nFinishDate.Text = ""
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

            'DrDown_nFIXITEM.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nFIXITEM").ToString
            If do_sql.G_table.Rows(0).Item("nConSume").ToString.Trim = "0" Then
                Rad_nConSume1.Checked = True
                Rad_nConSume2.Checked = False
            Else
                Rad_nConSume1.Checked = False
                Rad_nConSume2.Checked = True
            End If
            If do_sql.G_table.Rows(0).Item("nEffect").ToString.Trim = "0" Then
                Rad_nEffect1.Checked = True
                Rad_nEffect2.Checked = False
            Else
                Rad_nEffect1.Checked = False
                Rad_nEffect2.Checked = True
            End If
            Dim pt_date As String
            pt_date = do_sql.G_table.Rows(0).Item("nHandleDate").ToString.Trim
            If pt_date = "" Then
                Txt_nHandleDate.Text = ""
                DrDown_nHandle_HOUR.SelectedIndex = 0
            Else
                pt_date = CDate(pt_date).ToString("yyyy/MM/dd HH")
                Txt_nHandleDate.Text = Mid(pt_date, 1, 10)
                DrDown_nHandle_HOUR.SelectedItem.Text = Mid(pt_date, 12, 2)
            End If
            pt_date = do_sql.G_table.Rows(0).Item("nCheckDate").ToString.Trim
            If pt_date = "" Then
                Txt_nCheckDate.Text = ""
                DrDown_nCheck_HOUR.SelectedIndex = 0
            Else
                pt_date = CDate(pt_date).ToString("yyyy/MM/dd HH")
                Txt_nCheckDate.Text = Mid(pt_date, 1, 10)
                DrDown_nCheck_HOUR.SelectedItem.Text = Mid(pt_date, 12, 2)
            End If
            pt_date = do_sql.G_table.Rows(0).Item("nDoleDate").ToString.Trim
            If pt_date = "" Then
                Txt_nDoleDate.Text = ""
                DrDown_nDole_HOUR.SelectedIndex = 0
            Else
                pt_date = CDate(pt_date).ToString("yyyy/MM/dd HH")
                Txt_nDoleDate.Text = Mid(pt_date, 1, 10)
                DrDown_nDole_HOUR.SelectedItem.Text = Mid(pt_date, 12, 2)
            End If
            pt_date = do_sql.G_table.Rows(0).Item("nFinishDate").ToString.Trim
            If pt_date = "" Then
                Txt_nFinishDate.Text = ""
                DrDown_nFinish_HOUR.SelectedIndex = 0
            Else
                pt_date = CDate(pt_date).ToString("yyyy/MM/dd HH")
                Txt_nFinishDate.Text = Mid(pt_date, 1, 10)
                DrDown_nFinish_HOUR.SelectedItem.Text = Mid(pt_date, 12, 2)
            End If
            '----
            stmt = "select EFORMSN,PANAME,nAPPTIME,nPHONE,nPLACE,nFIXITEM from p_04 "
            stmt += "where EFORMSN ='" + do_sql.G_table.Rows(0).Item("nEFORMSN").ToString + "'"
            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Function
            End If
            If do_sql.G_table.Rows.Count > 0 Then
                Dim st_value As String = ""
                n_table = do_sql.G_table
                p = 0
                K = 0
                DrDown_nFIXITEM.Items.Clear()

                For Each dr In n_table.Rows
                    stmt = CDate(Trim(dr("nAPPTIME").ToString)).ToString("yyyy/MM/dd") + "--" + Trim(dr("PANAME").ToString)
                    stmt += "--" + Trim(dr("nPHONE").ToString)
                    stmt += "--" + Trim(dr("nPLACE").ToString)
                    'stmt += "--" + Trim(dr("nFIXITEM").ToString)
                    DrDown_nFIXITEM.Items.Add(stmt)
                    DrDown_nFIXITEM.Items(p).Value = Trim(dr("EFORMSN").ToString)
                    If p = 0 Then
                        st_value = Trim(dr("nFIXITEM").ToString)
                    End If
                    p += 1
                Next
                txtnFIXITEM.Text = st_value
            End If

        End If

        GridView1.DataBind()
        Call disabel_field()
        select_data = True
    End Function
    Sub disabel_field()

        DrDown_PANAME.Enabled = False
        DrDown_nFIXITEM.Enabled = False
        Rad_nConSume1.Enabled = False
        Rad_nConSume2.Enabled = False
        Rad_nEffect1.Enabled = False
        Rad_nEffect2.Enabled = False

        Txt_nHandleDate.Enabled = False
        DrDown_nHandle_HOUR.Enabled = False
        Txt_nCheckDate.Enabled = False
        DrDown_nCheck_HOUR.Enabled = False
        Txt_nDoleDate.Enabled = False
        DrDown_nDole_HOUR.Enabled = False
        Txt_nFinishDate.Enabled = False
        DrDown_nFinish_HOUR.Enabled = False

        GridView1.Enabled = False
        But_ins.Enabled = False

    End Sub
    'Function chg_Fdate(ByVal txt_value As String) As String
    '    chg_Fdate = txt_value
    '    If txt_value = "" Then
    '        Exit Function
    '    End If
    '    If txt_value.Length = 10 Then
    '        Exit Function
    '    End If
    '    chg_Fdate = CDate(txt_value).ToString("yyyy/MM/dd")
    'End Function
    Sub f1()
        If But_exe.Enabled = False Then
            But_exe.Enabled = True
        Else
            But_exe.Enabled = False
        End If
    End Sub

    Protected Sub Txt_nHandleDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_nHandleDate.TextChanged
        'Txt_nHandleDate.Text = "0019"
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

        Catch ex As Exception

        End Try

    End Sub

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

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "205px"
        Div_grid.Style("left") = "320px"

        'Calendar1.SelectedDate = Txt_nHandleDate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "205px"
        Div_grid2.Style("left") = "320px"

        'Calendar2.SelectedDate = Txt_nCheckDate.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Txt_nHandleDate.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        Txt_nCheckDate.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

    Protected Sub ImgDate3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate3.Click

        Div_grid3.Visible = True
        Div_grid3.Style("Top") = "205px"
        Div_grid3.Style("left") = "320px"

        'Calendar3.SelectedDate = Txt_nDoleDate.Text

    End Sub

    Protected Sub ImgDate4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate4.Click

        Div_grid4.Visible = True
        Div_grid4.Style("Top") = "205px"
        Div_grid4.Style("left") = "320px"

        'Calendar4.SelectedDate = Txt_nFinishDate.Text

    End Sub

    Protected Sub Calendar3_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar3.SelectionChanged

        Txt_nDoleDate.Text = Calendar3.SelectedDate.Date
        Div_grid3.Visible = False

    End Sub

    Protected Sub Calendar4_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar4.SelectionChanged

        Txt_nFinishDate.Text = Calendar4.SelectedDate.Date
        Div_grid4.Visible = False

    End Sub

    Protected Sub btnClose3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose3.Click

        Div_grid3.Visible = False

    End Sub

    Protected Sub btnClose4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose4.Click

        Div_grid4.Visible = False

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '新表單才可刪除
            If read_only <> "" Then

                '隱藏刪除按鈕
                e.Row.Cells(5).Visible = False

            End If

        End If

    End Sub

    Protected Sub DrDown_nFIXITEM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_nFIXITEM.SelectedIndexChanged
        stmt = "select EFORMSN,PANAME,nAPPTIME,nPHONE,nPLACE,nFIXITEM from p_04 "
        stmt += "where EFORMSN ='" + DrDown_nFIXITEM.SelectedValue + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            txtnFIXITEM.Text = do_sql.G_table.Rows(0).Item("nFIXITEM").ToString
        Else
            txtnFIXITEM.Text = ""
        End If

    End Sub

    Protected Sub GridView10_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView10.SelectedIndexChanged
        txtcomment.Text = GridView10.Rows(GridView10.SelectedRow.RowIndex).Cells(0).Text
        Div_grid10.Visible = False
    End Sub

    Protected Sub But_PHRASE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_PHRASE.Click
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        Dim Visitors As Integer
        Dim str_SQL As String = ""

        '否有申請過請假表單
        db.Open()
        Dim PerCountCom As New SqlCommand("SELECT count(*) as Visitors from Phrase WHERE employee_id = '" & Session("user_id") & "'", db)
        Dim PerRdv = PerCountCom.ExecuteReader()
        If PerRdv.read() Then
            Visitors = PerRdv("Visitors")
        End If
        db.Close()

        If Visitors > 0 Then   '是否需要依類別分類;如請假,派車等
            Div_grid10.Visible = True
            Div_grid10.Style("Top") = "650px"
            Div_grid10.Style("left") = "350px"
        Else

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('無個人批核片語資料,請至批核片語管理新增資料');")
            Response.Write(" </script>")

        End If
    End Sub

    Protected Sub Btn_PHclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_PHclose.Click
        Div_grid10.Visible = False
    End Sub
End Class
