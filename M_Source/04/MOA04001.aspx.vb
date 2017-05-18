Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_04_MOA04001
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
    Public stepChk, connstr As String
    Public pfilename As String = ""
    Public ErrMsg As String = ""
    Dim user_id, org_uid As String
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

            Dim strobject_uide As String = ""

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '找出表單審核者
            If read_only = "2" Then

                db.Open()
                Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' and hddate is null", db)
                Dim RdrAgentCheck = strAgentCheck.ExecuteReader()
                If RdrAgentCheck.read() Then
                    AgentEmpuid = RdrAgentCheck.item("empuid")
                End If
                db.Close()

            End If

            ''填表人資料
            'db.Open()
            'Dim strPer As New SqlCommand("SELECT O.object_uid FROM SYSTEMOBJ O,SYSTEMOBJUSE U WHERE O.object_uid=U.object_uid AND U.employee_id = '" & do_sql.G_user_id & "'", db)
            'Dim RdPer = strPer.ExecuteReader()
            'If RdPer.read() Then
            '    strobject_uide = RdPer("object_uid")
            'End If
            'db.Close()

            'If strobject_uide = "K489" Then
            '    '房舍水電修繕管制單位
            '    stepChk = "1"
            'ElseIf strobject_uide = "Q140" Then
            '    '派工單位
            '    stepChk = "2"
            'End If

            '填表人資料
            Dim strsql As String = "SELECT O.object_uid FROM SYSTEMOBJ O,SYSTEMOBJUSE U WHERE O.object_uid=U.object_uid AND U.employee_id = '" & do_sql.G_user_id & "'"
            db.Open()
            Dim dt1 As DataTable = New DataTable()
            Dim da1 As SqlDataAdapter = New SqlDataAdapter(strsql, db)
            da1.Fill(dt1)
            db.Close()


            If dt1.Rows.Count <> 0 Then

                For j As Integer = 0 To dt1.Rows.Count - 1
                    If dt1.Rows(j).Item(0) = "K489" Then
                        '房舍水電修繕管制單位
                        stepChk = "1"
                    ElseIf dt1.Rows(j).Item(0) = "Q140" Then
                        '派工單位
                        stepChk = "2"
                    End If
                Next

            End If

            'do_sql.G_user_id = "tempu178"
            'eformsn = "R5Y26NOFA5123456" 'd_pub.randstr(16) '
            'eformid = "R5Y26NOFA5"
            'read_only = "1"
            If IsPostBack Then
                Exit Sub
            End If
            'Txt_nPacthCount
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
                stmt = "select * from Employee where  ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") order by emp_chinese_name"
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

        If read_only = "2" Then
            If stepChk = "1" Then

                If Txt_nFacilityNo.Text = "" Then
                    'do_sql.G_errmsg = "設施編號必需輸入"
                    'Exit Sub
                End If

                stmt = "update P_04 set nExternal='" + rdo_nExternal.SelectedValue + "'" '承辦類別
                stmt += ",nFacilityNo='" + Txt_nFacilityNo.Text + "'" '設施編號
                stmt += " where EFORMSN='" + eformsn + "'"

                If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If

            ElseIf stepChk = "2" Then

                If Txt_nFIXDATE.Text = "" Then
                    do_sql.G_errmsg = "維修時間必需輸入"
                    Exit Sub
                End If
                If Txt_nPacthCount.Text = "" Then
                    do_sql.G_errmsg = "派員人數必需輸入"
                    Exit Sub
                End If
                'If Txt_nPacthPer.Text = "" Then
                '    do_sql.G_errmsg = "施工人員必需輸入"
                '    Exit Sub
                'End If

                stmt = "update P_04 set nFIXDATE='" + Txt_nFIXDATE.Text + "'" '維修時間
                stmt += ",nPacthCount=" + Txt_nPacthCount.Text  '派員人數
                stmt += ",nPacthPer='" + Txt_nPacthPer.Text + "'" '施工人員
                stmt += " where EFORMSN='" + eformsn + "'"

                If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If

            End If

        Else

            If check_text() = False Then
                Exit Sub
            End If

            stmt = "insert into P_04(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
            stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
            stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
            stmt += "nAPPTIME, nPHONE, " '申請時間,申請事由,
            stmt += "nPLACE,nFIXITEM) " '申請出入日期,使用地點        
            stmt += " values('" + eformsn + "','" + Lab_PWUNIT.Text + "','" + Lab_PWTITLE.Text + "',N'"
            stmt += Lab_PWNAME.Text + "','" + do_sql.G_user_id + "','" + Lab_PAUNIT.Text + "',N'"
            stmt += DrDown_PANAME.SelectedItem.Text + "','" + Lab_PATITLE.Text + "','" + DrDown_PANAME.SelectedItem.Value + "','"
            stmt += Label8.Text + "','" + Txt_nPHONE.Text + "','"
            stmt += Txt_nPLACE.Text + "','" + Txt_nFIXITEM.Text + "')"
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

        '當新表單判斷同單位是否有單位行政官
        '同單位沒有則導入下一頁
        '下一頁找出全部同一級單位的行政官

        If read_only = "" Then

            '判斷同單位是否有單位行政官
            Dim OAPer As Integer = 0
            Dim OAPerAll As Integer = 0
            Dim strOrgTop As String = ""
            Dim OAPerORG As String = ""

            Dim strsame As String = "SELECT count(EMPLOYEE.employee_id) as OAPer FROM SYSTEMOBJUSE INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID WHERE (SYSTEMOBJUSE.object_uid = 'p12') AND (EMPLOYEE.ORG_UID = '" & org_uid & "')"

            db.Open()
            Dim OAPerCom As New SqlCommand(strsame, db)
            Dim OAPerRdr = OAPerCom.ExecuteReader()
            If OAPerRdr.read() Then
                OAPer = OAPerRdr.item("OAPer")
            End If
            db.Close()

            '同單位無行政官找出一級單位全部行政官
            '同單位有行政官則送出
            '同單位有兩位以上則導入下一頁
            If OAPer = 0 Then

                Dim CP As New C_Public
                strOrgTop = CP.getUporg(org_uid, 1)

                db.Open()
                Dim OAPerAllCom As New SqlCommand("SELECT count(EMPLOYEE.employee_id) as OAPerAll FROM SYSTEMOBJUSE INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID WHERE (SYSTEMOBJUSE.object_uid = 'p12') AND (EMPLOYEE.ORG_UID IN (" & CP.getchildorg(strOrgTop) & "))", db)
                Dim OAPerAllRdr = OAPerAllCom.ExecuteReader()
                If OAPerAllRdr.read() Then
                    OAPerAll = OAPerAllRdr.item("OAPerAll")
                End If
                db.Close()

                If OAPerAll = 0 Then

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('無任何單位行政官');")
                    Response.Write(" </script>")

                    Exit Sub

                ElseIf OAPerAll = 1 Then

                    '找出一級單位唯一的單位行政官
                    db.Open()
                    Dim OAPerOneCom As New SqlCommand("SELECT DISTINCT EMPLOYEE.employee_id FROM SYSTEMOBJUSE INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID WHERE (SYSTEMOBJUSE.object_uid = 'p12') AND (EMPLOYEE.ORG_UID IN (" & CP.getchildorg(strOrgTop) & "))", db)
                    Dim OAPerOneRdr = OAPerOneCom.ExecuteReader()
                    If OAPerOneRdr.read() Then
                        OAPerORG = OAPerOneRdr.item("employee_id")
                    End If
                    db.Close()

                    'Dim SendValOA = eformid & "," & do_sql.G_user_id & "," & eformsn & "," & "1" & "," & OAPerORG
                    Dim SendValOA = eformid & "," & DrDown_PANAME.SelectedValue & "," & eformsn & "," & "1" & "," & OAPerORG

                    '表單審核
                    Val_P = FC.F_Send(SendValOA, do_sql.G_conn_string)
                    Dim PageUp As String = ""

                    If read_only = "" Then
                        PageUp = "New"
                    End If

                    Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                    do_sql.G_errmsg = "存檔成功"

                ElseIf OAPerAll > 0 Then
                    Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=3&strOrgTop=" & strOrgTop & "&AppUID=" & DrDown_PANAME.SelectedValue)
                End If

            ElseIf OAPer = 1 Then
                '表單審核
                Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)
                Dim PageUp As String = ""

                If read_only = "" Then
                    PageUp = "New"
                End If

                Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                do_sql.G_errmsg = "存檔成功"
            ElseIf OAPer > 1 Then
                Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=2" & "&AppUID=" & DrDown_PANAME.SelectedValue)
            End If



        Else

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
                do_sql.G_errmsg = "存檔成功"
            End If

        End If


    End Sub
    Function check_text() As Boolean
        check_text = False
        If Lab_PAUNIT.Text = "" Then
            do_sql.G_errmsg = "申請人姓名必須選取"
            Exit Function
        End If
        If Txt_nPHONE.Text = "" Then
            do_sql.G_errmsg = "電話必須輸入"
            Exit Function
        End If
        If Txt_nPLACE.Text = "" Then
            do_sql.G_errmsg = "地點:必須輸入"
            Exit Function
        End If
        If Txt_nFIXITEM.Text = "" Then
            do_sql.G_errmsg = "修理項目必須選取"
            Exit Function
        End If
        check_text = True
    End Function
    Function select_data() As Boolean
        select_data = False
        stmt = "select * from p_04 "
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
            DrDown_PANAME.Items.Add(New ListItem("", ""))
            'DrDown_PANAME.Items(0).Value = ""

            Lab_PATITLE.Text = ""
            Label8.Text = ""

            Txt_nPHONE.Text = ""
            Txt_nPLACE.Text = ""
            Txt_nFIXITEM.Text = ""
            If read_only = "2" Then
                Txt_nFIXDATE.Text = ""
                Txt_nPacthCount.Text = ""
                Txt_nPacthPer.Text = ""
            End If
            Call disabel_field()
            Exit Function
        End If
        Lab_PWUNIT.Text = do_sql.G_table.Rows(0).Item("PWUNIT").ToString
        Lab_PWNAME.Text = do_sql.G_table.Rows(0).Item("PWNAME").ToString
        Lab_PWTITLE.Text = do_sql.G_table.Rows(0).Item("PWTITLE").ToString
        Lab_PAUNIT.Text = do_sql.G_table.Rows(0).Item("PAUNIT").ToString
        DrDown_PANAME.Items.Clear()
        DrDown_PANAME.Items.Add(New ListItem(Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString), Trim(do_sql.G_table.Rows(0).Item("PAIDNO").ToString)))
        'DrDown_PANAME.Items(0).Value = Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString)

        Lab_PATITLE.Text = do_sql.G_table.Rows(0).Item("PATITLE").ToString
        Label8.Text = do_sql.G_table.Rows(0).Item("nAPPTIME").ToString

        Txt_nPHONE.Text = do_sql.G_table.Rows(0).Item("nPHONE").ToString
        Txt_nPLACE.Text = do_sql.G_table.Rows(0).Item("nPLACE").ToString
        Txt_nFIXITEM.Text = do_sql.G_table.Rows(0).Item("nFIXITEM").ToString

        If read_only = "1" Or read_only = "2" Then
            If do_sql.G_table.Rows(0).Item("nFIXDATE").ToString = "" Then
                Txt_nFIXDATE.Text = ""
            Else
                Txt_nFIXDATE.Text = CDate(do_sql.G_table.Rows(0).Item("nFIXDATE").ToString).ToString("yyyy/MM/dd")
            End If

            Txt_nPacthCount.Text = do_sql.G_table.Rows(0).Item("nPacthCount").ToString
            Txt_nPacthPer.Text = do_sql.G_table.Rows(0).Item("nPacthPer").ToString

            If do_sql.G_table.Rows(0).Item("nExternal").ToString = "外包" Then
                rdo_nExternal.Items(1).Selected = True
            Else
                rdo_nExternal.Items(0).Selected = True
            End If

            Txt_nFacilityNo.Text = do_sql.G_table.Rows(0).Item("nFacilityNo").ToString
        End If

        Call disabel_field()
        select_data = True
    End Function
    Sub disabel_field()

        DrDown_PANAME.Enabled = False
        Txt_nPHONE.Enabled = False
        Txt_nPLACE.Enabled = False
        Txt_nFIXITEM.Enabled = False

        If read_only = "1" Then
            But_exe.Visible = False

            rdo_nExternal.Enabled = False
            Txt_nFacilityNo.Enabled = False

            Txt_nFIXDATE.Enabled = False
            Txt_nPacthCount.Enabled = False
            Txt_nPacthPer.Enabled = False
            ImgDate1.Visible = False

        ElseIf read_only = "2" Then

            If stepChk = "1" Then

                Txt_nFIXDATE.Enabled = False
                Txt_nPacthCount.Enabled = False
                Txt_nPacthPer.Enabled = False
                ImgDate1.Visible = False

            ElseIf stepChk = "2" Then

                rdo_nExternal.Enabled = False
                Txt_nFacilityNo.Enabled = False
            Else

                rdo_nExternal.Enabled = False
                Txt_nFacilityNo.Enabled = False

                Txt_nFIXDATE.Enabled = False
                Txt_nPacthCount.Enabled = False
                Txt_nPacthPer.Enabled = False
                ImgDate1.Visible = False
            End If


        End If

    End Sub

    Protected Sub But_prt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_prt.Click
        'eformsn = GridView1.SelectedDataKey.Item(0)
        Dim filename As String = "prt_040010" & Rnd() & ".drs"
        Dim print As New C_Xprint
        print.C_Xprint("rpt040010.txt", filename)
        print.NewPage()
        Prt_04(print) '增加列印內容
        print.EndFile()
        If (print.ErrMsg <> "") Then '產生列印檔過程有錯誤
            ErrMsg = print.ErrMsg
        Else '顯示報表檔
            pfilename = "../../drs/" & filename
        End If
    End Sub
    Protected Sub Prt_04(ByVal prt As C_Xprint) '列印
        Dim dv As DataView
        Dim Sql As String

        Sql = "SELECT *"
        Sql += " FROM p_04"
        Sql += " where P_04.EFORMSN='" + eformsn + "'"
        SqlDataSource3.SelectCommand = Sql
        dv = CType(SqlDataSource3.Select(DataSourceSelectArguments.Empty), DataView)

        If (dv.Count() > 0) Then
            prt.Add("填表人單位", dv.Table.Rows(0)("PWUNIT").ToString, 0, 0)
            prt.Add("填表人級職", dv.Table.Rows(0)("PWTITLE").ToString, 0, 0)
            prt.Add("填表人姓名", dv.Table.Rows(0)("PWNAME").ToString, 0, 0)
            prt.Add("申請時間", CDate(dv.Table.Rows(0)("nAPPTIME")).ToString("yyyy/MM/dd hh:mm"), 0, 0)
            prt.Add("電話", dv.Table.Rows(0)("nPHONE").ToString, 0, 0)
            prt.Add("修理項目", dv.Table.Rows(0)("nFIXITEM").ToString, 0, 0)
            If Not dv.Table.Rows(0)("nFIXDATE") Is DBNull.Value Then
                prt.Add("維修時間", CDate(dv.Table.Rows(0)("nFIXDATE")).ToString("yyyy/MM/dd hh:mm"), 0, 0)
            Else
                prt.Add("維修時間", "", 0, 0)
            End If
            prt.Add("派員人數", dv.Table.Rows(0)("nPacthCount").ToString, 0, 0)
            prt.Add("施工人員", dv.Table.Rows(0)("nPacthPer").ToString, 0, 0)
            prt.Add("承辦類別", dv.Table.Rows(0)("nExternal").ToString, 0, 0)
            prt.Add("申請人單位", dv.Table.Rows(0)("PAUNIT").ToString, 0, 0)
            prt.Add("申請人姓名", dv.Table.Rows(0)("PANAME").ToString, 0, 0)
            prt.Add("申請人級職", dv.Table.Rows(0)("PATITLE").ToString, 0, 0)
            prt.Add("地點", dv.Table.Rows(0)("nPLACE").ToString, 0, 0)
            prt.Add("批核資料", "流水號:" & dv.Table.Rows(0)("P_Num").ToString, 0, 0)
        End If



        '開啟連線
        Dim db As New SqlConnection(connstr)

        '查房舍批核資料
        Dim strsql = "Select * From flowctl where eformsn='" & eformsn & "'"
        db.Open()
        Dim dt2 As DataTable = New DataTable("")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(strsql, db)
        da2.Fill(dt2)
        db.Close()

        Dim intAddr As Integer = 0
        Dim strFlowDetail As String = ""

        For x As Integer = 0 To dt2.Rows.Count - 1

            intAddr += 15

            strFlowDetail = dt2.Rows(x).Item("emp_chinese_name").ToString & " " & dt2.Rows(x).Item("group_name").ToString & " " & dt2.Rows(x).Item("hddate").ToString & " " & dt2.Rows(x).Item("comment").ToString

            prt.Add("批核資料", strFlowDetail, 0, intAddr)

        Next

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

            '判斷批核者是否為上一關送件者的副主管,假如批核者為副主管則不可送件只可呈轉
            Dim ParentFlag As String = ""
            Dim ParentVal = user_id & "," & eformsn

            Dim FC As New C_FlowSend.C_FlowSend
            ParentFlag = FC.F_NextChief(ParentVal, connstr)

            If ParentFlag = "1" Then
                'But_exe.Visible = False

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
            Else
                '判斷下一關是否為上一級主管
                '是的話判斷上一級主管多少人

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

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "250px"
        Div_grid.Style("left") = "280px"

    End Sub
    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Txt_nFIXDATE.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

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
            Div_grid10.Style("Top") = "525px"
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
