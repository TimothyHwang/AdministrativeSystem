Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_02_MOA02001
    Inherits System.Web.UI.Page

    Public read_only As String = ""
    Dim eformid, user_id, org_uid, streformsn, connstr As String
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

        'user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        eformid = Request.QueryString("eformid")
        streformsn = Request.QueryString("eformsn")
        read_only = Request.QueryString("read_only")

        'session被清空回首頁
        If user_id = "" And org_uid = "" And streformsn = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '判斷選擇哪個填表人
            Try

                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '找出表單審核者
                If read_only = "2" Then

                    db.Open()
                    Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & streformsn & "' and hddate is null", db)
                    Dim RdrAgentCheck = strAgentCheck.ExecuteReader()
                    If RdrAgentCheck.read() Then
                        AgentEmpuid = RdrAgentCheck.item("empuid")
                    End If
                    db.Close()

                End If

                '找出同級單位以下全部單位
                Dim Org_Down As New C_Public

                SqlDataSource1.SelectCommand = "SELECT * FROM V_EmpInfo WHERE (ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ")) ORDER BY emp_chinese_name"

                '找出登入者的一級單位
                Dim strParentOrg As String = ""
                strParentOrg = Org_Down.getUporg(org_uid, 1)

                '會議室
                SqlDataSource3.SelectCommand = "SELECT [MeetSn], [MeetName] FROM [P_0201] WHERE (share = 1 OR Org_Uid IN (" & Org_Down.getchildorg(strParentOrg) & ")) AND Enabled=1 ORDER BY Share DESC,[MeetName]"

                If read_only = "2" Then
                    send.Text = "核准"
                    '判斷是否由入口網站送件
                    If Session("user_id") = "" Then
                        user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                    End If


                End If

                '填表人資料
                db.Open()
                Dim strPer As New SqlCommand("SELECT ORG_NAME,emp_chinese_name,AD_Title FROM V_EmpInfo WHERE employee_id = '" & user_id & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    Lab_ORG_NAME_1.Text = RdPer("ORG_NAME")
                    Lab_emp_chinese_name.Text = RdPer("emp_chinese_name")
                    Lab_title_name_1.Text = RdPer("AD_Title")
                End If
                db.Close()

                '申請時間
                AppDate.Text = Now

                '顯示新增的會議室資料
                GridView1.DataBind()

            Catch ex As Exception

            End Try


        End If


    End Sub

    Protected Sub DrDown_emp_chinese_name_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_emp_chinese_name.SelectedIndexChanged

        '判斷選擇哪個申請人
        AppUserFun()

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        If IsPostBack = False Then

            '將下拉Default選擇為登入者
            Dim i As Integer
            For i = 0 To DrDown_emp_chinese_name.Items.Count - 1
                If UCase(DrDown_emp_chinese_name.Items(i).Value) = UCase(user_id) Then
                    DrDown_emp_chinese_name.Items(i).Selected = True

                    '判斷選擇哪個申請人
                    AppUserFun()
                End If
            Next

            '單純讀取表單不可送件
            If read_only = "1" Then

                InsMeeting.Visible = False
                send.Visible = False
                backBtn.Visible = False
                tranBtn.Visible = False
                supplyBtn.Visible = False
                'ImgDate1.Visible = False
                StartDay.Enabled = False
                htime_start.Enabled = False
                ShowDetail(streformsn)

                '讀取表單可送件
            ElseIf read_only = "2" Then

                InsMeeting.Visible = False
                supplyBtn.Visible = False
                'ImgDate1.Visible = False
                StartDay.Enabled = False
                htime_start.Enabled = False
                ShowDetail(streformsn)

                Dim PerCount As Integer = 0

                '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
                Dim ParentFlag As String = ""
                Dim ParentVal = user_id & "," & streformsn

                Dim FC As New C_FlowSend.C_FlowSend
                ParentFlag = FC.F_NextChief(ParentVal, connstr)

                If ParentFlag = "1" Then
                    'send.Visible = False

                    '截取下一層申請人之帳號
                    db.Open()
                    Dim PerNextCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                    Dim PerNextRdv = PerNextCom.ExecuteReader()
                    If PerNextRdv.Read() Then
                        PerCount = PerNextRdv("PerCount")
                    End If
                    db.Close()
                    Session("apply_user_id") = ""

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
                        send.Visible = True
                        tranBtn.Visible = False
                    Else
                        send.Visible = True
                        'send.Visible = False
                    End If

                End If

                '找出同級單位以下全部單位
                Dim Org_Down As New C_Public

                '沒有批核權限不可執行動作
                If (Org_Down.ApproveAuth(streformsn, user_id)) = "" Then
                    send.Visible = False
                    backBtn.Visible = False
                    tranBtn.Visible = False
                    supplyBtn.Visible = False
                End If

            Else
                backBtn.Visible = False
                tranBtn.Visible = False
            End If

        End If

    End Sub

    Public Function AppUserFun()

        '判斷選擇哪個申請人
        Try

            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '申請人資料
            db.Open()
            Dim strPer As New SqlCommand("SELECT ORG_NAME,AD_Title FROM V_EmpInfo WHERE employee_id = '" & DrDown_emp_chinese_name.SelectedValue & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                Lab_ORG_NAME_2.Text = RdPer("ORG_NAME")
                Lab_title_name_2.Text = RdPer("AD_Title")
            End If
            db.Close()

        Catch ex As Exception

        End Try

        AppUserFun = ""

    End Function

    Protected Sub send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles send.Click

        Try

            '新增資料
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '表單審核
            Dim FC As New C_FlowSend.C_FlowSend

            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                'SendVal = eformid & "," & user_id & "," & streformsn & "," & "1" & ","
                ''簽核時須依登入者ID為user_ID
                If read_only = "" Then
                    SendVal = eformid & "," & DrDown_emp_chinese_name.SelectedValue & "," & streformsn & "," & "1" & ","
                ElseIf read_only = "2" Then
                    SendVal = eformid & "," & user_id & "," & streformsn & "," & "1" & ","
                End If

            Else
                SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1" & ","
            End If

                Dim NextPer As Integer = 0

                '關卡為上一級主管有多少人
                NextPer = FC.F_NextStep(SendVal, connstr)

                Dim Chknextstep As Integer

                '判斷表單關卡
                db.Open()
                Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & streformsn & "' and empuid = '" & user_id & "' and hddate is null", db)
                Dim RdrCheck = strComCheck.ExecuteReader()
                If RdrCheck.Read() Then
                    Chknextstep = RdrCheck.Item("nextstep")
                End If
                db.Close()

                Dim strgroup_id As String = ""

                If read_only <> "" Then

                    '判斷上一級主管
                    db.Open()
                    Dim strUpComCheck As New SqlCommand("SELECT group_id FROM flow WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ", db)
                    Dim RdrUpCheck = strUpComCheck.ExecuteReader()
                    If RdrUpCheck.Read() Then
                        strgroup_id = RdrUpCheck.Item("group_id")
                    End If
                    db.Close()

                End If

                If NextPer = 0 And read_only = "" Then

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('無上一級主管');")
                    Response.Write(" </script>")

                    Exit Sub

                End If

                '判斷欄位是否輸入
                Dim ErrFlag As String = ""
                Dim ErrFlagPer As String = ""


                If PHONE.Text = "" Then
                    ErrFlag = "1"
                End If

                If MeetName.Text = "" Then
                    ErrFlag = "1"
                End If

                If speaker.Text = "" Then
                    ErrFlag = "1"
                End If

                If Status.SelectedValue = " " Then
                    ErrFlag = "1"
                End If

                If (JoinPerson.Text.Length.ToString) > 255 Then
                    ErrFlag = "1"
                    ErrFlagPer = "1"
                End If

                If (MeetContect.Text.Length.ToString) > 255 Then
                    ErrFlag = "1"
                    ErrFlagPer = "1"
                End If

                Dim MeetInsFlag As String = ""

                '判斷是否有新增會議室日期跟時段
                db.Open()
                Dim strMeet As New SqlCommand("SELECT MTT_Num FROM P_0204 WHERE EFORMSN = '" & streformsn & "'", db)
                Dim RdMeet = strMeet.ExecuteReader()
                If RdMeet.Read() Then
                    MeetInsFlag = "1"
                End If
                db.Close()

                If read_only = "2" Then

                    Dim strAgentName As String = ""

                    '判斷是否為代理人批核
                    If UCase(user_id) = UCase(AgentEmpuid) Then

                        '增加批核意見
                        insComment(txtcomment.Text, streformsn, user_id)

                    Else

                        Dim strComment As String = ""

                        '找尋批核者姓名
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.Read() Then
                            strAgentName = RdPer("emp_chinese_name")
                        End If
                        db.Close()

                        strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)"

                        '增加批核意見
                        insComment(strComment, streformsn, AgentEmpuid)

                    End If
                    But_PHRASE.Enabled = False
                    txtcomment.Enabled = False
                End If

                If MeetInsFlag = "1" And ErrFlag = "" Then

                    '表單送件
                    Dim Val_P As String
                    Val_P = ""

                    '判斷新表單
                    If read_only = "" Then

                        '新增申請會議室資料
                        InsData()

                    End If

                    '判斷下一關為上一級主管時人數是否超過一人
                    If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
                        Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & streformsn & "&SelFlag=1" & "&AppUID=" & DrDown_emp_chinese_name.SelectedValue)
                    Else

                        Val_P = FC.F_Send(SendVal, connstr)

                        Dim PageUp As String = ""

                        If read_only = "" Then
                            PageUp = "New"
                        End If

                        Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                        'do_sql.G_errmsg = "新增成功"

                    End If

                Else
                    If ErrFlag <> "" Then
                        If ErrFlagPer = "1" Then
                            labMeetNull.Text = "參加人員或會議內容超過255中文字"
                        Else
                            labMeetNull.Text = "請輸入全部必填欄位"
                        End If
                    Else
                        labMeetNull.Text = "請新增會議室開會日期"
                    End If

                End If


        Catch ex As Exception

        End Try

    End Sub

    Protected Sub InsMeeting_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles InsMeeting.Click

        Try
            '新增資料
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            Dim MeetErr As String = ""

            '判斷會議室日期跟時段
            db.Open()
            Dim StepSQL As String = "SELECT distinct MTT_Num,MeetHour FROM P_0204,flowctl WHERE P_0204.eformsn=flowctl.eformsn AND MeetSn = '" & Place.SelectedValue & "' AND MeetTime between '" & StartDay.Text & " 00:00" & "' AND '" & StartDay.Text & " 23:59" & "'"
            Dim ds As New DataSet
            Dim Dt As New DataTable
            Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
            da.Fill(ds)
            '設定DataTable
            Dt = ds.Tables(0)
            For x As Integer = 0 To Dt.Rows.Count - 1
                If Dt.Rows(x).Item(1) = "全天" Then
                    MeetErr = "會議室申請日期重複"
                Else
                    If htime_start.SelectedValue = Dt.Rows(x).Item(1) Then
                        MeetErr = "會議室申請日期重複"
                    ElseIf htime_start.SelectedValue = "全天" Then
                        MeetErr = "會議室申請日期重複"
                    End If
                End If
            Next
            db.Close()

            '判斷同表單是否申請相同會議室
            db.Open()
            Dim strPer As New SqlCommand("SELECT MTT_Num,MeetHour FROM P_0204 WHERE eformsn = '" & streformsn & "' AND MeetSn = '" & Place.SelectedValue & "'  AND MeetTime between '" & StartDay.Text & " 00:00" & "' AND '" & StartDay.Text & " 23:59" & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then

                If RdPer("MeetHour") = "全天" Then
                    MeetErr = "會議室申請日期重複"
                Else
                    If htime_start.SelectedValue = RdPer("MeetHour") Then
                        MeetErr = "會議室申請日期重複"
                    ElseIf htime_start.SelectedValue = "全天" Then
                        MeetErr = "會議室申請日期重複"
                    End If
                End If

            End If
            db.Close()

            If StartDay.Text <> "" And MeetErr = "" Then

                '新增會議日期資料
                db.Open()
                Dim insCom As New SqlCommand("INSERT INTO P_0204(EFORMSN,MeetSn,MeetTime,MeetHour) VALUES ('" & streformsn & "','" & Place.SelectedValue & "','" & StartDay.Text & "','" & htime_start.SelectedValue & "')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                labMeetErr.Text = ""

            Else
                labMeetErr.Text = MeetErr

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('" & MeetErr & "');")
                Response.Write(" </script>")

            End If

            '顯示新增的會議室資料
            GridView1.DataBind()

        Catch ex As Exception

        End Try


    End Sub

    Public Function ShowDetail(ByVal eformsn As String)

        '判斷選擇哪個申請人
        Try

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '會議室申請單資料
            db.Open()
            Dim strMeet As New SqlCommand("SELECT * FROM P_02 WHERE EFORMSN = '" & eformsn & "'", db)
            Dim RdMeet = strMeet.ExecuteReader()

            If RdMeet.read() Then

                Lab_ORG_NAME_1.Text = RdMeet("PWUNIT")
                Lab_ORG_NAME_2.Text = RdMeet("PAUNIT")
                Lab_emp_chinese_name.Text = RdMeet("PWNAME")

                '先清空使用者
                DrDown_emp_chinese_name.Items.Clear()
                DrDown_emp_chinese_name.Items.Add(New ListItem(RdMeet("PANAME"), RdMeet("PAIDNO")))

                Lab_title_name_1.Text = RdMeet("PWTITLE")
                Lab_title_name_2.Text = RdMeet("PATITLE")


                AppDate.Text = RdMeet("nAPPLYTIME")
                PHONE.Text = RdMeet("nPHONE")
                MeetName.Text = RdMeet("nMEETNAME")

                '判斷選擇會議類別
                Dim i As Integer
                For i = 0 To Status.Items.Count - 1
                    If Status.Items(i).Value = RdMeet("nSTATUS") Then
                        Status.Items(i).Selected = True
                    End If
                Next

                speaker.Text = RdMeet("nSPEAKER")

                '判斷選擇會議室
                MeetSel(RdMeet("nPLACE"))

                JoinPerson.Text = RdMeet("nJOINPERSON")
                MeetContect.Text = RdMeet("nCONTECT")
            End If
            db.Close()


            '表單欄位不可修改
            DrDown_emp_chinese_name.Enabled = False
            PHONE.Enabled = False
            MeetName.Enabled = False
            speaker.Enabled = False
            Status.Enabled = False
            Place.Enabled = False
            JoinPerson.Enabled = False
            MeetContect.Enabled = False

        Catch ex As Exception

        End Try

        ShowDetail = ""

    End Function

    Public Sub MeetSel(ByVal strplace As String)

        '判斷選擇會議室
        Try

            Dim j As Integer
            For j = 0 To Place.Items.Count - 1
                If Place.Items(j).Text = strplace Then
                    Place.Items(j).Selected = True
                Else
                    Place.Items(j).Selected = False
                End If
            Next

        Catch ex As Exception

        End Try

        'MeetSel = ""

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '新表單才可刪除
            If read_only <> "" Then

                '隱藏刪除按鈕
                e.Row.Cells(4).Visible = False

            End If

        End If

    End Sub

    Protected Sub tranBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tranBtn.Click
        Try
            If GetSuperiors(user_id).Rows.Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                Server.Transfer("../00/MOA00014.aspx?eformsn=" + streformsn)
            Else
                If read_only = "2" Then
                    '增加批核意見
                    insComment(txtcomment.Text, streformsn, user_id)
                End If

                Dim Val_P As String
                Val_P = ""

                '表單呈轉
                Dim FC As New C_FlowSend.C_FlowSend

                Dim SendVal As String = ""

                '判斷是否為代理人批核的表單
                If AgentEmpuid = "" Then
                    SendVal = eformid & "," & user_id & "," & streformsn & "," & "1"
                Else
                    SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1"
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

    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles backBtn.Click

        Try

            If read_only = "2" Then
                '增加批核意見
                insComment(txtcomment.Text, streformsn, user_id)
            End If

            Dim Val_P As String
            Val_P = ""

            '表單駁回
            Dim FC As New C_FlowSend.C_FlowSend


            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & streformsn & "," & "1"
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1"
            End If

            Val_P = FC.F_Back(SendVal, connstr)

            '將申請的會議室撤回

            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '將會議室申請時段放到移除資料表
            db.Open()
            Dim insCom As New SqlCommand("INSERT INTO P_0205 (MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, DelUser) SELECT MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, '" & user_id & "' FROM P_0204 WHERE EFORMSN='" & streformsn & "'", db)
            insCom.ExecuteNonQuery()
            db.Close()

            '刪除申請的會議室時段
            db.Open()
            Dim delCom As New SqlCommand("DELETE FROM P_0204 WHERE (EFORMSN='" & streformsn & "')", db)
            delCom.ExecuteNonQuery()
            db.Close()

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單已駁回給申請人');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub supplyBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles supplyBtn.Click

        Try

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷表單是否回上頁重新送件
            db.Open()
            Dim strReSend As New SqlCommand("SELECT * FROM P_02 WHERE EFORMSN = '" & streformsn & "'", db)
            Dim RdReSend = strReSend.ExecuteReader()

            If RdReSend.read() Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('表單編號重複，請重新申請');")
                Response.Write(" window.parent.location='../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N';")
                Response.Write(" </script>")
            Else

                '判斷欄位是否輸入
                Dim ErrFlag As String = ""

                If PHONE.Text = "" Then
                    ErrFlag = "1"
                End If

                If MeetName.Text = "" Then
                    ErrFlag = "1"
                End If

                If speaker.Text = "" Then
                    ErrFlag = "1"
                End If

                If Status.SelectedValue = " " Then
                    ErrFlag = "1"
                End If

                If (MeetContect.Text.Length.ToString) > 255 Then
                    ErrFlag = "1"
                End If

                Dim MeetInsFlag As String = ""

                '判斷是否有新增會議室日期跟時段
                db.Open()
                Dim strMeet As New SqlCommand("SELECT MTT_Num FROM P_0204 WHERE EFORMSN = '" & streformsn & "'", db)
                Dim RdMeet = strMeet.ExecuteReader()
                If RdMeet.read() Then
                    MeetInsFlag = "1"
                End If
                db.Close()

                '判斷新表單
                If read_only = "" Then

                    If MeetInsFlag = "1" Then
                        '新增申請會議室資料
                        InsData()
                    End If

                End If

                If MeetInsFlag = "1" And ErrFlag = "" Then

                    Dim Val_P As String
                    Val_P = ""

                    '表單補登
                    Dim FC As New C_FlowSend.C_FlowSend

                    Dim SendVal As String = ""

                    '判斷是否為代理人批核的表單
                    If AgentEmpuid = "" Then
                        SendVal = eformid & "," & user_id & "," & streformsn & "," & "1"
                    Else
                        SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1"
                    End If

                    Val_P = FC.F_Supply(SendVal, connstr)

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('補登完成');")
                    Response.Write(" window.parent.location='../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N';")
                    Response.Write(" </script>")

                Else
                    If ErrFlag <> "" Then
                        labMeetNull.Text = "請輸入全部必填欄位"
                    Else
                        labMeetNull.Text = "請新增會議室開會日期"
                    End If

                End If

            End If


        Catch ex As Exception

        End Try


    End Sub

    Public Sub InsData()

        Try
            '新增資料
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '新增會議日期資料
            db.Open()

            Dim InsP02 As String = ""
            '表單序號,填表人單位,填表人級職,填表人姓名,填表人身份證字號,申請人單位,申請人姓名,申請人級職,申請人身份證字號
            '聯絡電話,會議名稱,會議類別,主持人,會議地點,參加人員,會議內容
            InsP02 = "INSERT INTO P_02(EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME, PATITLE, PAIDNO "
            InsP02 += ",nPHONE,nMEETNAME,nSTATUS,nSPEAKER,nPLACE,nJOINPERSON,nCONTECT) "
            InsP02 += " VALUES ('" & streformsn & "','" & Lab_ORG_NAME_1.Text & "','" & Lab_title_name_1.Text & "',N'" & Lab_emp_chinese_name.Text & "','" & user_id & "','" & Lab_ORG_NAME_2.Text & "',N'" & DrDown_emp_chinese_name.SelectedItem.Text & "','" & Lab_title_name_2.Text & "','" & DrDown_emp_chinese_name.SelectedValue & "'"
            InsP02 += ",'" & PHONE.Text & "','" & MeetName.Text & "','" & Status.SelectedValue & "','" & speaker.Text & "','" & Place.SelectedItem.Text & "','" & JoinPerson.Text & "','" & MeetContect.Text & "')"

            Dim insCom As New SqlCommand(InsP02, db)
            insCom.ExecuteNonQuery()
            db.Close()

        Catch ex As Exception

        End Try

    End Sub

    Public Sub insComment(ByVal strComment As String, ByVal strEformsn As String, ByVal strid As String)

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

        'insComment = ""

    End Sub

    'Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

    '    Div_grid.Visible = True
    '    Div_grid.Style("Top") = "230px"
    '    Div_grid.Style("left") = "240px"

    'End Sub
    'Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

    '    StartDay.Text = Calendar1.SelectedDate.Date
    '    Div_grid.Visible = False

    'End Sub

    'Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

    '    Div_grid.Visible = False

    'End Sub

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        txtcomment.Text = GridView2.Rows(GridView2.SelectedRow.RowIndex).Cells(0).Text
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
            Div_grid10.Style("Top") = "570px"
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
