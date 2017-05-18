Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_04_MOA04100
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Public read_only As String = ""
    Public confirmFlag As String = ""
    Public stepChk As Integer
    Public strstepsid As String
    Dim eformid, user_id, org_uid, streformsn, connstr As String
    Dim AgentEmpuid As String = ""
    Dim AppUID As String

    Public PageUp As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then
            Dim LoginAll As String = Page.User.Identity.Name.ToString
            Dim LoginID() As String = Split(LoginAll, "\")
            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If
        org_uid = Session("ORG_UID")
        eformid = Request.QueryString("eformid")
        streformsn = Request.QueryString("eformsn")
        read_only = Request.QueryString("read_only")
        lbMsg.Text = ""
          
        'session被清空回首頁
        If user_id = "" And org_uid = "" And streformsn = "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            '判斷填寫批核意見
            confirmFlag = getConfirmFlag(streformsn, user_id)
            '判斷選擇哪個填表人
            Try
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '判斷顯示不同階段表單 
                stepChk = getStepChk(streformsn, user_id)

                '找出表單審核者
                If read_only = "2" Then
                    db.Open()
                    Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & streformsn & "' and hddate is null", db)
                    Dim RdrAgentCheck As IDataReader = strAgentCheck.ExecuteReader()
                    If RdrAgentCheck.Read() Then
                        AgentEmpuid = RdrAgentCheck.Item("empuid")
                    End If
                    db.Close()
                End If
                '找出同級單位以下全部單位
                Dim Org_Down As New C_Public
                SqlDataSource2.SelectCommand = "SELECT * FROM V_EmpInfo WHERE (ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ")) ORDER BY emp_chinese_name"
                '找出登入者的一級單位
                Dim strParentOrg As String = ""
                strParentOrg = Org_Down.getUporg(org_uid, 1)

                DDViewPer.Items.Clear()
                DDViewPer.DataSource = SqlDataSourceEle8
                DDViewPer.DataTextField = "House_Name"
                DDViewPer.DataValueField = "House_Num"
                DDViewPer.DataBind()

                If read_only = "2" Then
                    send.Text = "核准"
                    '判斷是否由入口網站送件
                    If Session("user_id") = "" Then
                        user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                    End If
                End If

                If getNewApp(streformsn) = "true" Then
                    '填表人資料
                    db.Open()
                    Dim strPer As New SqlCommand("SELECT ORG_NAME,emp_chinese_name,AD_Title FROM V_EmpInfo WHERE employee_id = '" & user_id & "'", db)
                    Dim RdPer As IDataReader = strPer.ExecuteReader()
                    If RdPer.Read() Then
                        Lab_PWUNIT.Text = RdPer("ORG_NAME")
                        Lab_PWNAME.Text = RdPer("emp_chinese_name")
                        Lab_PWTITLE.Text = RdPer("AD_Title")
                    End If
                    db.Close()
                    '申請時間
                    If AppDate.Text = "" Then
                        AppDate.Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
                    End If
                Else
                    lbleformsn.Text = streformsn
                End If
            Catch ex As Exception
                lbMsg.Text = ex.Message
            End Try
            Btn_PHclose.Width = GridView10.Width
        End If
    End Sub

    Protected Sub DrDown_PANAME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_PANAME.SelectedIndexChanged
        '判斷選擇哪個申請人
        AppUserFun()
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
            Dim strPer As New SqlCommand("SELECT ORG_NAME,AD_Title FROM V_EmpInfo WHERE employee_id = '" & DrDown_PANAME.SelectedValue & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                Lab_PAUNIT.Text = RdPer("ORG_NAME")
                Lab_PATITLE.Text = RdPer("AD_Title")
            End If
            db.Close()
        Catch ex As Exception
            lbMsg.Text = ex.Message
        End Try
        AppUserFun = ""
    End Function

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string
        '開啟連線
        Dim db As New SqlConnection(connstr)

        If IsPostBack = False Then
            ShowDetail(streformsn)
            If read_only = "1" Or read_only = "" Then
                backBtn.Visible = False
            ElseIf read_only = "2" Then
                setConfirmBtn(streformsn, user_id)
            End If
        End If
        If getNewApp(streformsn) = "true" Then
            AppUserFun()
        End If
        setConfirmBtn(streformsn, user_id)

        If Not Session("PageUp") Is Nothing Or read_only = "1" Then
            '完成任務後，動作按鈕即無法再做第二次點取
            send.Enabled = False
        End If

    End Sub
    Public Function ShowDetail(ByVal eformsn As String)
        Dim strStepSid As String = getstepsid(streformsn, user_id)
        Dim strStatus As String = getWriteFromFlag(streformsn, user_id)
        Dim strFormStatus As String = getFlowStatus(streformsn)
        '判斷選擇哪個申請人
        Try
            '若初次申請時，電話與設備樓層、請修事宜皆需可以key in
            If (read_only <> "" Or strFormStatus = "2" Or strFormStatus = "3") Then
                DrDown_PANAME.Enabled = False
                Txt_nPHONE.Enabled = False
                DDLBD.Enabled = False
                DDLFL.Enabled = False
                DDLRNUM.Enabled = False
                Txt_nFIXITEM.Enabled = False
            ElseIf Not Session("PageUp") Is Nothing Then '當送出完成功後，電話與設備樓層、請修事宜、送出都無法再被操作
                DrDown_PANAME.Enabled = False
                Txt_nPHONE.Enabled = False
                DDLBD.Enabled = False
                DDLFL.Enabled = False
                DDLRNUM.Enabled = False
                Txt_nFIXITEM.Enabled = False
                send.Enabled = False
            End If

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '修繕申請單資料
            db.Open()
            Dim strSQL As String
            strSQL = "SELECT p_0415.*,p_0404.bd_name,p_0406.fl_name,p_0411.rnum_name,p_0416.House_Name,nNowStatus,nFinalDate,nResult FROM P_0415 "
            strSQL += "join p_0404 on p_0404.bd_code=p_0415.nbd_code "
            strSQL += "join p_0406 on p_0406.fl_code=p_0415.nfl_code "
            strSQL += "join p_0411 on p_0411.rnum_code=p_0415.nrnum_code "
            strSQL += "left join p_0416 on p_0415.nViewPer = p_0416.House_Num "
            strSQL += "WHERE EFORMSN = '" & eformsn & "'"

            Dim strMeet As New SqlCommand(strSQL, db)
            Dim RdMeet As IDataReader = strMeet.ExecuteReader()

            If RdMeet.Read() Then

                Lab_PWUNIT.Text = RdMeet("PWUNIT")
                Lab_PAUNIT.Text = RdMeet("PAUNIT")
                Lab_PWNAME.Text = RdMeet("PWNAME")

                '先清空使用者
                DrDown_PANAME.Items.Clear()
                DrDown_PANAME.Items.Add(RdMeet("PANAME"))

                Lab_PWTITLE.Text = RdMeet("PWTITLE")
                Lab_PATITLE.Text = RdMeet("PATITLE")

                AppDate.Text = RdMeet("nAPPTIME")
                Txt_nPHONE.Text = RdMeet("nPHONE")

                DDLBD.Items.Clear()
                DDLBD.Items.Add(New ListItem(RdMeet("bd_name"), RdMeet("nbd_code")))

                DDLFL.Items.Clear()
                DDLFL.Items.Add(New ListItem(RdMeet("fl_name"), RdMeet("nfl_code")))

                DDLRNUM.Items.Clear()
                DDLRNUM.Items.Add(New ListItem(RdMeet("rnum_name"), RdMeet("nrnum_code")))

                Txt_nFIXITEM.Text = RdMeet("nFIXITEM")

                If Not IsDBNull(RdMeet("nFacilityNo")) Then
                    If Not RdMeet("nFacilityNo") = "" Then
                        If read_only <> "1" Then
                            DDViewPer.Items.Clear()
                            DDViewPer.Items.Add(New ListItem(RdMeet("House_Name"), RdMeet("nViewPer")))
                        End If
                        Txt_nViewDate.Text = RdMeet("nViewDate")
                        Txt_nStartDATE.Text = RdMeet("nStartDATE")

                        Txt_nCause.Text = RdMeet("nCause")
                        Txt_nPacthCount.Text = RdMeet("nPacthCount")
                        Txt_nPacthPer.Text = RdMeet("nPacthPer")
                        Txt_nFacilityNo.Text = RdMeet("nFacilityNo")
                        DDL_nErrCause.Items.Clear()
                        Select Case RdMeet("nErrCause")
                            Case "1"
                                DDL_nErrCause.Items.Add(New ListItem("人為因素", RdMeet("nErrCause")))
                            Case "2"
                                DDL_nErrCause.Items.Add(New ListItem("自然因素", RdMeet("nErrCause")))
                            Case "3"
                                DDL_nErrCause.Items.Add(New ListItem("維護查報", RdMeet("nErrCause")))
                        End Select
                        Select Case RdMeet("nExternal")
                            Case "內派"
                                rdo_nExternal.SelectedIndex = 0
                            Case "外包"
                                rdo_nExternal.SelectedIndex = 1
                        End Select
                    End If
                End If
                If Not IsDBNull(RdMeet("nNowStatus")) Then
                    If Not RdMeet("nNowStatus") = "" Then
                        Txt_nNowStatus.Text = RdMeet("nNowStatus")
                    End If
                End If
                If Not IsDBNull(RdMeet("nFinalDate")) Then
                    Txt_nFINALDATE.Text = RdMeet("nFinalDate")
                End If
                If Not IsDBNull(RdMeet("nResult")) Then
                    If Not RdMeet("nResult") = "" Then
                        Txt_nResult.Text = RdMeet("nResult")
                    End If
                End If
            End If
            db.Close()

            If strStatus = "2w" Then '只有在派工單位現勘時才可填寫
                send.Text = "修繕"
                prtP1btn.Visible = True
                Pl_2w.Visible = True
            Else
                Pl_2w.Visible = False
                DDViewPer.Enabled = False
                Txt_nViewDate.Enabled = False
                Txt_nStartDATE.Enabled = False
                ImgBtn_View.Visible = False
                ImgBtn_SDate.Visible = False
                Txt_nCause.Enabled = False
                Txt_nPacthCount.Enabled = False
                Txt_nPacthPer.Enabled = False
                Txt_nFacilityNo.Enabled = False
                DDL_nErrCause.Enabled = False
                rdo_nExternal.Enabled = False
                Button2.Visible = False
            End If
            If strStatus = "3w" Then '只有在派工單位完工時才可填寫
                send.Text = "完工"
                prtP2btn.Visible = True
                Pl_2w.Visible = True
                Pl_3w.Visible = True
            Else

                Txt_nNowStatus.Enabled = False
                Txt_nFINALDATE.Enabled = False
                Txt_nResult.Enabled = False
                btn_get.Visible = False
                ImgBtn_EDate.Visible = False
            End If

            SqlDataSource5.SelectCommand = "select DISTINCT P_0407.it_code,P_0407.it_name,P_0407.it_unit from P_0414" & _
                " join P_0407 on P_0407.it_code = substring(P_0414.shcode,0,7)" & _
                " where Job_Num ='" + streformsn + "' and shtype='0'"

            If strStatus = "1w" And getReject(streformsn) = "true" And strStepSid = "1" Then '退回至申請人的處理
                EndBtn.Visible = True
                backBtn.Visible = False
                send.Visible = False
            End If

            If strStepSid = "1032" Then '房舍水電修繕管制單位要能修改內外派
                rdo_nExternal.Enabled = True
                prtP1btn.Visible = False
                DDViewPer.Enabled = False
                Txt_nViewDate.Enabled = False
                Txt_nStartDATE.Enabled = False
                ImgBtn_View.Visible = False
                ImgBtn_SDate.Visible = False
                Txt_nCause.Enabled = False
                Txt_nPacthCount.Enabled = False
                Txt_nPacthPer.Enabled = False
                Txt_nFacilityNo.Enabled = False
                DDL_nErrCause.Enabled = False
                Button2.Visible = False
            End If

            If strStepSid = "1182" Then
                Txt_nNowStatus.Enabled = False
                Txt_nFINALDATE.Enabled = False
                Txt_nResult.Enabled = False
                btn_get.Visible = False
                ImgBtn_EDate.Visible = False
            End If

            If strStepSid = "1135" Then
                Txt_nNowStatus.Enabled = False
                Txt_nFINALDATE.Enabled = False
                Txt_nResult.Enabled = False
            End If

            If strStepSid = "21080" Then '派工單位填寫完工單時可用暫存
                SaveBtn.Visible = True
            End If
        Catch ex As Exception
            lbMsg.Text = "420:" + ex.Message
        End Try
    End Function

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Div_nFacilityNo.Visible = True
    End Sub

    Protected Sub btn_Close_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Div_nFacilityNo.Visible = False
    End Sub

    Protected Sub send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles send.Click
        Dim sPhone As String = Txt_nPHONE.Text.Trim()
        Dim s3FL As String = DDLFL.SelectedItem.Value
        Dim sRNUM As String = String.Empty
        Dim sViewPer As String = String.Empty
        Dim sViewDate As String = String.Empty
        Dim sCause As String = String.Empty
        Dim iPacthCount As Int32 = 0
        Dim checkErrorMsg As String = String.Empty
        Dim sPacthPer As String = String.Empty
        Dim sFacilityNo As String = String.Empty
        Dim sStartDATE As String = String.Empty
        Dim sNowStatus As String = String.Empty
        Dim sFINALDATE As String = String.Empty
        Dim sResult As String = String.Empty

        If Lab_PAUNIT.Text = "" Then
            AlertMsg("申請人姓名必須選取!!")
            Exit Sub
        End If
        If Txt_nPHONE.Text = "" Then
            AlertMsg("電話必須輸入!!")
            Exit Sub
        End If

        If (DDLRNUM.Items.Count > 0) Then
            If DDLRNUM.SelectedItem.Value <> "" Then
                sRNUM = DDLRNUM.SelectedItem.Value
            Else
                AlertMsg("請修地點選擇未完成哦~~")
                Exit Sub
            End If
        Else
            AlertMsg("請修地點選擇未完成哦~")
            Exit Sub
        End If

        If Txt_nFIXITEM.Text = "" Then
            AlertMsg("請修事項必須填寫哦!!")
            Exit Sub
        End If

        '判斷哪些階段的單位需要做表單 "1w" 報修單 ; "2w" 派工單 ;"3w" 完工單"
        Dim sWriteFromFlag As String = getWriteFromFlag(streformsn, user_id)

        If sWriteFromFlag = "2w" Then
            If DDViewPer.SelectedItem.Value <> "" Then
                sViewPer = DDViewPer.SelectedItem.Value
            Else
                AlertMsg("請選擇現勘人員~")
                Exit Sub
            End If
            If Txt_nViewDate.Text.Trim() <> "" Then
                sViewDate = Txt_nViewDate.Text.Trim()
                Try
                    Dim dtViewDate As DateTime = DateTime.Parse(sViewDate + " 23:59:59")
                    If (dtViewDate < DateTime.Now) Then
                        AlertMsg("您選擇現勘日期不可以早於今日哦~")
                        Exit Sub
                    End If
                Catch ex As Exception
                    AlertMsg("現勘日期格式不正確哦~")
                    Exit Sub
                End Try
            Else
                AlertMsg("請選擇現勘日期~")
                Exit Sub
            End If
            If Txt_nCause.Text.Trim() <> "" Then
                sCause = Txt_nCause.Text.Trim()
            Else
                AlertMsg("請輸入原因分析~")
                Exit Sub
            End If
            If Txt_nPacthCount.Text.Trim() <> "" Then
                Try
                    iPacthCount = Integer.Parse(Txt_nPacthCount.Text.Trim())
                Catch ex As System.Exception
                    iPacthCount = -1
                End Try
            End If
            If iPacthCount = -1 Or iPacthCount <= 0 Then
                AlertMsg("您輸入的派工人數非正確的大於0的數字哦~")
                Exit Sub
            End If

            If iPacthCount > 0 And Txt_nPacthPer.Text.Trim() <> "" Then
                sPacthPer = Txt_nPacthPer.Text.Trim()
            Else
                AlertMsg("當您有輸入派工人數時，請輸入派工人員~")
                Exit Sub
            End If
            If Txt_nFacilityNo.Text.Trim() <> "" Then
                sFacilityNo = Txt_nFacilityNo.Text.Trim()
            Else
                AlertMsg("請輸入設施(備)編號~")
                Exit Sub
            End If
            If Txt_nStartDATE.Text.Trim() <> "" Then
                sStartDATE = Txt_nStartDATE.Text.Trim()
                sViewDate = Txt_nViewDate.Text.Trim()
                Try
                    Dim dtViewDate As DateTime = DateTime.Parse(sViewDate + " 23:59:59")
                    Dim dtStartDATE As DateTime = DateTime.Parse(sStartDATE + " 23:59:59")
                    If (dtStartDATE < dtViewDate) Then
                        AlertMsg("您選擇的開工日期不可以早於現勘日期哦~")
                    End If
                Catch ex As Exception
                    lbMsg.Text = ("開工日期格式不正確哦~")
                End Try

            Else
                AlertMsg("請選擇預計開工日期~")
                Exit Sub
            End If
        End If

        If sWriteFromFlag = "3w" Then
            If Txt_nNowStatus.Text.Trim() <> "" Then
                sNowStatus = Txt_nNowStatus.Text.Trim()
            Else
                AlertMsg("請輸入目前現況~")
                Exit Sub
            End If
            If Txt_nFINALDATE.Text.Trim() <> "" Then
                sFINALDATE = Txt_nFINALDATE.Text.Trim()
                sStartDATE = Txt_nStartDATE.Text.Trim()
                Try
                    Dim dtFINALDATE As DateTime = DateTime.Parse(sFINALDATE + " 23:59:59")
                    Dim dtStartDATE As DateTime = DateTime.Parse(sStartDATE + " 23:59:59")
                    If (dtFINALDATE < dtStartDATE) Then
                        AlertMsg("您選擇的完工日期不可以早於開工日期哦~")
                        Exit Sub
                    End If
                Catch ex As Exception
                    lbMsg.Text = ("完工日期格式不正確哦~")
                End Try
            Else
                AlertMsg("請選擇完工日期~")
                Exit Sub
            End If
            If Txt_nResult.Text.Trim() <> "" Then
                sResult = Txt_nResult.Text.Trim()
            Else
                AlertMsg("請輸入處理結果~")
                Exit Sub
            End If
        End If
        Try
            Dim stmt As String = ""
            Dim FC As New C_FlowSend.C_FlowSend
            Dim FCC As New CFlowSend
            Dim SendVal As String = ""
            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                'SendVal = eformid & "," & user_id & "," & streformsn & "," & "1" & ","
                SendVal = eformid & "," & DrDown_PANAME.SelectedValue & "," & streformsn & "," & "1" & ","
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1" & ","
            End If
            Dim NextPer As Integer = 0
            '開啟連線
            Dim db As New SqlConnection(connstr)

            '關卡為上一級主管有多少人
            NextPer = FC.F_NextStep(SendVal, connstr)
            Dim Chknextstep As Integer
            '判斷表單關卡
            db.Open()
            Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & streformsn & "' and empuid = '" & user_id & "' and hddate is null", db)
            Dim RdrCheck As IDataReader = strComCheck.ExecuteReader()
            If RdrCheck.Read() Then
                Chknextstep = RdrCheck.Item("nextstep")
            End If
            db.Close()
            Dim strgroup_id As String = ""
            If read_only <> "" Then
                '判斷上一級主管
                db.Open()
                Dim strUpComCheck As New SqlCommand("SELECT group_id FROM flow WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ", db)
                Dim RdrUpCheck As IDataReader = strUpComCheck.ExecuteReader()
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
            Dim iCmdResult As Int16 = 0
            Dim ComDB As New SqlCommand
            If read_only = "2" Then
                '判斷狀態
                Dim strFlowStatus As String = getFlowStatus(streformsn)
                '1w 中 FlowStatus 會由1-->2 ; 3w 中  FlowStatus 會由3-->4 
                If (sWriteFromFlag = "1w" And strFlowStatus <> "1") Or (sWriteFromFlag = "3w" And strFlowStatus = "4") Then
                    stmt = "update P_0415 set FlowStatus=@FlowStatus where EFORMSN=@EFORMSN"
                    ComDB.Parameters.Add(New SqlParameter("@FlowStatus", SqlDbType.VarChar, 10)).Value = strFlowStatus
                    ComDB.Parameters.Add(New SqlParameter("@EFORMSN", SqlDbType.VarChar, 16)).Value = streformsn
                    iCmdResult = 1
                ElseIf (getstepsid(streformsn, user_id) <> "1032" And sWriteFromFlag = "2w") Or (getstepsid(streformsn, user_id) = "1032" And sWriteFromFlag = "2w") Then '未完成派工單位送單至房舍修繕單位
                    stmt = "update P_0415 set nExternal=@nExternal,nStartDATE=@nStartDATE" '承辦類別,開工時間
                    stmt += ",nFacilityNo=@nFacilityNo,nPacthCount=@nPacthCount" '設施編號,派員人數
                    stmt += ",nPacthPer=@nPacthPer,FlowStatus=@FlowStatus,nCause=@nCause" '派工人員,流程狀態,原因分析
                    stmt += ",nErrCause=@nErrCause,nViewPer=@nViewPer,nViewDate=@nViewDate" '故障原因,現勘人員,現勘日期
                    stmt += " where EFORMSN=@EFORMSN"
                    ComDB.Parameters.Add(New SqlParameter("@nExternal", SqlDbType.VarChar, 10)).Value = rdo_nExternal.SelectedValue
                    ComDB.Parameters.Add(New SqlParameter("@nStartDATE", SqlDbType.VarChar, 10)).Value = sStartDATE
                    ComDB.Parameters.Add(New SqlParameter("@nFacilityNo", SqlDbType.VarChar, 50)).Value = sFacilityNo
                    ComDB.Parameters.Add(New SqlParameter("@nPacthCount", SqlDbType.VarChar, 10)).Value = iPacthCount.ToString()
                    ComDB.Parameters.Add(New SqlParameter("@nPacthPer", SqlDbType.NVarChar, 100)).Value = sPacthPer
                    ComDB.Parameters.Add(New SqlParameter("@FlowStatus", SqlDbType.VarChar, 10)).Value = strFlowStatus
                    ComDB.Parameters.Add(New SqlParameter("@nCause", SqlDbType.VarChar, 255)).Value = sCause
                    ComDB.Parameters.Add(New SqlParameter("@nErrCause", SqlDbType.VarChar, 100)).Value = DDL_nErrCause.SelectedValue
                    ComDB.Parameters.Add(New SqlParameter("@nViewPer", SqlDbType.NVarChar, 2000)).Value = DDViewPer.SelectedValue
                    ComDB.Parameters.Add(New SqlParameter("@nViewDate", SqlDbType.VarChar)).Value = Txt_nViewDate.Text
                    ComDB.Parameters.Add(New SqlParameter("@EFORMSN", SqlDbType.VarChar, 16)).Value = streformsn
                    iCmdResult = 1
                ElseIf sWriteFromFlag = "3w" And strFlowStatus = "3" Then
                    stmt = "update P_0415 set "
                    stmt += "nNowStatus=@nNowStatus,nFinalDate=@nFinalDate,nResult=@nResult,FlowStatus=@FlowStatus" '目前現況,完工日期,處理結果,更新狀態"
                    stmt += " where EFORMSN=@EFORMSN"
                    ComDB.Parameters.Add(New SqlParameter("@nNowStatus", SqlDbType.VarChar, 100)).Value = sNowStatus
                    ComDB.Parameters.Add(New SqlParameter("@nFinalDate", SqlDbType.VarChar)).Value = sFINALDATE
                    ComDB.Parameters.Add(New SqlParameter("@nResult", SqlDbType.NVarChar, 2000)).Value = sResult
                    ComDB.Parameters.Add(New SqlParameter("@FlowStatus", SqlDbType.VarChar, 10)).Value = strFlowStatus
                    ComDB.Parameters.Add(New SqlParameter("@EFORMSN", SqlDbType.VarChar, 16)).Value = streformsn
                    iCmdResult = 1
                End If
            Else
                stmt = "insert into P_0415(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
                stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
                stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
                stmt += "nAPPTIME, nPHONE, " '申請時間,聯絡電話
                stmt += "nbd_code,nfl_code,nrnum_code,nFIXITEM) " '請修地點,請修事項  
                stmt += " values(@EFORMSN,@PWUNIT,@PWTITLE,@PWNAME, @PWIDNO, @PAUNIT, @PANAME, @PATITLE, "
                stmt += "@PAIDNO,@nAPPTIME, @nPHONE,@nbd_code,@nfl_code,@nrnum_code,@nFIXITEM)"
                ComDB.Parameters.Add(New SqlParameter("@EFORMSN", SqlDbType.VarChar, 16)).Value = streformsn
                ComDB.Parameters.Add(New SqlParameter("@PWUNIT", SqlDbType.VarChar, 50)).Value = Lab_PWUNIT.Text
                ComDB.Parameters.Add(New SqlParameter("@PWTITLE", SqlDbType.NVarChar, 50)).Value = Lab_PWTITLE.Text
                ComDB.Parameters.Add(New SqlParameter("@PWNAME", SqlDbType.VarChar, 20)).Value = Lab_PWNAME.Text
                ComDB.Parameters.Add(New SqlParameter("@PWIDNO", SqlDbType.VarChar, 10)).Value = user_id
                ComDB.Parameters.Add(New SqlParameter("@PAUNIT", SqlDbType.VarChar, 50)).Value = Lab_PAUNIT.Text
                ComDB.Parameters.Add(New SqlParameter("@PANAME", SqlDbType.VarChar, 20)).Value = DrDown_PANAME.SelectedItem.Text
                ComDB.Parameters.Add(New SqlParameter("@PATITLE", SqlDbType.NVarChar, 50)).Value = Lab_PATITLE.Text
                ComDB.Parameters.Add(New SqlParameter("@PAIDNO", SqlDbType.NVarChar, 50)).Value = DrDown_PANAME.SelectedItem.Value
                ComDB.Parameters.Add(New SqlParameter("@nAPPTIME", SqlDbType.NVarChar, 50)).Value = AppDate.Text
                ComDB.Parameters.Add(New SqlParameter("@nPHONE", SqlDbType.NVarChar, 50)).Value = Txt_nPHONE.Text
                ComDB.Parameters.Add(New SqlParameter("@nbd_code", SqlDbType.VarChar, 1)).Value = DDLBD.SelectedItem.Value
                ComDB.Parameters.Add(New SqlParameter("@nfl_code", SqlDbType.VarChar, 2)).Value = DDLFL.SelectedItem.Value
                ComDB.Parameters.Add(New SqlParameter("@nrnum_code", SqlDbType.VarChar, 5)).Value = DDLRNUM.SelectedItem.Value
                ComDB.Parameters.Add(New SqlParameter("@nFIXITEM", SqlDbType.NVarChar, 2000)).Value = Txt_nFIXITEM.Text
                iCmdResult = 1
            End If
            If iCmdResult = 1 Then
                Try
                    ComDB.CommandText = stmt
                    ComDB.Connection = db
                    ComDB.Connection.Open()
                    ComDB.ExecuteNonQuery()
                    iCmdResult = 2
        
                Catch ex As Exception
                    iCmdResult = -1
                Finally
                    If ComDB.Connection.State.Equals(ConnectionState.Open) Then
                        ComDB.Connection.Close()
                    End If
                    ComDB.Dispose()
                    ComDB = Nothing
                End Try

                If iCmdResult = -1 Then
                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('異動資料庫失敗,請重新再試或連絡資訊管理人員!!');")
                    Response.Write(" </script>")
                End If

            End If

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
                    Dim RdPer As IDataReader = strPer.ExecuteReader()
                    If RdPer.Read() Then
                        strAgentName = RdPer("emp_chinese_name")
                    End If
                    db.Close()
                    strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)"
                    '增加批核意見
                    insComment(strComment, streformsn, AgentEmpuid)
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
                Dim OAPerRdr As IDataReader = OAPerCom.ExecuteReader()
                If OAPerRdr.Read() Then
                    OAPer = OAPerRdr.Item("OAPer")
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
                    Dim OAPerAllRdr As IDataReader = OAPerAllCom.ExecuteReader()
                    If OAPerAllRdr.Read() Then
                        OAPerAll = OAPerAllRdr.Item("OAPerAll")
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
                        Dim OAPerOneRdr As IDataReader = OAPerOneCom.ExecuteReader()
                        If OAPerOneRdr.Read() Then
                            OAPerORG = OAPerOneRdr.Item("employee_id")
                        End If
                        db.Close()

                        'Dim SendValOA = eformid & "," & do_sql.G_user_id & "," & streformsn & "," & "1" & "," & OAPerORG
                        Dim SendValOA As String = eformid & "," & DrDown_PANAME.SelectedValue & "," & streformsn & "," & "1" & "," & OAPerORG

                        '表單審核
                        Val_P = FCC.F_Send(SendValOA, do_sql.G_conn_string)
                        If read_only = "" Then
                            PageUp = "New"
                            Session("PageUp") = "Done"
                        End If
                        Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                    ElseIf OAPerAll > 0 Then
                        Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & streformsn & "&SelFlag=3&strOrgTop=" & strOrgTop & "&AppUID=" & DrDown_PANAME.SelectedValue)
                    End If
                ElseIf OAPer = 1 Then
                    '表單審核
                    Val_P = FCC.F_Send(SendVal, do_sql.G_conn_string)
                    If read_only = "" Then
                        PageUp = "New"
                    Else
                        Session("PageUp") = "Done"
                    End If
                    Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                ElseIf OAPer > 1 Then
                    Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & streformsn & "&SelFlag=2" & "&AppUID=" & DrDown_PANAME.SelectedValue)
                End If
            Else
                '判斷下一關為上一級主管時人數是否超過一人
                If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
                    Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & streformsn & "&SelFlag=1" & "&AppUID=" & DrDown_PANAME.SelectedValue)
                Else
                    '表單審核
                    Val_P = FCC.F_Send(SendVal, do_sql.G_conn_string)
                    If read_only = "" Then
                        PageUp = "New"
                    Else
                        Session("PageUp") = "Done"
                    End If
                    Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                End If
            End If

        Catch ex As Exception
            lbMsg.Text = "858" + ex.Message
        End Try

    End Sub

    Public Function getFlowStatus(ByVal _streformsn As String) As String
        '流程狀態
        Dim strFlowStatus As String = ""
        Dim _strstepsid As String = ""
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        Dim strstepChk As New SqlCommand("select stepsid from dbo.flowctl where eformsn='" & _streformsn & "' and gonogo ='?'", db)
        Dim RdrstepChk As IDataReader = strstepChk.ExecuteReader()

        If RdrstepChk.read() Then
            _strstepsid = RdrstepChk.item("stepsid")
        End If
        db.Close()

        Select Case _strstepsid
            Case "29468"
                strFlowStatus = "1" '新送件
            Case "21079", "21027", "1032", "19420"
                strFlowStatus = "2" '處理中
            Case "21080"
                strFlowStatus = "3" '待料中
            Case "1182", "1135"
                strFlowStatus = "4" '完工
        End Select
        Return strFlowStatus

    End Function
    Public Function getStepChk(ByVal _streformsn As String, ByVal _empuid As String) As Integer
        '不同階段顯示不同表單
        Dim StepFlag As Integer
        Dim _strstepsid As String = ""
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        Dim strstepChk As New SqlCommand("select stepsid from dbo.flowctl where empuid='" & _empuid & "' and eformsn='" & _streformsn & "' and gonogo ='?'", db)
        Dim RdrstepChk As IDataReader = strstepChk.ExecuteReader()

        If RdrstepChk.read() Then
            _strstepsid = RdrstepChk.item("stepsid")
        End If
        db.Close()
        db.Open()
        If _strstepsid = "" Then
            Dim strSql As String = "select stepsid from flowctl where flowsn = (select max(flowsn) flowsn  from flowctl where eformsn='" & _streformsn & "' and gonogo <> '?')"
            Dim _strstepChk As New SqlCommand(strSql, db)
            RdrstepChk = _strstepChk.ExecuteReader()

            If RdrstepChk.read() Then
                _strstepsid = RdrstepChk.item("stepsid")
            End If
        End If
        db.Close()

        Select Case _strstepsid
            Case "21027", "1032", "21079"
                StepFlag = 2 '顯示派工單區域
            Case "21080", "1182", "1135"
                StepFlag = 3 '顯示完工單區域
            Case "19420", "29468"
                StepFlag = 1 '顯示報修單區域
        End Select
        Return StepFlag
    End Function
    Public Function getNewApp(ByVal _streformsn As String) As String
        Dim conn As New C_SQLFUN
        Dim flowsn As String
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        Dim dt As New SqlCommand("select flowsn from flowctl where eformsn='" & _streformsn & "'", db)
        Dim dr = dt.ExecuteReader()

        If dr.read() Then
            flowsn = dr.item("flowsn")
        End If
        db.Close()
        If flowsn <> "" Then
            Return "false"
        Else
            Return "true"
        End If

    End Function
    Public Function getWriteFromFlag(ByVal _streformsn As String, ByVal _empuid As String) As String
        '判斷哪些階段的單位需要做表單
        Dim ConfirmFlag As String
        Dim _strstepsid As String = ""

        _strstepsid = getstepsid(_streformsn, _empuid)
        ConfirmFlag = ""
        Select Case _strstepsid
            Case "21027", "1032"
                ConfirmFlag = "2w" '派工單
            Case "21080", "1182", "1135", "2"
                ConfirmFlag = "3w" '完工單
            Case "1", "29468", "19420", "21079"
                ConfirmFlag = "1w" '報修單
        End Select
        Return ConfirmFlag
    End Function

    Public Function getConfirmFlag(ByVal _streformsn As String, ByVal _empuid As String) As String
        '判斷哪些階段的單位需要做批核意見
        Dim ConfirmFlag As String = ""
        Dim _strstepsid As String = ""

        _strstepsid = getstepsid(_streformsn, _empuid)

        Select Case _strstepsid
            Case "29468", "1032", "19420", "21027", "21080"
                ConfirmFlag = "0" '不填寫批核意見
            Case "21079", "1182", "1135"
                ConfirmFlag = "1" '填寫批核意見
        End Select
        Return ConfirmFlag

    End Function

    Public Sub setConfirmBtn(ByVal _streformsn As String, ByVal _empuid As String)
        Dim _strstepsid As String = ""
        Dim ConfirmFlag As String = ""
        _strstepsid = getstepsid(_streformsn, _empuid)
        Select Case _strstepsid
            Case "21027"
                send.Text = "修繕"
                backBtn.Text = "不修繕"
            Case "1032"
                send.Text = "修繕"
                backBtn.Visible = False
            Case "21080"
                send.Text = "完工"
                backBtn.Visible = False
            Case "1182"
                prtP2btn.Visible = False
                send.Text = "核准"
            Case "1135"
                prtP2btn.Visible = False
                send.Visible = False
                EndBtn.Visible = True
                backBtn.Visible = False
            Case ""
                'send.Visible = False
                backBtn.Visible = False
        End Select

    End Sub

    Public Function getstepsid(ByVal _streformsn As String, ByVal _empuid As String) As String
        Dim _strstepsid As String = ""
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        Dim strstepChk As New SqlCommand("select top 1 stepsid from dbo.flowctl where empuid='" & _empuid & "' and eformsn='" & _streformsn & "' and gonogo ='?' order by createdate desc", db)
        Dim RdrstepChk As IDataReader = strstepChk.ExecuteReader()

        If RdrstepChk.Read() Then
            _strstepsid = RdrstepChk.Item("stepsid")
        End If
        db.Close()
        Return _strstepsid
    End Function
    Public Function getReject(ByVal _streformsn As String) As String
        Dim conn As New C_SQLFUN
        Dim strSql, strgonogo As String
        connstr = conn.G_conn_string

        Dim db As New SqlConnection(connstr)

        db.Open()
        strSql = "select * from " & _
                "(select ROW_NUMBER() OVER (order by flowsn) as RowNO,* from flowctl " & _
                "where eformsn='" & _streformsn & "' ) as b " & _
                "where b.RowNO= " & _
                "(select max(RNo)-1 from " & _
                "(select ROW_NUMBER() OVER (order by flowsn) as RNO,* from flowctl " & _
                "where eformsn='" & _streformsn & "') as a)"
        Dim dt As New SqlCommand(strSql, db)
        Dim dr As IDataReader = dt.ExecuteReader()

        If dr.Read() Then
            strgonogo = dr.Item("gonogo")
        End If
        db.Close()
        If strgonogo = "0" Then
            Return "true"
        Else
            Return "false"
        End If
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
            lbMsg.Text = ex.Message
        End Try
        insComment = ""
    End Function

    Protected Sub DDLelement_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLelement.SelectedIndexChanged

        Txt_nFacilityNo.Text = DDLelement.SelectedItem.Text
        Div_nFacilityNo.Visible = False

    End Sub

    Protected Sub btn_get_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_get.Click

        Dim strPath As String = ""
        Dim itcode As String = Mid(Trim(Txt_nFacilityNo.Text), 12, 3)
        strPath = "MOA04101.aspx?eformsn=" & streformsn & "&itcode=" & itcode

        Response.Write(" <script language='javascript'>")
        Response.Write(" sPath = '" & strPath & "';")
        Response.Write(" strFeatures = 'dialogWidth=1000px;dialogHeight=600px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
        Response.Write(" showModalDialog(sPath,self,strFeatures);")
        Response.Write(" </script>")

        ShowDetail(streformsn)
    End Sub

    Protected Sub ImgBtn_View_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn_View.Click

        DivView_grid.Visible = True
        DivView_grid.Style("Top") = "330px"
        DivView_grid.Style("left") = "520px"

    End Sub

    Protected Sub ImgBtn_SDate_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn_SDate.Click

        DivSDate_grid.Visible = True
        DivSDate_grid.Style("Top") = "415px"
        DivSDate_grid.Style("left") = "520px"

    End Sub

    Protected Sub ImgBtn_EDate_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn_EDate.Click

        DivEDate_grid.Visible = True
        DivEDate_grid.Style("Top") = "480px"
        DivEDate_grid.Style("left") = "520px"
        send.Text = "完工"
    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        DivView_grid.Visible = False

    End Sub
    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        DivSDate_grid.Visible = False

    End Sub
    Protected Sub btnClose3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose3.Click

        DivEDate_grid.Visible = False
        send.Text = "完工"
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Txt_nViewDate.Text = Calendar1.SelectedDate.Date
        DivView_grid.Visible = False

    End Sub
    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        Txt_nStartDATE.Text = Calendar2.SelectedDate.Date
        DivSDate_grid.Visible = False

    End Sub
    Protected Sub Calendar3_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar3.SelectionChanged

        Txt_nFINALDATE.Text = Calendar3.SelectedDate.Date
        DivEDate_grid.Visible = False
        send.Text = "完工"

    End Sub

    Function Check_All() As Boolean

        Dim Check_Flag As String = ""

        If Txt_nPHONE.Text = "" Then
            Check_Flag = "1"
        End If

        If DDLBD.SelectedItem.Value = "" Then
            Check_Flag = "1"
        End If

        If DDLFL.SelectedItem.Value = "" Then
            Check_Flag = "1"
        End If

        If DDLRNUM.SelectedItem.Value = "" Then
            Check_Flag = "1"
        End If

        If getWriteFromFlag(streformsn, user_id) = "2w" Then
            If DDViewPer.SelectedItem.Value = "" Then
                Check_Flag = "1"
            End If
            If Txt_nViewDate.Text = "" Then
                Check_Flag = "1"
            End If
            If Txt_nCause.Text = "" Then
                Check_Flag = "1"
            End If
            If Txt_nPacthCount.Text = "" Then
                Check_Flag = "1"
            End If
            If Txt_nPacthPer.Text = "" Then
                Check_Flag = "1"
            End If
            If Txt_nFacilityNo.Text = "" Then
                Check_Flag = "1"
            End If
            If Txt_nStartDATE.Text = "" Then
                Check_Flag = "1"
            End If
        End If

        If getWriteFromFlag(streformsn, user_id) = "3w" Then
            If Txt_nNowStatus.Text = "" Then
                Check_Flag = "1"
            End If
            If Txt_nFINALDATE.Text = "" Then
                Check_Flag = "1"
            End If
            If Txt_nResult.Text = "" Then
                Check_Flag = "1"
            End If
        End If

        If Check_Flag = "" Then
            Check_All = True
        End If

    End Function

    Protected Sub DDLFL_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLFL.SelectedIndexChanged

        Dim db As New SqlConnection(connstr)
        Dim strSQL As String = "SELECT *,rnum_code+'-'+rnum_name as ShowName FROM [P_0411] WHERE (([bd_code] = '" + DDLBD.SelectedValue + "') AND ([fl_code] = '" + DDLFL.SelectedValue + "')) ORDER BY [rnum_name]"
        db.Open()
        Dim dt As New SqlCommand(strSQL, db)
        Dim Rdrnum = dt.ExecuteReader()
        DDLRNUM.Items.Clear()
        Do While Rdrnum.Read()
            DDLRNUM.Items.Add(New ListItem(Rdrnum("ShowName"), Rdrnum("rnum_code")))
        Loop
        db.Close()
    End Sub

    Protected Sub DDLBD_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLBD.SelectedIndexChanged
        Dim db As New SqlConnection(connstr)
        Dim strSQL As String = "SELECT *,rnum_code+'-'+rnum_name as ShowName FROM [P_0411] WHERE (([bd_code] = '" + DDLBD.SelectedValue + "') AND ([fl_code] = '" + DDLFL.SelectedValue + "')) ORDER BY [rnum_name]"
        db.Open()
        Dim dt As New SqlCommand(strSQL, db)
        Dim Rdrnum = dt.ExecuteReader()
        DDLRNUM.Items.Clear()
        Do While Rdrnum.Read()
            DDLRNUM.Items.Add(New ListItem(Rdrnum("ShowName"), Rdrnum("rnum_code")))
        Loop
        db.Close()
    End Sub

    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles backBtn.Click
        Try
            Dim strAgentName As String
            Dim db As New SqlConnection(connstr)
            If read_only = "2" Then
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
                    If RdPer.read() Then
                        strAgentName = RdPer("emp_chinese_name")
                    End If
                    db.Close()
                    strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)"
                    '增加批核意見
                    insComment(strComment, streformsn, AgentEmpuid)
                End If
            End If

            Dim Val_P As String
            Val_P = ""

            '表單駁回
            Dim FCC As New CFlowSend

            Dim SendVal As String = ""

            Dim backNum As String

            If getstepsid(streformsn, user_id) = "21079" Then
                backNum = "2"
            Else
                backNum = "1"
            End If

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & streformsn & "," & "1" & "," & backNum
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1" & "," & backNum
            End If

            Val_P = FCC.F_BackM(SendVal, connstr)

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單已駁回');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")
        Catch ex As Exception
            lbMsg.Text = ex.Message
        End Try
    End Sub

    Protected Sub EndBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EndBtn.Click
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        Dim db As New SqlConnection(connstr)
        Dim strSQL As String
        strSQL = "update flowctl set gonogo='E',hddate=getdate(),nextstep=-1,comment = '" & txtcomment.Text.Trim() & _
                  "' where eformsn='" & streformsn & "' and empuid='" & user_id & "' "
        db.Open()
        Dim dt As New SqlCommand(strSQL, db)
        dt.ExecuteNonQuery()
        db.Close()
        '重新整理頁面
        Response.Write("<script language='javascript'>")
        Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
        Response.Write("window.close();")
        Response.Write("</script>")
    End Sub

    Protected Sub GridView11_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView11.RowDataBound
        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then
            Dim lblIt_Use As Label
            Dim Hidit_code As HiddenField
            Dim lblIt_App As Label
            Dim strSQL As String
            Dim num As String

            lblIt_App = e.Row.Cells(2).FindControl("lblIt_App")
            lblIt_Use = e.Row.Cells(4).FindControl("lblIt_Use")
            Hidit_code = e.Row.Cells(4).FindControl("Hidit_code")

            strSQL = "select count(*) as usenum from P_0414 " & _
                        "where substring(P_0414.shcode,0,7)='" + Hidit_code.Value + "' " & _
                        "and UseCheck='2' and shtype='0' " & _
                        "and Job_Num='" + streformsn + "'"
            num = 0
            Dim db As New SqlConnection(connstr)
            db.Open()
            Dim dt As New SqlCommand(strSQL, db)
            Dim dr = dt.ExecuteReader()
            If dr.read() Then
                num = dr("usenum").ToString()
            End If
            db.Close()
            lblIt_Use.Text = num


            strSQL = "select count(*) as appnum from P_0414 " & _
                        "where substring(P_0414.shcode,0,7)='" + Hidit_code.Value + "' " & _
                        "and shtype='0' " & _
                        "and Job_Num='" + streformsn + "'"
            num = 0
            db.Open()
            Dim dt1 As New SqlCommand(strSQL, db)
            Dim dr1 = dt1.ExecuteReader()
            If dr1.read() Then
                num = dr1("appnum").ToString()
            End If
            db.Close()
            lblIt_App.Text = num

        End If
    End Sub

    Protected Sub SaveBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveBtn.Click
        Dim db As New SqlConnection(connstr)
        Dim strSql As String
        db.Open()
        Try
            If Txt_nFINALDATE.Text = "" Then
                strSql = "update P_0415 set nNowStatus='" + Txt_nNowStatus.Text + "',nFinalDate=null,nResult='" + Txt_nResult.Text + "' " & _
                            "where EFORMSN='" + streformsn + "' "
            Else
                strSql = "update P_0415 set nNowStatus='" + Txt_nNowStatus.Text + "',nFinalDate='" + Txt_nFINALDATE.Text + "',nResult='" + Txt_nResult.Text + "' " & _
                             "where EFORMSN='" + streformsn + "' "
            End If
            Dim dt As New SqlCommand(strSql, db)
            dt.ExecuteNonQuery()
            db.Close()
            strSql = String.Empty
        Catch ex As Exception
            strSql = ex.Message
        End Try
        If strSql = String.Empty Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('此筆資料已為您暫存成功!');")
            Response.Write(" </script>")
        Else
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('" + strSql + "');")
            Response.Write(" </script>")
        End If
        send.Text = "完工"
    End Sub

    Protected Sub prtP1btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles prtP1btn.Click
        Response.Redirect("MOA04102.aspx?streformsn=" + streformsn, True)
    End Sub

    Protected Sub prtP2btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles prtP2btn.Click
        Response.Redirect("MOA04103.aspx?streformsn=" + streformsn, True)
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
        Else

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('無個人批核片語資料,請至批核片語管理新增資料');")
            Response.Write(" </script>")

        End If
    End Sub

    Protected Sub Btn_PHclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_PHclose.Click
        Div_grid10.Visible = False
    End Sub
    Protected Sub GridView10_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView10.SelectedIndexChanged
        txtcomment.Text = GridView10.Rows(GridView10.SelectedRow.RowIndex).Cells(0).Text
        Div_grid10.Visible = False
    End Sub

    Private Sub AlertMsg(ByVal alertMsgString As String)
        Response.Write(" <script language='javascript'>")
        Response.Write(" alert('" + alertMsgString + "');")
        Response.Write(" </script>")
    End Sub
End Class
