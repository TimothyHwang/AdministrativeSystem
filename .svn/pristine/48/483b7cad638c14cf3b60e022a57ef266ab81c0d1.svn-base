
Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq


Namespace M_Source._11

    Partial Class M_Source_11_MOA11001
        Inherits Page
        Public pfilename As String = ""
        Public do_sql As New C_SQLFUN
        Public ErrMsg As String = ""
        Public read_only As String = ""
        Public boolFixAsigned As Boolean
        Public strCaseStatus As String = ""

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public flowAdmin As String
        Public user_id As String
        Public org_uid As String
        Public eformsn As String
        Public eformid As String
        Public AgentEmpuid As String
        Public AppUID As String


        Public Sub insComment(ByVal strComment As String, ByVal strEformsn As String, ByVal strid As String)

            '新增批核意見
            Try

                '開啟連線
                Dim DC As New SQLDBControl
                Dim strSql As String
                '新增批何意見
                strSql = "UPDATE flowctl SET comment='" & strComment & "' WHERE hddate IS NULL AND eformsn='" & strEformsn & "' AND empuid='" & strid & "'"
                DC.ExecuteSQL(strSql)
                DC.Dispose()

            Catch ex As Exception

            End Try
        End Sub

        Protected Function CheckApplyForm() As Boolean
            Dim boolReturn As Boolean = True

            If txtPhone.Text.Length = 0 Then
                boolReturn = False
                do_sql.G_errmsg = "電話必須輸入"
            End If

            If ddlBuilding.SelectedValue.Length = 0 Then
                boolReturn = False
                do_sql.G_errmsg = "必須選取建築物"
            End If
            If ddlLevel.SelectedValue.Length = 0 Then
                boolReturn = False
                do_sql.G_errmsg = "必須選取樓層"
            End If
            If txtRoom.Text.Length = 0 Then
                boolReturn = False
                do_sql.G_errmsg = "房間號碼必須輸入"
            End If

            If ddlRepairMainKind.SelectedValue.Length = 0 Then
                boolReturn = False
                do_sql.G_errmsg = "必須選取主要維修種類"
            End If

            If ddlProblemKind.SelectedValue.Length = 0 Then
                boolReturn = False
                do_sql.G_errmsg = "必須選取次要維修種類"
            End If

            If txtDESCRIPTION.Text.Length = 0 Then
                boolReturn = False
                do_sql.G_errmsg = "必須輸入請修事項"
            End If

            If ddlRepairMainKind.SelectedItem.Text = "網路故障" Then
                If txtIPAddr.Text.Length = 0 Or txtMACAddr.Text.Length = 0 Or txtPlugNo.Text.Length = 0 Then
                    boolReturn = False
                    do_sql.G_errmsg = "必須輸入IP位址、MAC位址及插座編號"
                End If
            End If

            Dim DC As New SQLDBControl
            Dim DR As SqlDataReader
            DR = DC.CreateReader("SELECT EFORMSN FROM P_11 WHERE EFORMSN='" + eformsn + "'")
            If DR.HasRows Then boolReturn = False Else boolReturn = True
            DC.Dispose()
            CheckApplyForm = boolReturn
        End Function

        Protected Function CheckSignForm() As Boolean
            Dim boolReturn As Boolean = True
            ''叫修時間
            Dim CallTime As DateTime
            ''到修時間
            Dim ArriveTime As DateTime
            ''完修時間
            Dim FixTime As DateTime

            If (txtCallDate.Text.Length > 0) Then CallTime = CType(txtCallDate.Text + " " + ddlCallTimeHour.SelectedValue + ":" + ddlArriveTimeMin.SelectedValue, DateTime)
            If (txtArriveDate.Text.Length > 0) Then ArriveTime = CType(txtArriveDate.Text + " " + ddlArriveTimeHour.SelectedValue + ":" + ddlArriveTimeMin.SelectedValue, DateTime)
            If (txtFinalDate.Text.Length > 0) Then FixTime = CType(txtFinalDate.Text + " " + ddlFinalTimeHour.SelectedValue + ":" + ddlFinalTimeMin.SelectedValue, DateTime)

            If (txtCallDate.Text.Length > 0 And txtArriveDate.Text.Length > 0 And txtFinalDate.Text.Length > 0) Then
                If (CallTime > ArriveTime Or CallTime > FixTime) Then
                    boolReturn = False
                ElseIf (ArriveTime > FixTime) Then
                    boolReturn = False
                End If
            ElseIf (txtCallDate.Text.Length > 0 And txtArriveDate.Text.Length > 0) Then
                If (CallTime > ArriveTime) Then
                    boolReturn = False
                End If
            ElseIf (txtCallDate.Text.Length > 0 And txtFinalDate.Text.Length > 0) Then
                If (CallTime > FixTime) Then
                    boolReturn = False
                End If
            ElseIf (txtArriveDate.Text.Length > 0 And txtFinalDate.Text.Length > 0) Then
                If (ArriveTime > FixTime) Then
                    boolReturn = False
                End If
            End If

            If (strCaseStatus = "F" Or ViewState("P11CaseStatus") = "F") And flowAdmin = "3" Then
                If txtCallDate.Text.Length = 0 Or txtArriveDate.Text.Length = 0 Or txtFinalDate.Text.Length = 0 Then
                    boolReturn = False
                End If
            End If
            CheckSignForm = boolReturn
        End Function

        Protected Function LoadData() As Boolean
            Dim DC As New SQLDBControl
            Dim DR As SqlDataReader
            Dim strSql As String
            Dim boolReturn As Boolean
            Dim tool As New C_Public

            boolReturn = False
            Try
                strSql = "SELECT * FROM P_11 WHERE EFORMSN='" & eformsn & "'"
                DR = DC.CreateReader(strSql)
                If DR.HasRows Then
                    If DR.Read Then
                        Lab_PWUNIT.Text = DR("PWUNIT").ToString
                        Lab_PWNAME.Text = DR("PWNAME").ToString
                        Lab_PWTITLE.Text = DR("PWTITLE").ToString
                        'ddlPAUNIT.SelectedValue = DR("PAIDNO").ToString
                        DrDown_PANAME.Items.Clear()
                        DrDown_PANAME.Items.Add(New ListItem(Trim(DR("PANAME").ToString), Trim(DR("PAIDNO").ToString())))
                        DrDown_PANAME.SelectedValue = Trim(DR("PAIDNO").ToString)

                        ddlPAUNIT.Items.Clear()
                        ddlPAUNIT.Items.Add(New ListItem(Trim(DR("PAUNIT").ToString), tool.GetOrgIDByIDNo(Trim(DR("PAIDNO").ToString()))))
                        ddlPAUNIT.SelectedValue = tool.GetOrgIDByIDNo(Trim(DR("PAIDNO").ToString()))

                        Lab_PATITLE.Text = DR("PATITLE").ToString
                        lblApplyTime.Text = DR("APPTIME").ToString
                        hdApplyTime.Value = DR("APPTIME").ToString

                        txtPhone.Text = DR("PHONE").ToString

                        If DR("FIXIDNO").ToString().Length > 0 Then
                            'ddlRepairMan.SelectedValue = DR("FIXIDNO").ToString()
                            lblRepairMan.Text = tool.GetUserNameByID(DR("FIXIDNO").ToString())
                            hdnRepairMan.Value = DR("FIXIDNO").ToString()
                            boolFixAsigned = True
                        Else
                            boolFixAsigned = False
                        End If
                        ViewState("P11IsUPLOADNAME") = DR("UPLOADNAME").ToString().Length > 0
                        lnkUploadFile.Text = DR("UPLOADNAME").ToString()
                        lnkUploadFile.NavigateUrl = "../DownloadFile.aspx?file=" + eformsn + "-" + DR("UPLOADNAME").ToString()

                        ViewState("P11IsUPLOADFINISHNAME") = DR("UPLOADFINISHNAME").ToString().Length > 0
                        lnkFinishUploadFile.Text = DR("UPLOADFINISHNAME").ToString()
                        lnkFinishUploadFile.NavigateUrl = "../DownloadFile.aspx?file=" + eformsn + "-Final-" + DR("UPLOADFINISHNAME").ToString()

                        If read_only = "1" Or read_only = "2" Then
                            If DR("CALLTIME").ToString = "" Then
                                txtCallDate.Text = ""
                                ddlCallTimeHour.SelectedValue = DateTime.Now.Hour.ToString()
                                ddlCallTimeMin.SelectedValue = DateTime.Now.Minute.ToString()
                            Else
                                txtCallDate.Text = CDate(DR("CALLTIME").ToString).ToString("yyyy/MM/dd")
                                ddlCallTimeHour.Text = CType(CDate(DR("CALLTIME").ToString).Hour, String)
                                ddlCallTimeMin.Text = CType(CDate(DR("CALLTIME").ToString).Minute, String)
                                ViewState("P11IsCallTime") = True
                            End If

                            If DR("ARRIVETIME").ToString = "" Then
                                txtArriveDate.Text = ""
                                ddlArriveTimeHour.SelectedValue = DateTime.Now.Hour.ToString()
                                ddlArriveTimeMin.SelectedValue = DateTime.Now.Minute.ToString()
                            Else
                                txtArriveDate.Text = CDate(DR("ARRIVETIME").ToString).ToString("yyyy/MM/dd")
                                ddlArriveTimeHour.Text = CType(CDate(DR("ARRIVETIME").ToString).Hour, String)
                                ddlArriveTimeMin.Text = CType(CDate(DR("ARRIVETIME").ToString).Minute, String)
                                ViewState("P11IsArriveTime") = True
                            End If

                            If DR("FINALDATE").ToString = "" Then
                                txtFinalDate.Text = ""
                                ddlFinalTimeHour.SelectedValue = DateTime.Now.Hour.ToString()
                                ddlFinalTimeMin.SelectedValue = DateTime.Now.Minute.ToString()
                            Else
                                txtFinalDate.Text = CDate(DR("FINALDATE").ToString).ToString("yyyy/MM/dd")
                                ddlFinalTimeHour.Text = CType(CDate(DR("FINALDATE").ToString).Hour, String)
                                ddlFinalTimeMin.Text = CType(CDate(DR("FINALDATE").ToString).Minute, String)
                                ViewState("P11IsFinalTime") = True
                            End If

                            ddlRepairMainKind.SelectedValue = DR("BROKENTYPE").ToString().Substring(0, 1)
                            ddlProblemKind.SelectedValue = DR("BROKENTYPE").ToString().Substring(1, 2)
                            txtIPAddr.Text = CType(IIf(IsDBNull(DR("IPADDRESS")), "", DR("IPADDRESS")), String)
                            txtMACAddr.Text = CType(IIf(IsDBNull(DR("MACADDRESS")), "", DR("MACADDRESS")), String)
                            txtPlugNo.Text = CType(IIf(IsDBNull(DR("PLUGPORT")), "", DR("PLUGPORT")), String)

                            ddlBuilding.SelectedValue = DR("LOCATION").ToString().Substring(0, 2)
                            ddlLevel.SelectedValue = DR("LOCATION").ToString().Substring(2, 2)
                            txtRoom.Text = DR("LOCATION").ToString().Substring(4, DR("LOCATION").ToString().Length - 4)

                            txtDESCRIPTION.Text = DR("DESCRIPTION").ToString()



                            'Select Case DR("FIXSTATUS").ToString
                            '    Case "0"
                            '        rdbStatus.Items(0).Selected = True
                            '    Case "1"
                            '        rdbStatus.Items(1).Selected = True
                            '    Case "2"
                            '        rdbStatus.Items(2).Selected = True
                            '    Case Else
                            '        rdbStatus.Items(0).Selected = True
                            'End Select

                            If DR("FIXRECORD").ToString().Length > 0 Then
                                txtFIXRECORD.Text = DR("FIXRECORD").ToString()
                                ViewState("P11IsFIXRECORD") = True
                            End If
                        End If
                        If DR("PENDFLAG").ToString().Length > 0 Then
                            ViewState("P11CaseStatus") = DR("PENDFLAG").ToString()
                            strCaseStatus = ViewState("P11CaseStatus").ToString()
                        End If
                        Call FormController()
                        boolReturn = True
                    End If
                Else
                    Lab_PWUNIT.Text = ""
                    Lab_PWNAME.Text = ""
                    Lab_PWTITLE.Text = ""
                    'Lab_PAUNIT.Text = ""
                    DrDown_PANAME.Items.Clear()
                    DrDown_PANAME.Items.Add("")
                    DrDown_PANAME.Items(0).Value = ""

                    Lab_PATITLE.Text = ""
                    Label8.Text = ""

                    txtPhone.Text = ""

                    If read_only = "2" Then
                        'Txt_nFIXDATE.Text = ""

                    End If
                    Call FormController()
                End If

                DC.Dispose()
            Catch ex As Exception
                boolReturn = True
            End Try


            LoadData = boolReturn
        End Function

        Protected Sub FormController()
            ddlPAUNIT.Enabled = False
            DrDown_PANAME.Enabled = False
            pnlNetWorkProb.Enabled = False
            txtPhone.Enabled = False
            'ddlRepairMan.Visible = boolFixAsigned
            ''檢視
            If read_only = "1" Or IsPostBack Then
                ''**********************************
                ''流程控制
                ''**********************************
                But_exe.Visible = False
                But_prt.Visible = False
                backBtn.Visible = CheckFlowDrawVisible()
                tranBtn.Visible = False
                ''**********************************
                ''維修人員
                ''**********************************                
                'ddlRepairMan.Enabled = False
                ''**********************************
                ''請修地點
                ''**********************************
                ddlBuilding.Enabled = False
                ddlLevel.Enabled = False
                txtRoom.Enabled = False
                ''**********************************
                ''維修種類
                ''**********************************
                ddlRepairMainKind.Enabled = False
                ddlProblemKind.Enabled = False
                If ddlRepairMainKind.SelectedItem.Text = "網路故障" Then pnlNetWorkProb.Visible = True
                ''**********************************
                ''請修事項
                ''**********************************
                txtDESCRIPTION.Enabled = False
                ''**********************************
                ''附件上傳
                ''**********************************
                pnlApplyUploadFile.Visible = False
                pnlViewUploadFile.Visible = True
                ''**********************************
                ''完修資料-叫修時間
                ''**********************************
                txtCallDate.Enabled = False
                ddlCallTimeHour.Enabled = False
                ddlCallTimeMin.Enabled = False
                ''**********************************
                ''完修資料-到修時間
                ''**********************************
                txtArriveDate.Enabled = False
                ddlArriveTimeHour.Enabled = False
                ddlArriveTimeMin.Enabled = False
                ''**********************************
                ''完修資料-完修時間
                ''**********************************
                txtFinalDate.Enabled = False
                ddlFinalTimeHour.Enabled = False
                ddlFinalTimeMin.Enabled = False
                ''**********************************
                ''完修附件上傳
                ''**********************************
                pnlFinishUploadFile.Visible = False
                pnlViewFinishUploadFile.Visible = True
                ''**********************************
                ''維修紀錄
                ''**********************************
                txtFIXRECORD.Enabled = False
                ''**********************************
                ''重分按鈕
                ''**********************************
                btnReAppointment.Visible = False

                ''簽核
            ElseIf read_only = "2" Then
                ''**********************************
                ''請修地點
                ''**********************************
                ddlBuilding.Enabled = False
                ddlLevel.Enabled = False
                txtRoom.Enabled = False
                ''**********************************
                ''維修種類
                ''**********************************
                ddlRepairMainKind.Enabled = False
                ddlProblemKind.Enabled = False
                If ddlRepairMainKind.SelectedItem.Text = "網路故障" Then pnlNetWorkProb.Visible = True
                ''**********************************
                ''請修事項
                ''**********************************
                txtDESCRIPTION.Enabled = False
                ''**********************************
                ''附件上傳
                ''**********************************
                pnlApplyUploadFile.Visible = False
                pnlViewUploadFile.Visible = True
                ''**********************************
                ''維修人員
                ''**********************************                
                'ddlRepairMan.Enabled = flowAdmin = 2

                ''**********************************
                ''完修資料-叫修時間
                ''**********************************
                txtCallDate.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ddlCallTimeHour.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ddlCallTimeMin.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ''**********************************
                ''完修資料-到修時間
                ''**********************************
                txtArriveDate.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ddlArriveTimeHour.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ddlArriveTimeMin.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ''**********************************
                ''完修資料-完修時間
                ''**********************************
                txtFinalDate.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ddlFinalTimeHour.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ddlFinalTimeMin.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ''**********************************
                ''完修附件上傳
                ''**********************************
                pnlFinishUploadFile.Visible = (flowAdmin = 3 And strCaseStatus = "F")
                pnlViewFinishUploadFile.Visible = flowAdmin <> 3
                ''**********************************
                ''維修紀錄
                ''**********************************
                txtFIXRECORD.Enabled = (flowAdmin = 3 And strCaseStatus = "F")
                ''**********************************
                ''流程退回上一步
                ''**********************************
                btnReturn.Visible = CheckFlowReturnVisible()
                ''**********************************
                ''流程退回資訊報修管制單位
                ''**********************************
                btnReAppointment.Visible = CheckFlowReAppointmenVisible()
            End If
        End Sub

        Protected Function CheckFlowReAppointmenVisible() As Boolean
            Dim boolReturn As Boolean = False
            Select Case flowAdmin
                Case "1"
                Case "2"
                Case "3"
                    Dim LastFlow As New flowctl(eformid, eformsn, user_id, "?")
                    Dim PreFlow As flowctl = LastFlow.PreStep()
                    If PreFlow.group_name = "資訊報修管制單位" Then
                        boolReturn = True
                    End If
                Case Else
                    Dim LastFlow As New flowctl(eformid, eformsn, user_id, "?")
                    Dim PreFlow As flowctl = LastFlow.PreStep()
                    If PreFlow.PreStep().group_name = "資訊報修管制單位" Then
                        boolReturn = True
                    End If
            End Select

            CheckFlowReAppointmenVisible = boolReturn
        End Function

        Protected Function CheckFlowDrawVisible() As Boolean
            Dim boolReturn As Boolean = False
            Select Case flowAdmin
                Case "1"
                Case "2"
                Case "3"
                    ' ''取得上一個流程
                    'Dim LastFlow As New flowctl(eformid, eformsn, user_id, "?")
                    'Dim PreFlow As flowctl = LastFlow.PreStep()
                    'If PreFlow.group_name <> "上一級主管" Then
                    '    boolReturn = True
                    'End If
                Case Else
                    Dim LastFlowCtl As New flowctl(eformid, eformsn, "?")
                    Dim PreFlowCtl As flowctl = LastFlowCtl.PreStep()
                    If LastFlowCtl.group_name = "單位資訊官" And PreFlowCtl.group_name = "申請人" And PreFlowCtl.empuid = user_id Then
                        boolReturn = True
                    End If
            End Select
            CheckFlowDrawVisible = boolReturn
        End Function

        Protected Function CheckFlowReturnVisible() As Boolean
            Dim boolReturn As Boolean = False
            Select Case flowAdmin
                Case "1"
                Case "2"
                Case "3"
                    ' ''取得上一個流程
                    'Dim LastFlow As New flowctl(eformid, eformsn, user_id, "?")
                    'Dim PreFlow As flowctl = LastFlow.PreStep()
                    'If PreFlow.group_name <> "上一級主管" Then
                    '    boolReturn = True
                    'End If
                Case Else
                    Dim LastFlowCtl As New flowctl(eformid, eformsn, user_id, "?")
                    Dim PreFlowCtl As flowctl = LastFlowCtl.PreStep()
                    If PreFlowCtl.group_name = "資訊維修單位" And PreFlowCtl.PreStep().group_name = "上一級主管" Then
                        boolReturn = True
                    ElseIf LastFlowCtl.group_name = "資訊維修單位" And PreFlowCtl.group_name = "資訊報修管制單位" Then
                        boolReturn = False
                    End If
            End Select

            CheckFlowReturnVisible = boolReturn
        End Function

        Protected Sub LoadBasic()
            Dim tool As New C_Public
            lblApplyTime.Text = tool.GetChineseDate("2")
            hdApplyTime.Value = Now.ToString("yyyy/MM/dd HH:mm:dd")

        End Sub

        Protected Function Show(ByVal sReadonly As String, ByVal sFlowadmin As String, ByVal sType As String) As String
            ''sReadOnly="":表單申請  ="1":檢視  ="2":簽核
            Dim sReturn As String = "" ''class="hide"
            Select Case sType
                Case "Comment" ''批核意見
                    If sReadonly <> "2" Then
                        sReturn = " class=""hide"""
                    End If
                Case "CallTime", "ArriveTime", "FixRecord", "FixDate", "FixUploadFinish"
                    If sReadonly = "" Then
                        sReturn = " class=""hide"""
                    ElseIf sReadonly = "1" Then
                        sReturn = ""
                        Select Case sType
                            Case "CallTime"
                                sReturn = CType(IIf(ViewState("P11IsCallTime") = True, "", " class=""hide"""), String)
                            Case "ArriveTime"
                                sReturn = CType(IIf(ViewState("P11IsArriveTime") = True, "", " class=""hide"""), String)
                            Case "FixRecord"
                                sReturn = CType(IIf(ViewState("P11IsFIXRECORD") = True, "", " class=""hide"""), String)
                            Case "FixDate"
                                sReturn = CType(IIf(ViewState("P11IsFinalTime") = True, "", " class=""hide"""), String)
                            Case "FixUploadFinish"
                                sReturn = CType(IIf(ViewState("P11IsUPLOADFINISHNAME") = True, "", " class=""hide"""), String)
                        End Select
                    Else
                        If strCaseStatus <> "F" Then
                            sReturn = " class=""hide"""
                            If (Session("Role") = "1") Then
                                sReturn = ""
                                Select Case sType
                                    Case "CallTime"
                                        sReturn = CType(IIf(ViewState("P11IsCallTime") = True, "", " class=""hide"""), String)
                                    Case "ArriveTime"
                                        sReturn = CType(IIf(ViewState("P11IsArriveTime") = True, "", " class=""hide"""), String)
                                    Case "FixRecord"
                                        sReturn = CType(IIf(ViewState("P11IsFIXRECORD") = True, "", " class=""hide"""), String)
                                    Case "FixDate"
                                        sReturn = CType(IIf(ViewState("P11IsFinalTime") = True, "", " class=""hide"""), String)
                                    Case "FixUploadFinish"
                                        sReturn = CType(IIf(ViewState("P11IsUPLOADFINISHNAME") = True, "", " class=""hide"""), String)
                                End Select
                            End If
                        End If
                    End If
                Case "FixMan"
                    If sReadonly = "" Then
                        sReturn = " class=""hide"""
                    ElseIf sReadonly = "1" Then
                        sReturn = ""
                    End If
                    If Not boolFixAsigned And sFlowadmin <> "2" Then
                        sReturn = " class=""hide"""
                    End If
            End Select

            Show = sReturn
        End Function

        Protected Function SelectWholeTreeORG_UID(ByVal sUser_id As String) As String
            Dim sReturn As String = ""
            Dim tool As New C_Public

            sReturn = tool.GetWholeOrgIDs(sUser_id, ",", "'")
            SelectWholeTreeORG_UID = sReturn
        End Function

        Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

            Dim tool As New C_Public
            Dim strSql As String
            Dim n_table As DataTable
            Dim p As Integer
            Dim K As Integer
            Dim DC As SQLDBControl
            Dim DR As SqlDataReader

            LoadBasic()

            do_sql.G_user_id = Session("user_id").ToString()
            eformsn = Request("eformsn") '"VJ" 'd_pub.randstr(2) '
            eformid = Request("eformid")

            do_sql.G_errmsg = ""
            do_sql.G_user_id = Session("user_id").ToString() '"tempu180"
            'user_id = Session("user_id")
            org_uid = Session("ORG_UID").ToString()
            eformsn = Request("eformsn") '"VJ" 'd_pub.randstr(2) '
            eformid = Request("eformid")
            read_only = Request("read_only") '"":申請 1:檢視 2:簽核

            '取得登入者帳號
            If Page.User.Identity.Name.ToString.IndexOf("\", StringComparison.Ordinal) > 0 Then

                Dim LoginAll As String = Page.User.Identity.Name.ToString

                Dim LoginID() As String = Split(LoginAll, "\")

                user_id = LoginID(1)
            Else
                user_id = Page.User.Identity.Name.ToString
            End If

            If tool.CheckStepGroupEMPByName("單位資訊官", user_id, tool.GetObjectTypeFromStep("單位資訊官")) Then
                flowAdmin = "1"
                ''資訊官簽核時改變按鈕顯示名稱
                'If read_only = "2" Then But_exe.Text = "結案"
            ElseIf tool.CheckStepGroupEMPByName("資訊報修管制單位", user_id, tool.GetObjectTypeFromStep("資訊報修管制單位")) Then
                If tool.CheckGroupSigner("資訊報修管制單位", eformsn) Then
                    flowAdmin = "2"
                Else
                    flowAdmin = "0"
                End If
            ElseIf tool.CheckStepGroupEMPByName("資訊維修單位", user_id, tool.GetObjectTypeFromStep("資訊維修單位")) Then
                flowAdmin = "3"
            Else
                flowAdmin = "0"
            End If

            If IsPostBack Then Exit Sub 'PostBack才執行 

            SqlDataSource4.SelectCommand = "SELECT DISTINCT KIND_NUM,KIND_NAME FROM [SYSKIND] WHERE State_enabled=1 AND ([Kind_Num] IN ('12','13')) "
            ddlBuilding.DataBind()

            SqlDataSource2.SelectCommand = "SELECT DISTINCT KIND_NUM,KIND_NAME FROM [SYSKIND] WHERE State_enabled=1 AND ([Kind_Num] IN ('7','8','9')) "
            ddlRepairMainKind.DataBind()


            'session被清空回首頁
            If user_id = "" And org_uid = "" And eformsn = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                If do_sql.select_urname(do_sql.G_user_id) = False Then
                    Lab_PWUNIT.Text = ""
                    Exit Sub
                End If
                If do_sql.G_usr_table.Rows.Count > 0 Then
                    Lab_PWUNIT.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
                    Lab_PWNAME.Text = do_sql.G_usr_table.Rows(0).Item("emp_chinese_name").ToString.Trim
                    Lab_PWTITLE.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim

                End If

                ''讀入審核詳細資料
                If read_only = "1" Or read_only = "2" Then
                    If LoadData() = False Then
                        Exit Sub
                    End If
                    Exit Sub
                End If

                If org_uid <> "" Then
                    'stmt = "select * from Employee where ORG_UID='" + str_ORG_UID + "' order by emp_chinese_name"
                    strSql = "select * from Employee where  ORG_UID IN (" & tool.getchildorg(org_uid) & ") order by emp_chinese_name"
                    If do_sql.db_sql(strSql, do_sql.G_conn_string) = False Then
                        Exit Sub
                    End If
                    If do_sql.G_table.Rows.Count > 0 Then
                        n_table = do_sql.G_table
                        p = 0
                        K = 0
                        DrDown_PANAME.Items.Clear()

                        For Each datarrow In n_table.Rows
                            DrDown_PANAME.Items.Add(Trim(datarrow("emp_chinese_name").ToString))
                            DrDown_PANAME.Items(p).Value = Trim(datarrow("employee_id").ToString)
                            If UCase(do_sql.G_user_id) = UCase(Trim(datarrow("employee_id").ToString)) Then
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

                '找出表單審核者
                If read_only = "2" Then
                    DC = New SQLDBControl
                    strSql = "SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' and hddate is null"
                    DR = DC.CreateReader(strSql)
                    If DR.Read() Then
                        AgentEmpuid = DR("empuid").ToString()
                    End If
                    DC.Dispose()

                End If

            End If
            If Not IsPostBack Then
                ddlPAUNIT.Items.Clear()
                '判斷登入者權限
                If Session("Role") = "1" Then
                    SqlDataSource7.SelectCommand = "SELECT RTRIM(LTRIM(ORG_UID)) AS ORG_UID, RTRIM(LTRIM(ORG_NAME)) AS ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                ElseIf flowAdmin = "1" Then
                    SqlDataSource7.SelectCommand = "SELECT RTRIM(LTRIM(ORG_UID)) AS ORG_UID, RTRIM(LTRIM(ORG_NAME)) AS ORG_NAME FROM [ADMINGROUP] WHERE ORG_UID IN (" + SelectWholeTreeORG_UID(user_id) + ") ORDER BY [ORG_NAME]"
                Else
                    SqlDataSource7.SelectCommand = "SELECT RTRIM(LTRIM(ORG_UID)) AS ORG_UID, RTRIM(LTRIM(ORG_NAME)) AS ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                    ddlPAUNIT.Enabled = False
                End If
                ddlPAUNIT.DataBind()
            End If
        End Sub

        Protected Sub DrDown_PANAME_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DrDown_PANAME.SelectedIndexChanged
            If do_sql.select_urname(DrDown_PANAME.Items(DrDown_PANAME.SelectedIndex).Value) = False Then
                DrDown_PANAME.Text = ""
                Exit Sub
            End If

            If do_sql.G_usr_table.Rows.Count > 0 Then
                'ddlPAUNIT.SelectedItem.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
                Lab_PATITLE.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
            End If
        End Sub

        Protected Sub Page_PreRenderComplete(sender As Object, e As EventArgs) Handles Me.PreRenderComplete
            Dim conn As New C_SQLFUN
            Dim connstr As String = conn.G_conn_string

            strCaseStatus = CType(IIf(IsNothing(ViewState("P11CaseStatus")), "", ViewState("P11CaseStatus")), String)

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '單純讀取表單不可送件
            If read_only = "1" Then

                But_exe.Visible = False                     '送件
                backBtn.Visible = CheckFlowDrawVisible()    '駁回
                tranBtn.Visible = False                     '呈轉
                btnReturn.Visible = False                   '退回
                btnReAppointment.Visible = False            '重分

                '讀取表單可送件
            ElseIf read_only = "2" Then

                Dim PerCount As Integer = 0

                '判斷批核者是否為上一關送件者的副主管,假如批核者為副主管則不可送件只可呈轉
                Dim ParentFlag As String
                Dim ParentVal = user_id & "," & eformsn

                Dim FC As New C_FlowSend.C_FlowSend
                ParentFlag = FC.F_NextChief(ParentVal, connstr).ToString()

                If ParentFlag = "1" Then
                    'But_exe.Visible = False

                    '上一級主管多少人
                    db.Open()
                    Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                    Dim PerRdv = PerCountCom.ExecuteReader()
                    If PerRdv.Read() Then
                        PerCount = CType(PerRdv("PerCount"), Integer)
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
                Else
                    Dim LastFlowCtl As New flowctl(eformid, eformsn, user_id, "?") ''現在流程
                    Dim PreFlowCtl As flowctl = LastFlowCtl.PreStep() ''上一關流程

                    ''上一個流程是派工單位
                    If LastFlowCtl.group_name = "資訊報修管制單位" Or LastFlowCtl.group_name = "資訊維修單位" Or PreFlowCtl.group_name = "資訊維修單位" Then
                        backBtn.Visible = False
                    End If
                End If

            Else
                backBtn.Visible = False
                tranBtn.Visible = False
            End If

            ''各簽核按鈕名稱設定
            Select Case flowAdmin
                Case "1"
                    'But_exe.Text = CType(IIf(read_only = "", "送件", "呈轉"), String)             '送件                    
                Case "2"
                    But_exe.Text = CType(IIf(read_only = "", "送件", "分派"), String)             '送件                    
                Case "3"
                    'But_exe.Text = CType(IIf(read_only = "", "送件", "呈轉"), String)             '送件                    
                    btnReAppointment.Text = "重派"    '重分
                Case Else
                    Dim tool As New C_Public
                    'But_exe.Text = CType(IIf(read_only = "", "送件", "呈轉"), String)
                    backBtn.Text = CType(IIf(read_only = "1", "撤銷", backBtn.Text), String)
                    If tool.IsFixmanSupervisor(user_id) Then
                        But_exe.Text = CType(IIf(strCaseStatus = "F", "結案", "核准"), String)             '送件                        
                        'btnReturn.Text = "駁回"           '退回
                        btnReAppointment.Text = "重派"    '重分
                    End If
            End Select
        End Sub

        Protected Sub But_exe_Click(sender As Object, e As EventArgs) Handles But_exe.Click

            Dim FC As New CFlowSend
            Dim DC As SQLDBControl
            Dim DR As SqlDataReader
            Dim strSql As String
            Dim tool As New C_Public

            Dim SendVal As String
            AppUID = DrDown_PANAME.SelectedValue
            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                'SendVal = eformid & "," & user_id & "," & eformsn & "," & "1" & ","
                SendVal = eformid & "," & AppUID & "," & eformsn & "," & "1" & ","
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1" & ","
            End If

            'Dim NextPer As Integer = 0


            '關卡為上一級主管有多少人
            'DC = New SQLDBControl
            'NextPer = FC.F_NextStep(SendVal, DC.ConnStr)
            'DC.Dispose()

            '判斷同單位是否有單位行政官
            '指派資訊官
            Dim OAPer As Integer = 0
            Dim OAPerORG As String = ""
            OAPer = Split(tool.GetUnitIDsByStepName(user_id, "單位資訊官", ",", "", 3), ",").Length

            Dim Chknextstep As Integer

            '判斷表單關卡
            DC = New SQLDBControl
            DR = DC.CreateReader("SELECT nextstep FROM flowctl WHERE eformsn = '" & eformsn & "' and empuid = '" & user_id & "' and hddate is null")
            If DR.Read() Then
                Chknextstep = CType(DR("nextstep"), Integer)
            End If
            DC.Dispose()

            Dim strgroup_id As String = ""

            'If read_only <> "" Then

            '    '判斷上一級主管
            '    DC = New SQLDBControl
            '    DR = DC.CreateReader("SELECT group_id FROM flow WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ")
            '    If DR.Read() Then
            '        strgroup_id = DR("group_id").ToString()
            '    End If
            '    DC.Dispose()

            'End If

            If OAPer = 0 And read_only = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('無上一級主管');")
                Response.Write(" </script>")

                Exit Sub

            End If

            do_sql.G_errmsg = ""

            If read_only = "" Then

                If CheckApplyForm() = False Then
                    Exit Sub
                End If
                ''上傳檔案
                Dim Path As String = Server.MapPath("~/M_Source/11/Upload/")    '定義上傳後存檔路徑

                'Dim FileOk As Boolean = False   '宣告一個FileOk用來判別是否上傳成功，預設為False                
                'Dim i As Integer
                If FileUpload1.HasFile Then  '透過HasFile判斷有檔案上傳                    
                    Try
                        If FileUpload1.PostedFile.ContentLength > (2 * 1024 * 1024) Then
                            Response.Write(" <script language='javascript'>")
                            Response.Write(" alert('檔案大小超過2MB，請重新選擇');")
                            Response.Write(" </script>")
                            Exit Sub
                        End If
                        FileUpload1.PostedFile.SaveAs(Path + eformsn + "-" + FileUpload1.FileName)    '將上傳的檔案儲存
                        ''Label1.Text = "Upload Success!!"     '傳回成功
                    Catch ex As Exception
                        ''Label1.Text = "Upload False!! <br>" + ex.Message
                        Exit Sub
                    End Try
                End If

                strSql = "INSERT INTO P_11 (EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
                strSql += "PWNAME,PWIDNO,PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
                strSql += "PANAME,PATITLE,PAIDNO," '申請人姓名,申請人級職,申請人身份證字號
                strSql += "APPTIME,PHONE," '申請時間,連絡電話
                strSql += "BROKENTYPE,DESCRIPTION,LOCATION," '維修種類,維修事項,維修地點
                strSql += "IPADDRESS,MACADDRESS,PLUGPORT,UPLOADNAME," 'IP位址,MAC位址,插座編號,上傳檔案名稱
                strSql += "FIXIDNO"
                strSql += ") VALUES ("

                strSql += "'" + eformsn + "'" '表單序號
                strSql += ",'" + Lab_PWUNIT.Text + "'" '填表人單位
                strSql += ",'" + Lab_PWTITLE.Text + "'" '填表人級職
                strSql += ",N'" + Lab_PWNAME.Text + "'" '填表人姓名
                strSql += ",'" + do_sql.G_user_id + "'" '填表人身份證字號
                strSql += ",'" + ddlPAUNIT.SelectedItem.Text + "'" '申請人單位
                strSql += ",N'" + DrDown_PANAME.SelectedItem.Text + "'" '申請人姓名
                strSql += ",'" + Lab_PATITLE.Text + "'" '申請人級職
                strSql += ",'" + DrDown_PANAME.SelectedItem.Value + "'" '申請人身份證字號
                strSql += ",'" + hdApplyTime.Value + "'" '申請時間
                strSql += ",'" + txtPhone.Text + "'" '連絡電話
                strSql += ",'" + ddlRepairMainKind.SelectedValue + ddlProblemKind.SelectedValue + "'" '維修種類
                strSql += ",N'" + txtDESCRIPTION.Text + "'" '維修事項
                strSql += ",'" + ddlBuilding.SelectedValue + ddlLevel.SelectedValue + txtRoom.Text + "'" '維修地點
                strSql += ",'" + txtIPAddr.Text + "'" 'IP位址
                strSql += ",'" + txtMACAddr.Text + "'" 'MAC位址
                strSql += ",'" + txtPlugNo.Text + "'" '插座編號
                strSql += ",'" + FileUpload1.FileName + "'" '上傳檔案名稱
                strSql += ",''" '維修人員
                strSql += ")"
                If do_sql.db_exec(strSql, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
            Else
                'If txtFinalDate.Text.Length = 0 Then
                '    MessageBox.Show("請填寫完工時間及更改處理狀態")
                '    Exit Sub
                'End If
                ''上傳檔案
                If flowAdmin = 3 Then
                    If strCaseStatus = "F" Or ViewState("P11CaseStatus") = "F" Then
                        ''判斷必填時間欄位
                        If (CheckSignForm() = False) Then
                            MessageBox.Show("日期有誤，請確認必填日期或先後順序")
                            Exit Sub
                        End If

                        Dim Path As String = Server.MapPath("~/M_Source/11/Upload/")    '定義上傳後存檔路徑
                        If FileUpload2.HasFile Then  '透過HasFile判斷有檔案上傳                    
                            Try
                                If FileUpload2.PostedFile.ContentLength > (2 * 1024 * 1024) Then
                                    Response.Write(" <script language='javascript'>")
                                    Response.Write(" alert('檔案大小超過2MB，請重新選擇');")
                                    Response.Write(" </script>")
                                    Exit Sub
                                End If
                                FileUpload2.PostedFile.SaveAs(Path + eformsn + "-Final-" + FileUpload2.FileName)    '將上傳的檔案儲存
                                ''Label1.Text = "Upload Success!!"     '傳回成功
                            Catch ex As Exception
                                ''Label1.Text = "Upload False!! <br>" + ex.Message
                                Exit Sub
                            End Try
                        End If


                        strSql = "UPDATE P_11 SET "
                        'strSql += "FIXTYPE='" & rdo_nExternal.SelectedItem.Value & "'"
                        strSql += "CALLTIME='" + txtCallDate.Text + " " + ddlCallTimeHour.SelectedValue + ":" + ddlCallTimeMin.SelectedValue + "'"
                        strSql += ",ARRIVETIME='" + txtArriveDate.Text + " " + ddlArriveTimeHour.SelectedValue + ":" + ddlArriveTimeMin.SelectedValue + "'"
                        strSql += ",FINALDATE='" + txtFinalDate.Text + " " + ddlFinalTimeHour.SelectedValue + ":" + ddlFinalTimeMin.SelectedValue + "'"
                        strSql += ",FIXRECORD='" + txtFIXRECORD.Text + "'"
                        strSql += ",UPLOADFINISHNAME='" + FileUpload2.FileName + "'"
                        strSql += " WHERE EFORMSN='" & eformsn & "'"
                        If do_sql.db_exec(strSql, do_sql.G_conn_string) = False Then
                            Exit Sub
                        End If


                    End If
                Else
                    'Dim flow As New flowctl(eformid, eformsn, user_id, "?")
                    'Dim Preflow As flowctl = flow.PreStep()
                    'If Preflow.group_name = "資訊維修單位" And Preflow.PreStep().group_name = "資訊報修管制單位" Then
                    '    strSql = "UPDATE P_11 SET "
                    '    'strSql += "FIXTYPE='" & rdo_nExternal.SelectedItem.Value & "'"
                    '    strSql += "PENDFLAG='F'"
                    '    strSql += " WHERE EFORMSN='" & eformsn & "'"
                    '    If do_sql.db_exec(strSql, do_sql.G_conn_string) = False Then
                    '        Exit Sub
                    '    End If
                    'End If
                End If
                '增加批核意見
                insComment(txtcomment.Text, eformsn, user_id)
            End If

            If read_only = "2" Then

                Dim strAgentName As String = ""

                '判斷是否為代理人批核
                If UCase(user_id) = UCase(AgentEmpuid) Then

                    '增加批核意見
                    insComment(txtcomment.Text, eformsn, user_id)

                Else

                    Dim strComment As String

                    '找尋批核者姓名
                    DC = New SQLDBControl
                    DR = DC.CreateReader("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'")
                    If DR.Read() Then
                        strAgentName = DR("emp_chinese_name").ToString()
                    End If
                    DC.Dispose()

                    strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)"

                    '增加批核意見
                    insComment(strComment, eformsn, AgentEmpuid)

                End If
                ''管制單位簽核
                If flowAdmin = 2 Then
                    'If ddlRepairMan.SelectedValue = "" Then
                    '    Response.Write(" <script language='javascript'>")
                    '    Response.Write(" alert('請選擇維修人員');")
                    '    Response.Write(" </script>")
                    '    Exit Sub
                    'Else
                    '    strSql = "UPDATE P_11 SET "
                    '    strSql += "FIXIDNO='" & ddlRepairMan.SelectedValue & "'"
                    '    strSql += " WHERE EFORMSN='" & eformsn & "'"
                    '    If do_sql.db_exec(strSql, do_sql.G_conn_string) = False Then
                    '        Exit Sub
                    '    End If
                    'End If
                End If
            End If

            Dim Val_P As String
            'Val_P = ""

            '當新表單判斷同單位是否有單位行政官
            '同單位沒有則導入下一頁
            '下一頁找出全部同一級單位的行政官
            '新表單
            If read_only = "" Then



                'Dim strsame As String '= "SELECT COUNT(*) AS OAPER FROM SYSTEMOBJUSE A LEFT JOIN SYSTEMOBJ B ON A.OBJECT_UID=B.OBJECT_UID LEFT JOIN EMPLOYEE C ON A.EMPLOYEE_ID =C.EMPLOYEE_ID WHERE B.OBJECT_UID='" & tool.GetInformationIDByName("單位資訊官") & "' AND C.ORG_UID IN (" & tool.GetAllStepOrgIDs(user_id, "'") & ")"
                'OAPer = Split(tool.GetUnitIDsByStepName(user_id, "單位資訊官", ",", "", 3), ",").Length
                'DC = New SQLDBControl
                'DR = DC.CreateReader(strsame)
                'If DR.Read() Then
                '    OAPer = CType(DR("OAPer"), Integer)
                'End If
                'DC.Dispose()

                If OAPer = 0 Then '沒有指定資訊官
                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('無任何單位資訊官');")
                    Response.Write(" </script>")
                ElseIf OAPer = 1 Then '只有一位資訊官
                    DC = New SQLDBControl
                    strSql = "SELECT * FROM SYSTEMOBJUSE A LEFT JOIN SYSTEMOBJ B ON A.OBJECT_UID=B.OBJECT_UID LEFT JOIN EMPLOYEE C ON A.EMPLOYEE_ID =C.EMPLOYEE_ID WHERE B.OBJECT_UID='" & tool.GetInformationIDByName("單位資訊官") & "' AND C.ORG_UID IN (" & tool.GetWholeOrgIDs(user_id, ",", "'") & ")"
                    DR = DC.CreateReader(strSql)
                    If DR.Read() Then
                        OAPerORG = DR("employee_id").ToString()
                    End If
                    DC.Dispose()

                    Dim SendValOA = eformid & "," & do_sql.G_user_id & "," & eformsn & "," & "1" & "," & OAPerORG

                    '表單審核
                    Val_P = FC.F_Send(SendValOA, do_sql.G_conn_string).ToString()
                    Dim PageUp As String = ""

                    If read_only = "" Then
                        PageUp = "New"
                    End If

                    Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                    do_sql.G_errmsg = "存檔成功"
                ElseIf OAPer > 1 Then '多位資訊官
                    'MessageBox.Show("單位資訊官設定過多，請聯絡管理員")
                    'Exit Sub

                    Server.Transfer("../00/MOA00015.aspx?SendVal=" + eformid & "," & do_sql.G_user_id & "," & eformsn & "," & "1" & "," & OAPerORG)
                End If
                '同單位無資訊官找出一級單位全部資訊官
                '同單位有資訊官則送出
                '同單位有兩位以上則導入下一頁
                'If OAPer = 0 Then

                '    Dim CP As New C_Public
                '    strOrgTop = CP.getUporg(org_uid, 1)

                '    DC = New SQLDBControl
                '    DR = DC.CreateReader("SELECT count(EMPLOYEE.employee_id) as OAPerAll FROM SYSTEMOBJUSE INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID WHERE (SYSTEMOBJUSE.object_uid = '" + tool.GetInformationIDByName("單位資訊官") + "') AND (EMPLOYEE.ORG_UID IN (" & CP.getchildorg(strOrgTop) & "))")

                '    If DR.Read() Then
                '        OAPerAll = DR.Item("OAPerAll")
                '    End If
                '    DC.Dispose()

                '    If OAPerAll = 0 Then

                '        Response.Write(" <script language='javascript'>")
                '        Response.Write(" alert('無任何單位資訊官');")
                '        Response.Write(" </script>")

                '        Exit Sub

                '    ElseIf OAPerAll = 1 Then

                '        '找出一級單位唯一的單位資訊官
                '        DC = New SQLDBControl
                '        DR = DC.CreateReader("SELECT DISTINCT EMPLOYEE.employee_id FROM SYSTEMOBJUSE INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID WHERE (SYSTEMOBJUSE.object_uid = '" + tool.GetInformationIDByName("單位資訊官") + "') AND (EMPLOYEE.ORG_UID IN (" & CP.getchildorg(strOrgTop) & "))")
                '        If DR.Read() Then
                '            OAPerORG = DR("employee_id")
                '        End If
                '        DC.Dispose()

                '        Dim SendValOA = eformid & "," & do_sql.G_user_id & "," & eformsn & "," & "1" & "," & OAPerORG

                '        '表單審核
                '        Val_P = FC.F_Send(SendValOA, do_sql.G_conn_string)
                '        Dim PageUp As String = ""

                '        If read_only = "" Then
                '            PageUp = "New"
                '        End If

                '        Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                '        do_sql.G_errmsg = "存檔成功"

                '    ElseIf OAPerAll > 0 Then
                '        If GetSuperiors(user_id).Rows.Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                '            Server.Transfer("../00/MOA00014.aspx?eformsn=" + eformsn)
                '        End If
                '        'Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=3&strOrgTop=" & strOrgTop)
                '    End If

                'ElseIf OAPer = 1 Then
                '    '表單審核
                '    Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)
                '    Dim PageUp As String = ""

                '    If read_only = "" Then
                '        PageUp = "New"
                '    End If

                '    Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                '    do_sql.G_errmsg = "存檔成功"
                'ElseIf OAPer > 1 Then
                '    Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=2")
                'End If


                '簽核表單
            Else
                '關卡為上一級主管有多少人
                Dim NextPer As Int32 = FC.F_NextStep(SendVal, do_sql.G_conn_string)

                '判斷下一關為上一級主管時人數是否超過一人
                Dim flowctl As New flowctl(eformid, eformsn, user_id, "?")
                Dim flow As New flow(flowctl.stepsid)
                Dim NextFlow As New flow(flow.nextstep)
                If NextPer > 1 And flow.NextFlow().group_id = "Z860" Then
                    Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=1&AppUID=" & DrDown_PANAME.SelectedValue)
                Else                    
                    Dim Preflow As flowctl = flowctl.PreStep()
                    If Preflow.group_name = "資訊維修單位" And Preflow.PreStep().group_name = "資訊報修管制單位" Then
                        SendVal = eformid & "," & user_id & "," & eformsn & "," & "1" & "," & hdnRepairMan.Value '& ddlRepairMan.SelectedValue

                        strSql = "UPDATE P_11 SET "
                        'strSql += "FIXTYPE='" & rdo_nExternal.SelectedItem.Value & "'"
                        strSql += "PENDFLAG='F'"
                        strSql += " WHERE EFORMSN='" & eformsn & "'"
                        If do_sql.db_exec(strSql, do_sql.G_conn_string) = False Then
                            Exit Sub
                        End If
                    End If

                    Dim da_group As DataTable
                    da_group = CType(FC.F_GroupEmp(NextFlow.group_id, do_sql.G_conn_string), DataTable)
                    If da_group.Rows.Count > 1 AndAlso hdnRepairMan.Value = "" Then '主資料未有維修人資料表示流程在管制單位上
                        Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=4&group_id=" & NextFlow.group_id & "&AppUID=" & DrDown_PANAME.SelectedValue)
                    End If
                End If

                ''確定流程在維修人員接單確定的上一級主管中
                If (flowctl.group_name = "資訊報修管制單位") Or (flowctl.group_name = "上一級主管" And flowctl.PreStep().group_name = "資訊維修單位") Then
                    If flowctl.nextstep <> "-1" Then ''確定流程不是完修解管
                        SendVal = eformid & "," & user_id & "," & eformsn & "," & "1" & "," & hdnRepairMan.Value '& ddlRepairMan.SelectedValue
                    End If
                End If

                '表單審核
                Val_P = FC.F_Send(SendVal, do_sql.G_conn_string).ToString()
                Dim PageUp As String = ""

                If read_only = "" Then
                    PageUp = "New"
                End If

                Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                do_sql.G_errmsg = "存檔成功"

                End If

        End Sub

        Protected Sub backBtn_Click(sender As Object, e As EventArgs) Handles backBtn.Click

            Try
                '表單駁回
                Dim FC As New CFlowSend
                Dim Val_P As String = ""
                Dim SendVal As String

                If read_only = "2" Then
                    '增加批核意見
                    insComment(txtcomment.Text, eformsn, user_id)

                    '判斷是否為代理人批核的表單
                    If AgentEmpuid = "" Then
                        SendVal = eformid & "," & user_id & "," & eformsn & "," & "1"
                    Else
                        SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1"
                    End If

                    Val_P = FC.F_Back(SendVal, do_sql.G_conn_string).ToString()

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('表單已駁回給申請人');")
                    '重新整理頁面
                    Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
                    Response.Write(" window.close();")
                    Response.Write(" </script>")

                ElseIf read_only = "1" Then
                    If flowAdmin = 0 Then
                        SendVal = eformid & "," & eformsn
                        Val_P = FC.F_DrawBack(SendVal, do_sql.G_conn_string).ToString()
                        If Val_P = "1" Then
                            MessageBox.Show("表單已批核不可撤銷")
                        Else
                            Dim DC As New SQLDBControl
                            Dim strSql As String = "UPDATE P_11 SET PENDFLAG='B' WHERE EFORMSN='" + eformsn+"'"

                            DC.ExecuteSQL(strSql)
                            DC.Dispose()
                            MessageBox.Show("表單撤銷完成")
                        End If
                    End If
                End If

            Catch ex As Exception

            End Try
        End Sub

        Protected Sub tranBtn_Click(sender As Object, e As EventArgs) Handles tranBtn.Click
            Try
                If GetSuperiors(user_id).Rows.Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                    Server.Transfer("../00/MOA00014.aspx?eformsn=" + eformsn)
                Else
                    If read_only = "2" Then
                        '增加批核意見
                        insComment(txtcomment.Text, eformsn, user_id)
                    End If

                    Dim Val_P As String = ""

                    '表單呈轉
                    Dim FC As New C_FlowSend.C_FlowSend

                    Dim SendVal As String

                    '判斷是否為代理人批核的表單
                    If AgentEmpuid = "" Then
                        SendVal = eformid & "," & user_id & "," & eformsn & "," & "1"
                    Else
                        SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1"
                    End If

                    Val_P = FC.F_Transfer(SendVal, do_sql.G_conn_string).ToString()

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
            Dim db As New SqlConnection(do_sql.G_conn_string)
            Dim ds As New DataSet()

            db.Open()
            Dim comm As New SqlCommand("select employee_id,emp_chinese_name from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id = @employee_id))", db)
            comm.Parameters.Add("@employee_id", SqlDbType.VarChar, 10).Value = employee_id.Trim()
            Dim da As New SqlDataAdapter(comm)
            da.Fill(ds)
            db.Close()
            Return ds.Tables(0)
        End Function

        'Protected Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        '    Try
        '        If (CheckSignForm() = False) Then
        '            MessageBox.Show("日期有誤，請確認日期先後順序")
        '            Exit Sub
        '        End If
        '        ''更新處理資料
        '        ''進度
        '        ''按下更新資料自動將狀態改為處理中
        '        rdbStatus.Items(1).Selected = True
        '        Dim sStatus As String = rdbStatus.Items(1).Value
        '        ''報修時間
        '        Dim sCallDateTime As String = ""
        '        If txtCALLDATE.Text.Length > 0 Then
        '            sCallDateTime = txtCALLDATE.Text + " " + ddlCALLTIMEHour.SelectedValue + ":" + ddlCALLTIMEMin.SelectedValue
        '        End If

        '        ''到修時間
        '        Dim sArrivedDateTime As String = ""
        '        If txtARRIVEDATE.Text.Length > 0 Then
        '            sArrivedDateTime = txtARRIVEDATE.Text + " " + ddlARRIVETIMEHour.SelectedValue + ":" + ddlARRIVETIMEMin.SelectedValue
        '        End If
        '        ''維修紀錄
        '        Dim sRepairRecord As String = ""
        '        If txtRecord.Text.Length > 0 Then
        '            sRepairRecord = txtRecord.Text
        '        End If

        '        Dim strSql As String


        '        If flowAdmin <> "0" Then
        '            Dim DC As New SQLDBControl

        '            strSql = "UPDATE P_11 SET "
        '            strSql += "FIXSTATUS='" & sStatus & "'"

        '            If sCallDateTime.Length > 0 Then strSql += ",CALLTIME = '" & sCallDateTime & "'"
        '            If sArrivedDateTime.Length > 0 Then strSql += ",ARRIVETIME='" & sArrivedDateTime & "'"
        '            If sRepairRecord.Length > 0 Then strSql += ",FIXRECORD='" & sRepairRecord & "'"

        '            strSql += " WHERE EFORMSN='" & eformsn & "'"
        '            DC.ExecuteSQL(strSql)
        '            DC.Dispose()
        '            MessageBox.Show("更新完成")
        '        End If
        '    Catch ex As Exception

        '    End Try

        'End Sub


        Protected Sub But_PHRASE_Click(sender As Object, e As EventArgs) Handles But_PHRASE.Click
            Dim conn As New C_SQLFUN
            Dim connstr As String = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            Dim Visitors As Integer
            'Dim str_SQL As String = ""

            '否有申請過請假表單
            db.Open()
            Dim PerCountCom As New SqlCommand(CType(("SELECT count(*) as Visitors from Phrase WHERE employee_id = '" & Session("user_id") & "'"), String), db)
            Dim PerRdv = PerCountCom.ExecuteReader()
            If PerRdv.Read() Then
                Visitors = CType(PerRdv("Visitors"), Integer)
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

        Protected Sub ddlBuilding_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlBuilding.SelectedIndexChanged
            ddlLevel.Items.Clear()
            ddlLevel.DataBind()
        End Sub

        Protected Sub ddlRepairMainKind_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlRepairMainKind.SelectedIndexChanged
            Dim ddl As DropDownList = CType(sender, DropDownList)
            If IsNothing(ddl) Then Exit Sub
            If CInt(ddl.SelectedValue) = 9 Then
                pnlNetWorkProb.Visible = True
            Else
                pnlNetWorkProb.Visible = False
            End If
        End Sub

        Protected Sub ddlRepairMan_DataBound(sender As Object, e As EventArgs) Handles ddlRepairMan.DataBound
            'Dim ddl As DropDownList = CType(sender, DropDownList)
            'If Not IsNothing(ddl) Then
            '    ddl.Items.Insert(0, New ListItem("請選擇", ""))
            'End If
        End Sub

        Protected Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
            Dim LastFlow As New flowctl(eformid, eformsn, user_id, "?") ''現在流程
            Dim PreFlow As flowctl = LastFlow.PreStep() ''上一關流程
            Dim Val_P As String = ""
            ''先新增批核意見
            insComment(txtcomment.Text, eformsn, LastFlow.empuid)

            Select Case flowAdmin
                Case "1"
                Case "2"
                Case "3"
                Case Else
                    ''上一個流程是派工單位
                    If PreFlow.group_name = "資訊維修單位" Then
                        Val_P = LastFlow.JunpBack()
                    End If
            End Select
            ''傳送至如影隨行畫面
            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & "")

        End Sub

        Protected Sub btnReAppointment_Click(sender As Object, e As EventArgs) Handles btnReAppointment.Click
            Dim LastFlow As New flowctl(eformid, eformsn, user_id, "?") ''現在流程
            Dim Val_P As String = ""
            ''先新增批核意見
            insComment(txtcomment.Text, eformsn, LastFlow.empuid)

            Select Case flowAdmin
                Case "1", "2"
                Case "3"
                    ''送件給資訊報修管制單位
                    Val_P = LastFlow.JumpByGroupName("資訊報修管制單位")
                Case Else
                    ''送件給資訊報修管制單位
                    Val_P = LastFlow.JumpByGroupName("資訊報修管制單位")
            End Select

            '退回分派後需清空維修人員
            Dim Dc As New SQLDBControl
            Dim strSql = "UPDATE P_11 SET FIXIDNO='' WHERE EFORMSN='" & eformsn & "'"
            Dc.ExecuteSQL(strSql)
            Dc.Dispose()
            ''傳送至如影隨行畫面
            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & "")

        End Sub

        Protected Sub DrDown_PANAME_DataBound(sender As Object, e As EventArgs) Handles DrDown_PANAME.DataBound
            'If DrDown_PANAME.Enabled = False Then Exit Sub
            If read_only = "" Then
                If ddlPAUNIT.SelectedValue = org_uid Then
                    For Each o As ListItem In DrDown_PANAME.Items
                        If o.Value = UCase(user_id) Then
                            DrDown_PANAME.SelectedValue = UCase(user_id)
                            Exit For
                        ElseIf o.Value = LCase(user_id) Then
                            DrDown_PANAME.SelectedValue = LCase(user_id)
                            Exit For
                        End If
                    Next
                Else
                    DrDown_PANAME.SelectedValue = DrDown_PANAME.Items(0).Value
                End If
            End If
        End Sub

        Protected Sub ddlPAUNIT_DataBound(sender As Object, e As EventArgs) Handles ddlPAUNIT.DataBound
            If read_only = "" Then
                CType(sender, DropDownList).SelectedValue = org_uid
            End If
        End Sub

        Protected Sub ddlPAUNIT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPAUNIT.SelectedIndexChanged
            '清空User重新讀取
            DrDown_PANAME.Items.Clear()

            If ddlPAUNIT.SelectedValue = "" Then
                SqlDataSource12.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE 1=2"
            Else
                If Session("Role") = "1" Or flowAdmin = "1" Then
                    SqlDataSource12.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & ddlPAUNIT.SelectedValue & "' ORDER BY emp_chinese_name"
                Else
                    SqlDataSource12.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & ddlPAUNIT.SelectedValue & "' AND employee_id='" & user_id & "' ORDER BY emp_chinese_name"
                    DrDown_PANAME.Enabled = False
                End If
            End If
            DrDown_PANAME.DataBind()
        End Sub

        Protected Sub GridView10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridView10.SelectedIndexChanged
            txtcomment.Text += GridView10.Rows(GridView10.SelectedRow.RowIndex).Cells(0).Text
            Div_grid10.Visible = False
        End Sub

        Protected Sub Btn_PHclose_Click(sender As Object, e As EventArgs) Handles Btn_PHclose.Click
            Div_grid10.Visible = False
        End Sub
    End Class
End Namespace