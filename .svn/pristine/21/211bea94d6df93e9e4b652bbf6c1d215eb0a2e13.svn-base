Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_09_MOA09001
    Inherits System.Web.UI.Page
    Dim org_uid As String
    Dim sql_function As New C_SQLFUN
    Dim connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '取得登入者帳號
            If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then
                Dim LoginAll As String = Page.User.Identity.Name.ToString
                Dim LoginID() As String = Split(LoginAll, "\")
                ViewState("userID") = LoginID(1)
            Else
                ViewState("userID") = Page.User.Identity.Name.ToString
            End If

            org_uid = Session("ORG_UID") '使用者單位ID
            ViewState("eFormID") = Request.QueryString("eformid") '表單種類ID
            ViewState("eFormSN") = Request.QueryString("eformsn") '表單ID
            '申請單頁面開啟模式 "":新增申請單 1:瀏覽表單 2:審核表單
            ViewState("read_only") = Request.QueryString("read_only")
            'session被清空回首頁
            If String.IsNullOrEmpty(ViewState("userID")) And String.IsNullOrEmpty(org_uid) And String.IsNullOrEmpty(ViewState("eFormSN")) Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.parent.location='../../index.aspx';")
                Response.Write(" </script>")
                Response.End()
            End If

            If String.IsNullOrEmpty(ViewState("read_only").ToString()) Then 'read_only="" 為新增表單
                SettingNewFormData(ViewState("userID").ToString()) '設定新增表單顯示資料
            ElseIf ViewState("read_only").ToString().Equals("1") Then 'read_only="1" 為瀏覽表單
                SettingVerifyFormData(ViewState("eFormSN").ToString(), "1") '設定瀏覽表單顯示資料
            ElseIf ViewState("read_only").ToString().Equals("2") Then 'read_only="2" 為新增表單
                SettingVerifyFormData(ViewState("eFormSN").ToString(), "2") '設定審核表單顯示資料
            End If
        End If
    End Sub

    ''' <summary>
    ''' 設定新增表單顯示資料
    ''' </summary>
    ''' <param name="fillOutBy">填表人ID</param>
    ''' <remarks></remarks>
    Private Sub SettingNewFormData(ByVal fillOutBy As String)
        '填表人資料
        Dim dtFillOut As DataTable = GetEmployeeAndUnit(fillOutBy)
        lbFillOutUnit.Text = dtFillOut.Rows(0)("ORG_NAME").ToString()         '填表人單位
        lbFillOutName.Text = dtFillOut.Rows(0)("emp_chinese_name").ToString() '填表人姓名
        lbFillOutTitle.Text = dtFillOut.Rows(0)("AD_Title").ToString()        '填表人級職

        SettingLevel2EmployeeDropDownList(fillOutBy) '設定申請人下拉選單
        lbCreatorName.Visible = False

        '申請人資料(預設同填表人)
        lbCreatorUnit.Text = dtFillOut.Rows(0)("ORG_NAME").ToString()  '申請人單位
        ddlCreatorName.SelectedValue = fillOutBy.ToUpper()             '申請人姓名
        lbCreatorTitle.Text = dtFillOut.Rows(0)("AD_Title").ToString() '申請人級職

        lbCreateDate.Text = DateTime.Now.ToString()               '申請時間
        lbSponsor.Text = dtFillOut.Rows(0)("ORG_NAME").ToString() '承辦單位(同申請人單位)

        btnSubmit.Visible = True   '送出
        btnApprove.Visible = False '核准
        btnReject.Visible = False  '駁回
    End Sub

    ''' <summary>
    ''' 取得員工及其所屬單位資料
    ''' </summary>
    ''' <param name="employeeId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetEmployeeAndUnit(ByVal employeeId As String) As DataTable
        db.Open()
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT ORG_NAME,emp_chinese_name,AD_Title FROM EMPLOYEE AS E JOIN ADMINGROUP AS A ON E.ORG_UID = A.ORG_UID WHERE employee_id = '" + employeeId + "'", db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        GetEmployeeAndUnit = dt
    End Function

    '設定申請人下拉選單(即填表人所屬二級單位內所有成員)
    Private Sub SettingLevel2EmployeeDropDownList(ByVal fillOutBy As String)
        Dim dt As DataTable = GetOrganizeLevel2Employees(fillOutBy)
        For Each dr As DataRow In dt.Rows
            ddlCreatorName.Items.Add(New ListItem(dr("emp_chinese_name").ToString(), dr("employee_id").ToUpper().ToString()))
        Next
    End Sub

    ''' <summary>
    ''' 取得成員所屬二級單位內所有成員資料
    ''' </summary>
    ''' <param name="employee_id">目標成員員工ID</param>
    ''' <returns>所屬二級單位內所有成員資料</returns>
    ''' <remarks></remarks>
    Private Function GetOrganizeLevel2Employees(ByVal employee_id As String) As DataTable
        db.Open()
        Dim organizeId2 As String = GetLevelOrganizeId2(db, employee_id) '成員所屬二級單位ID
        Dim orgTreeId2 As String = GetOrganizeIdTree(db, organizeId2) '二級單位樹狀結構底下所有之單位ID字串 格式: ORG_UID1,ORG_UID2.....
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM EMPLOYEE WHERE ORG_UID IN (" + orgTreeId2 + ") ORDER BY emp_chinese_name", db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        GetOrganizeLevel2Employees = dt
    End Function

    '取得成員所屬一或二級單位ID
    'employeeId 成員ID
    'OrgUid 成員所屬單位ID,遞迴用,初始給空值
    'OrgKind 尋找之單位級別
    Private Function GetLevelOrganizeId2(ByRef db As SqlConnection, ByVal employeeId As String) As String
        Dim SqlCom As String = "SELECT * FROM ADMINGROUP WHERE ORG_UID = (SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" + employeeId + "' )"
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(SqlCom, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        'If dt.Rows(0)("ORG_KIND").Equals(Convert.ToInt32(OrgKind)) Then
        GetLevelOrganizeId2 = dt.Rows(0)("ORG_UID")
        'Else
        'GetLevelOrganizeId2 = GetLevelOrganizeId(db, employeeId, dt.Rows(0)("PARENT_ORG_UID"), OrgKind)
        'End If
    End Function

    '取得成員所屬一或二級單位ID
    'employeeId 成員ID
    'OrgUid 成員所屬單位ID,遞迴用,初始給空值
    'OrgKind 尋找之單位級別
    Private Function GetLevelOrganizeId(ByRef db As SqlConnection, ByVal employeeId As String, ByVal OrgUid As String, ByVal OrgKind As String) As String
        Dim SqlCom As String = IIf(String.IsNullOrEmpty(OrgUid), _
                                   "SELECT * FROM ADMINGROUP WHERE ORG_UID = (SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" + employeeId + "' )", _
                                   "SELECT * FROM ADMINGROUP WHERE ORG_UID = '" + OrgUid + "'")
        Dim dt As New DataTable
        Dim ds As New DataSet
        
        Dim da As SqlDataAdapter = New SqlDataAdapter(SqlCom, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        If dt.Rows(0)("ORG_KIND").Equals(Convert.ToInt32(OrgKind)) Then
            GetLevelOrganizeId = dt.Rows(0)("ORG_UID")
        Else
            GetLevelOrganizeId = GetLevelOrganizeId(db, employeeId, dt.Rows(0)("PARENT_ORG_UID"), OrgKind)
        End If
    End Function

    '取得一或二級單位及其樹狀結構底下所有之單位ID字串 格式: ORG_UID1,ORG_UID2.....
    'OrgUid 目標一或二級單位
    Private Function GetOrganizeIdTree(ByRef db As SqlConnection, ByVal OrgUid As String) As String
        Dim strResult As String = ""
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID = '" + OrgUid + "'", db)
        da.Fill(ds)
        dt = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                strResult += "'" + OrgUid + "'," + GetOrganizeIdTree(db, dr("ORG_UID"))
            Next
        Else
            strResult = "'" + OrgUid + "'"
        End If
        GetOrganizeIdTree = strResult
    End Function

    ''' <summary>
    ''' 設定審核表單顯示資料
    ''' </summary>
    ''' <param name="eFormSN">表單ID</param>
    ''' <param name="read_Only">表單讀取型態</param>
    ''' <remarks></remarks>
    Private Sub SettingVerifyFormData(ByVal eFormSN As String, ByVal read_Only As String)
        db.Open()
        Dim strSQL As String = "SELECT P.*,AF.ORG_NAME AS ORG_NAME_F,EF.emp_chinese_name AS emp_chinese_name_F,EF.AD_Title AS AD_Title_F"
        strSQL = strSQL + ",AC.ORG_NAME AS ORG_NAME_C,EC.emp_chinese_name AS emp_chinese_name_C,EC.AD_Title AS AD_Title_C FROM P_09 AS P "
        strSQL = strSQL + "JOIN EMPLOYEE AS EF ON P.FillOutBy = EF.employee_id JOIN ADMINGROUP AS AF ON EF.ORG_UID = AF.ORG_UID "
        strSQL = strSQL + "JOIN EMPLOYEE AS EC ON P.CreateBy = EC.employee_id JOIN ADMINGROUP AS AC ON EC.ORG_UID = AC.ORG_UID "
        strSQL = strSQL + "WHERE EFORMSN = @EFORMSN"
        Dim strComm As New SqlCommand(strSQL, db)
        strComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eFormSN.Trim()
        Dim reader = strComm.ExecuteReader()
        If reader.Read() Then
            lbFillOutUnit.Text = reader("ORG_NAME_F")         '填表人單位
            lbFillOutName.Text = reader("emp_chinese_name_F") '填表人姓名
            lbFillOutTitle.Text = reader("AD_Title_F")        '填表人級職
            lbCreatorUnit.Text = reader("ORG_NAME_C")         '申請人單位
            ddlCreatorName.Visible = False
            lbCreatorName.Text = reader("emp_chinese_name_C") '申請人姓名
            lbCreatorTitle.Text = reader("AD_Title_C")        '申請人級職
            txtSubject.Text = reader("Subject")               '開會事由
            txtSubject.Enabled = False
            lbCreateDate.Text = reader("CreateDate")          '申請時間
            txtLocation.Text = reader("Location")             '開會地點
            txtLocation.Enabled = False
            lbSponsor.Text = reader("Sponsor")                '承辦單位(同申請人單位)
            txtModerator.Text = reader("Moderator")           '主持人
            txtModerator.Enabled = False
            txtPhoneNumber.Text = reader("PhoneNumber")       '聯絡人電話
            txtPhoneNumber.Enabled = False
            txtMeetingDate.Text = Convert.ToDateTime(reader("MeetingDate")).ToString("yyyy/MM/dd") '開會日期
            txtMeetingDate.Enabled = False
            ImgDate1.Visible = False
            ddlHour.SelectedValue = reader("MeetingHour")     '開會時間
            ddlHour.Enabled = False
            ddlMinute.SelectedValue = reader("MeetingMinute") '開會時間
            ddlMinute.Enabled = False
            lbDocumentNo.Text = reader("DocumentNo")          '發文字號
            txtEnteringPeopleNumber.Text = reader("EnteringPeopleNumber") '進出人員
            txtEnteringPeopleNumber.Enabled = False
        End If
        reader.Close()
        strComm.Dispose()

        Dim fileComm As New SqlCommand("SELECT * FROM UPLOAD WHERE eformsn = @eformsn", db)
        fileComm.Parameters.Add("@eformsn", SqlDbType.VarChar, 16).Value = eFormSN.Trim()
        Dim fileReader = fileComm.ExecuteReader()
        While fileReader.Read() '增加附件檔案連結
            Dim strFileName As String = fileReader("FileName").ToString().Split("-")(1)
            Dim strFilePath As String = fileReader("FilePath") + fileReader("FileName")
            divFiles.InnerHtml = divFiles.InnerHtml + "<a href='" + strFilePath + "'>" + strFileName + "</a>" + "<br />"
        End While
        fileReader.Close()
        fileComm.Dispose()
        db.Close()

        hidEnteringGate.Value = GetEnteringGateString(eFormSN) '進出營門

        btnSubmit.Visible = False '送出
        If read_Only.Equals("1") Then 'read_only = "1", 讀取狀態為瀏覽表單,故不顯示核准及駁回按鈕
            btnApprove.Visible = False '核准
            btnReject.Visible = False  '駁回
        ElseIf read_Only.Equals("2") Then 'read_only = "2", 讀取狀態為審核表單,故顯示核准及駁回按鈕
            btnApprove.Visible = True '核准
            btnReject.Visible = True  '駁回
        End If
    End Sub

    Private Function GetEnteringGateString(ByVal eformsn As String) As String
        db.Open()
        Dim strSQL As String = "SELECT * FROM P_0901 WHERE EFORMSN = @EFORMSN"
        Dim sqlComm As New SqlCommand(strSQL, db)
        sqlComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eformsn
        Dim ds = New DataSet()
        Dim da = New SqlDataAdapter(sqlComm)
        da.Fill(ds)
        db.Close()
        Dim strEnteringGates As String = ""
        Dim dt As DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            strEnteringGates = dt.Rows(0)("GateNumber").ToString()
            For i As Integer = 1 To dt.Rows.Count - 1
                strEnteringGates = strEnteringGates + "," + dt.Rows(i)("GateNumber").ToString()
            Next
        End If

        Return strEnteringGates
    End Function

    Protected Sub ImgDate1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Div_grid.Visible = True
        Div_grid.Style("Top") = "200px"
        Div_grid.Style("left") = "310px"
    End Sub

    Protected Sub btnClose1_Click(sender As Object, e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar1_SelectionChanged(sender As Object, e As System.EventArgs) Handles Calendar1.SelectionChanged
        txtMeetingDate.Text = Calendar1.SelectedDate.ToString("yyyy/MM/dd")
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar1_DayRender(sender As Object, e As System.Web.UI.WebControls.DayRenderEventArgs) Handles Calendar1.DayRender
        If e.Day.Date < DateTime.Now.Date Then
            e.Day.IsSelectable = False
        End If
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim files As HttpFileCollection = Request.Files
        If Not CheckFileUpload(files) Then '檢查是否有上傳附件檔案
            Response.Write("<script type='text/javascript'>")
            Response.Write("alert('請上傳附件');")
            Response.Write("</script>")
            Exit Sub
        End If
        If CheckFileExtension(files) Then '檢查上傳檔案中是否包含tec檔
            Response.Write("<script type='text/javascript'>")
            Response.Write("alert('請勿上傳tec附件檔案');")
            Response.Write("</script>")
            Exit Sub
        End If

        Dim dtMeetingControlMembers As DataTable = GetMeetingControlMembers() '取得門禁會議管制群組成員員工ID資料
        If dtMeetingControlMembers.Rows.Count > 1 Then '若審核群組成員超過兩人則轉跳至選擇審核者畫面
            Dim strParameters = "eFormID=" + ViewState("eFormID").ToString() + "&eFormSN=" + ViewState("eFormSN").ToString() _
                                + "&FillOutBy=" + ViewState("userID").ToString() + "&Sponsor=" + lbSponsor.Text
            Server.Transfer("MOA09005.aspx?" + strParameters, True)
        End If

        Dim strCreateBy As String = ddlCreatorName.SelectedValue '申請人員工ID
        Dim strSubject As String = txtSubject.Text         '開會事由
        Dim strLocation As String = txtLocation.Text       '開會地點
        Dim strSponsor As String = lbSponsor.Text          '承辦單位
        Dim strModerator As String = txtModerator.Text     '主持人
        Dim strPhoneNumber As String = txtPhoneNumber.Text '聯絡人電話
        Dim strMeetingDate As String = txtMeetingDate.Text '開會日期
        Dim strHour As String = ddlHour.SelectedValue      '開會時間(時)
        Dim strMinute As String = ddlMinute.SelectedValue  '開會時間(分)
        Dim strDocumentNo As String = txtDocumentWord.Text + "字第" & txtDocumentNo.Text + "號" '發文字號
        Dim strEnteringPeopleNumber As String = txtEnteringPeopleNumber.Text '進出人員(部外)
        '進出營門 0:博愛營區一號門 1:博愛營區二號門 2:博愛營區三號門 3:博愛營區四號門 4:採購中心 5:博一大樓 6:博二大樓
        Dim strEnteringGate As String = hidEnteringGate.Value '資料字串紀錄格式: 營門編號1,營門編號2,...

        Dim redirectFlag As Boolean = True                 '執行是否成功,若成功則轉址
        Dim Val_P As String = ""                           '表單送件回傳值用變數
        '取得欲呈送之門禁會議管制群組成員資料(審核者),若有兩人以上則會被跳轉至選擇頁面,只有一人則直接取其資料
        Dim drEmployeeApprove As DataRow = GetEmployee(dtMeetingControlMembers.Rows(0)("employee_id")).Rows(0)
        Try
            db.Open()
            '新增UPLOAD附件檔案上傳資料,並儲存檔案
            Dim saveFileResult As String = SavingUploadFiles(db, ViewState("eFormSN").ToString(), files)
            If Not saveFileResult.Equals("OK") Then
                Throw New Exception(saveFileResult)
            End If

            '表單審核,FCC.SendVal用參數,格式: 表單種類ID,申請者ID,eFormID,1,審核者ID
            Dim SendVal = ViewState("eFormID").ToString() & "," & strCreateBy & "," & ViewState("eFormSN").ToString() & "," & "1" & "," & drEmployeeApprove("employee_id").ToString()

            '表單送件,寄信及增加flowctl資料表內申請人及其下一關審核者資料
            Dim FCC As New CFlowSend
            Val_P = FCC.F_Send(SendVal, connstr)

            '新增P_09門禁會議管制表單資料,insertP_09函數若執行成功則回傳該筆Insert資料之PK,失敗則回傳錯誤訊息
            Dim insertP_09Result As String = InserP_09(db, ViewState("eFormSN").ToString(), strCreateBy, strSubject _
                                                       , strLocation, strSponsor, strModerator, strPhoneNumber _
                                                       , strMeetingDate, strHour, strMinute, strDocumentNo _
                                                       , strEnteringPeopleNumber, strEnteringGate)
            If Not insertP_09Result.Equals("OK") Then
                Throw New Exception(insertP_09Result)
            End If
        Catch ex As Exception
            redirectFlag = False
            Response.Write("<script type='text/javascript'>")
            Response.Write("alert('申請單送出失敗:" + ex.Message + "');")
            Response.Write("</script>")
            Exit Sub
        Finally
            db.Close()
        End Try
        If redirectFlag Then '若程式執行成功,則將頁面轉址
            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
        End If
    End Sub

    ''' <summary>
    ''' 檢查是否有上傳檔案
    ''' </summary>
    ''' <param name="files">FileUpload集合</param>
    ''' <returns>是否有上傳檔案</returns>
    ''' <remarks></remarks>
    Private Function CheckFileUpload(ByVal files As HttpFileCollection) As Boolean
        Dim result As Boolean = False
        For i As Integer = 0 To files.Count - 1
            If Not String.IsNullOrEmpty(files(i).FileName) Then
                Return True '只要其中有一個FileUpload元件有上傳檔案即可
            End If
        Next
        Return result
    End Function

    ''' <summary>
    ''' 檢查上傳檔案中是否包含tec檔
    ''' </summary>
    ''' <param name="files">FileUpload集合</param>
    ''' <returns>上傳檔案中是否包含tec檔</returns>
    ''' <remarks></remarks>
    Private Function CheckFileExtension(ByVal files As HttpFileCollection) As Boolean
        Dim result As Boolean = False
        For i As Integer = 0 To files.Count - 1
            Dim array As String() = files(i).FileName.Split(".")
            If array(array.Length - 1).ToUpper().Equals("TEC") Then
                Return True
            End If
        Next
        Return result
    End Function

    '取得門禁會議管制群組成員員工ID資料
    Private Function GetMeetingControlMembers() As DataTable
        db.Open()
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT SU.employee_id FROM SYSTEMOBJUSE AS SU JOIN SYSTEMOBJ AS S ON SU.object_uid = S.object_uid WHERE S.object_name = N'門禁會議管制群組'", db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        GetMeetingControlMembers = dt
    End Function

    ''' <summary>
    ''' 取得員工資料
    ''' </summary>
    ''' <param name="employeeId">員工ID</param>
    ''' <returns>員工資料</returns>
    ''' <remarks></remarks>
    Private Function GetEmployee(ByVal employeeId As String) As DataTable
        db.Open()
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM EMPLOYEE WHERE employee_id = '" + employeeId + "'", db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        GetEmployee = dt
    End Function

    ''' <summary>
    ''' 儲存附件檔案並新增資料庫資料
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eformsn">表單審核流程表單資料ID</param>
    ''' <param name="files">HttpFileCollection</param>
    ''' <returns>執行成功訊息OK或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function SavingUploadFiles(ByRef db As SqlConnection, ByVal eformsn As String, ByVal files As HttpFileCollection) As String
        Dim result As String = "OK"
        Try
            For i As Integer = 0 To files.Count - 1
                If Not String.IsNullOrEmpty(files(i).FileName) Then
                    '儲存檔案名稱加上eformsn產生唯一性
                    Dim fileNameArray As String() = files(i).FileName.Split("\")
                    Dim fileName As String = fileNameArray(fileNameArray.Length - 1)
                    Dim saveFileName = eformsn & "-" & fileName
                    Dim uploadPath = Server.MapPath("/MOA/M_Source/99/" & saveFileName)
                    files(i).SaveAs(uploadPath)

                    Dim insertComm = New SqlCommand("INSERT INTO UPLOAD(eformsn,FileName,FilePath,Upload_Time) VALUES(@eformsn,@FileName,@FilePath,GETDATE())", db)
                    insertComm.Parameters.Add("@eformsn", SqlDbType.VarChar, 16).Value = eformsn.Trim()
                    insertComm.Parameters.Add("@FileName", SqlDbType.VarChar, 100).Value = saveFileName.Trim()
                    insertComm.Parameters.Add("@FilePath", SqlDbType.VarChar, 255).Value = "/MOA/M_Source/99/"
                    insertComm.ExecuteNonQuery()
                    insertComm.Dispose()
                End If
            Next
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 新增P_09門禁管制會議申請表資料
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eformsn">表單審核流程表單資料ID</param>
    ''' <param name="strCreateBy">申請人員工ID</param>
    ''' <param name="strSubject">開會事由</param>
    ''' <param name="strLocation">開會地點</param>
    ''' <param name="strSponsor">承辦單位</param>
    ''' <param name="strModerator">主持人</param>
    ''' <param name="strPhoneNumber">聯絡人電話</param>
    ''' <param name="strMeetingDate">開會日期</param>
    ''' <param name="strHour">開會時間(時)</param>
    ''' <param name="strMinute">開會時間(分)</param>
    ''' <param name="strDocumentNo">發文字號</param>
    ''' <param name="strEnteringPeopleNumber">進出人員數量</param>
    ''' <param name="strEnteringGate">進出營門編號</param>
    ''' <returns>執行成功訊息OK或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function InserP_09(ByRef db As SqlConnection, ByVal eformsn As String, ByVal strCreateBy As String _
                               , ByVal strSubject As String, ByVal strLocation As String, ByVal strSponsor As String _
                               , ByVal strModerator As String, ByVal strPhoneNumber As String, ByVal strMeetingDate As String _
                               , ByVal strHour As String, ByVal strMinute As String, ByVal strDocumentNo As String _
                               , ByVal strEnteringPeopleNumber As String, ByVal strEnteringGate As String) As String
        Dim result As String = "OK"
        Dim strSQL As String = "INSERT INTO P_09(EFORMSN,FillOutBy,CreateBy,Subject,MeetingDate,MeetingHour,MeetingMinute,Location"
        strSQL = strSQL + ",Sponsor,Moderator,PhoneNumber,DocumentNo,EnteringPeopleNumber,Status,CreateDate,PENDFLAG) "
        strSQL = strSQL + "VALUES(@EFORMSN,@FillOutBy,@CreateBy,@Subject,@MeetingDate,@MeetingHour,@MeetingMinute,@Location"
        'Status:表單狀態 [ 0：未登管 1：已登管 2：退件 3：刪除 ]
        strSQL = strSQL + ",@Sponsor,@Moderator,@PhoneNumber,@DocumentNo,@EnteringPeopleNumber,0,GETDATE(),'-')"
        Try
            Dim sqlComm = New SqlCommand(strSQL, db)
            sqlComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eformsn                  '表單審核流程表單資料ID
            sqlComm.Parameters.Add("@FillOutBy", SqlDbType.VarChar, 10).Value = ViewState("userID").ToString() '填表人員工編號
            sqlComm.Parameters.Add("@CreateBy", SqlDbType.VarChar, 10).Value = strCreateBy             '申請人員工編號
            sqlComm.Parameters.Add("@Subject", SqlDbType.NVarChar, 100).Value = strSubject             '開會事由
            sqlComm.Parameters.Add("@MeetingDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strMeetingDate) '開會日期
            sqlComm.Parameters.Add("@MeetingHour", SqlDbType.Int).Value = Convert.ToInt32(strHour)     '開會時間(時)
            sqlComm.Parameters.Add("@MeetingMinute", SqlDbType.Int).Value = Convert.ToInt32(strMinute) '開會時間(時)
            sqlComm.Parameters.Add("@Location", SqlDbType.NVarChar, 100).Value = strLocation           '開會地點
            sqlComm.Parameters.Add("@Sponsor", SqlDbType.NVarChar, 50).Value = strSponsor              '承辦單位
            sqlComm.Parameters.Add("@Moderator", SqlDbType.NVarChar, 50).Value = strModerator          '主持人
            sqlComm.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar, 50).Value = strPhoneNumber      '聯絡人電話
            sqlComm.Parameters.Add("@DocumentNo", SqlDbType.NVarChar, 50).Value = strDocumentNo        '發文字號
            sqlComm.Parameters.Add("@EnteringPeopleNumber", SqlDbType.Int).Value = Convert.ToInt32(strEnteringPeopleNumber) '進出人員數量
            sqlComm.ExecuteNonQuery()
            Dim gateNoArray As String() = strEnteringGate.Split(",")
            For i As Integer = 0 To gateNoArray.Length - 1
                '進出營門 0:博愛營區一號門 1:博愛營區二號門 2:博愛營區三號門 3:博愛營區四號門 4:採購中心 5:博一大樓 6:博二大樓
                Dim insertP_0901Result As String = InserP_0901(db, eformsn, gateNoArray(i))
                If Not insertP_0901Result.Equals("OK") Then
                    Throw New Exception(insertP_0901Result)
                End If
            Next
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 新增P_0901門禁管制會議申請表資料
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eFormSN">表單審核流程表單資料ID</param>
    ''' <param name="gateNo">出入營門編號</param>
    ''' <returns>執行成功訊息OK或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function InserP_0901(ByRef db As SqlConnection, ByVal eFormSN As String, ByVal gateNo As String) As String
        Dim result As String = "OK"
        Dim strSQL As String = "INSERT INTO P_0901(EFORMSN,GateNumber) VALUES(@EFORMSN,@GateNumber)"
        Try
            Dim sqlComm = New SqlCommand(strSQL, db)
            sqlComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eFormSN '表單審核流程表單資料ID
            '進出營門 0:博愛營區一號門 1:博愛營區二號門 2:博愛營區三號門 3:博愛營區四號門 4:採購中心 5:博一大樓 6:博二大樓
            sqlComm.Parameters.Add("@GateNumber", SqlDbType.Int).Value = Convert.ToInt32(gateNo)
            sqlComm.ExecuteNonQuery()
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    Protected Sub btnApprove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        Dim bl_executeResult = False
        Dim conn As SqlConnection = New SqlConnection(connstr)
        Try
            conn.Open()
            '表單狀態 [ 0：未登管 1：已登管 2：退件 3：刪除 ]
            Dim myCommand As New SqlCommand("UPDATE P_09 SET Status = 1, ModifyBy = @ModifyBy, ModifyDate = GETDATE() WHERE EFORMSN = @EFORMSN", conn)
            myCommand.Parameters.Add("@ModifyBy", SqlDbType.VarChar, 10).Value = ViewState("userID").ToString()
            myCommand.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = ViewState("eFormSN").ToString()
            myCommand.ExecuteNonQuery()
            bl_executeResult = True
        Catch ex As Exception
            bl_executeResult = False
        Finally
            conn.Close()
        End Try

        If bl_executeResult = False Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('審核失敗，請重新再試，或連絡系統管理人員！');")
            Response.Write(" </script>")
            Exit Sub
        ElseIf ViewState("read_only").ToString() = "2" Then '若為審核表單,則審核通過後,將母視窗資料重新整理更新
            Response.Write(" <script language='javascript'>")
            Response.Write(" window.dialogArguments.document.location.href = window.dialogArguments.document.location;")
            Response.Write(" </script>")
        End If

        insComment(txtcomment.Text.Trim(), ViewState("eFormSN").ToString(), ViewState("userID").ToString()) '增加批核意見
        '表單審核
        Dim FCC As New CFlowSend
        Dim SendVal As String = ViewState("eFormID").ToString() & "," & ViewState("userID").ToString() & "," & ViewState("eFormSN").ToString() & "," & "1" & ","
        Dim Val_P As String = FCC.F_Send(SendVal, connstr)
        Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=")
    End Sub

    ''' <summary>
    ''' 增加批核意見
    ''' </summary>
    ''' <param name="strComment">批核意見</param>
    ''' <param name="strEformsn">表單ID</param>
    ''' <param name="strid">批核者ID</param>
    ''' <remarks></remarks>
    Private Sub insComment(ByVal strComment As String, ByVal strEformsn As String, ByVal strid As String)
        db.Open()
        Dim insCom As New SqlCommand("UPDATE flowctl SET comment='" & strComment & "' WHERE hddate IS NULL AND eformsn='" & strEformsn & "' AND empuid='" & strid & "'", db)
        insCom.ExecuteNonQuery()
        db.Close()
    End Sub

    Protected Sub btnReject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReject.Click
        Try
            Dim conn As SqlConnection = New SqlConnection(connstr)
            conn.Open()
            '表單狀態 [ 0：未登管 1：已登管 2：退件 3：刪除 ]
            Dim myCommand As New SqlCommand("UPDATE P_09 SET Status = 2, ModifyBy = @ModifyBy, ModifyDate = GETDATE() WHERE EFORMSN = @EFORMSN", conn)
            myCommand.Parameters.Add("@ModifyBy", SqlDbType.VarChar, 10).Value = ViewState("userID").ToString()
            myCommand.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = ViewState("eFormSN").ToString()
            myCommand.ExecuteNonQuery()
            conn.Close()

            insComment(txtcomment.Text.Trim(), ViewState("eFormSN").ToString(), ViewState("userID").ToString()) '增加批核意見
            Dim Val_P As String = ""
            '表單駁回
            'Dim FC As New C_FlowSend.C_FlowSend
			Dim FC As New CFlowSend
            Dim SendVal As String = ViewState("eFormID").ToString() & "," & ViewState("userID").ToString() & "," & ViewState("eFormSN").ToString() & "," & "1"
            Val_P = FC.F_Back(SendVal, connstr)

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單已駁回給申請人');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")

        Catch ex As Exception
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('駁回表單失敗：" + ex.Message + "');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")
        End Try
    End Sub

    ''' <summary>
    ''' 取得員工姓名
    ''' </summary>
    ''' <param name="employeeID">員工ID</param>
    ''' <returns>員工中文姓名</returns>
    ''' <remarks></remarks>
    Private Function GetEmployeeName(ByVal employeeID As String) As String
        Dim empName As String = ""
        db.Open()
        Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = @employee_id", db)
        strPer.Parameters.Add("@employee_id", SqlDbType.VarChar, 10).Value = employeeID.Trim()
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.Read() Then
            empName = RdPer("emp_chinese_name")
        End If
        RdPer.Close()
        db.Close()
        Return empName
    End Function

    Protected Sub But_PHRASE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_PHRASE.Click
        Dim Visitors As Integer
        Dim str_SQL As String = ""

        '是否有個人批核片語資料
        db.Open()
        Dim PerCountCom As New SqlCommand("SELECT count(*) as Visitors from Phrase WHERE employee_id = '" & ViewState("userID").ToString() & "'", db)
        Dim PerRdv = PerCountCom.ExecuteReader()
        If PerRdv.Read() Then
            Visitors = PerRdv("Visitors")
        End If
        db.Close()

        If Visitors > 0 Then
            Div_grid10.Visible = True
            Div_grid10.Style("Top") = "300px"
            Div_grid10.Style("left") = "350px"
        Else
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('無批核片語資料,請至批核片語管理新增資料');")
            Response.Write(" </script>")
        End If
    End Sub

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        txtcomment.Text = GridView2.Rows(GridView2.SelectedRow.RowIndex).Cells(0).Text
        Div_grid10.Visible = False
    End Sub

    Protected Sub Btn_PHclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_PHclose.Click
        Div_grid10.Visible = False
    End Sub

    Protected Sub ddlCreatorName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCreatorName.SelectedIndexChanged
        '申請人資料
        Dim dtCreateBy As DataTable = GetEmployeeAndUnit(ddlCreatorName.SelectedValue.Trim())
        lbCreatorUnit.Text = dtCreateBy.Rows(0)("ORG_NAME").ToString()     '申請人單位
        ddlCreatorName.SelectedValue = ddlCreatorName.SelectedValue.Trim() '申請人姓名
        lbCreatorTitle.Text = dtCreateBy.Rows(0)("AD_Title").ToString()    '申請人級職
        lbSponsor.Text = dtCreateBy.Rows(0)("ORG_NAME").ToString()         '承辦單位(同申請人單位)
    End Sub
End Class
