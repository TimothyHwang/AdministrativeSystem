Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.Common
Imports System.Web.UI
Imports System.Collections.Generic

Partial Class M_Source_08_MOA08012
    Inherits System.Web.UI.Page
    Dim sql_function As New C_SQLFUN
    Dim connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim user_id As String = Session("user_id")
        'Session被清空回首頁
        If user_id = "" Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")
            Exit Sub
        End If
        If Not Page.IsPostBack Then
            Dim dtRoleGroup As DataTable = GetROLEGROUPGroupUid() '權限群組資料
            ViewState("ManagersOfCopierID") = dtRoleGroup.Select("Group_Name='影印機管理者'")(0)("Group_Uid")
            ViewState("ConfidentialSupervisoryOfficerID") = dtRoleGroup.Select("Group_Name='保密督導官'")(0)("Group_Uid")
            Dim ListEmployeeGroup = GetEmployeeRoleGroup(user_id) '使用者所組群組ID
            Dim emploeeRole As String = "" '使用者所組群組 1:影印機管理者 3:保密督導官
            If ListEmployeeGroup.Contains(ViewState("ConfidentialSupervisoryOfficerID")) Then
                emploeeRole = "3"
            ElseIf ListEmployeeGroup.Contains(ViewState("ManagersOfCopierID")) Then
                emploeeRole = "1"
            Else
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('您無權限進入本頁面觀看功能!!');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
                Exit Sub
            End If
            ViewState("emploeeRole") = emploeeRole

            ddlORG_UID.Items.Clear()
            ddlORG_UID.Items.Add(New ListItem("", ""))
            ddlORG_UID.Items.AddRange(GetAdminGroup().ToArray())
            Dim dtP_08 As DataTable = GetP_08("", "", "", "", "", "", "", "")
            ViewState("DataTableP_08") = dtP_08
            gvP_08.DataSource = dtP_08
            gvP_08.DataBind()
        End If
    End Sub

    '取得權限群組資料
    Private Function GetROLEGROUPGroupUid() As DataTable
        Dim dt As New DataTable
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()

        Dim strSQL = "SELECT * FROM ROLEGROUP"
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        dt = ds.Tables(0)

        db.Close()
        GetROLEGROUPGroupUid = dt
    End Function

    '取得使用者所屬權限群組ID
    Private Function GetEmployeeRoleGroup(ByVal employee_id As String) As List(Of String)
        Dim employeeRoleGroups As New List(Of String)
        Dim dt As DataTable = Nothing
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()

        Dim strSQL = "SELECT RGI.Group_Uid, RG.Group_Name FROM ROLEGROUPITEM AS RGI JOIN ROLEGROUP AS RG ON RGI.Group_Uid = RG.Group_Uid"
        strSQL = strSQL + " WHERE employee_id='" + employee_id + "'"
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        For Each dr As DataRow In dt.Rows
            employeeRoleGroups.Add(dr("Group_Uid").ToString())
        Next
        db.Close()

        GetEmployeeRoleGroup = employeeRoleGroups
    End Function

    '取得全部單位資料之ListItem型態 Text:ORG_NAME Value:ORG_UID
    Private Function GetAdminGroup() As List(Of ListItem)
        Dim adminGroupList As New List(Of ListItem)
        db.Open()

        Dim strSQL = "SELECT * FROM ADMINGROUP"
        Dim dt As DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        For Each dr As DataRow In dt.Rows
            adminGroupList.Add(New ListItem(dr("ORG_NAME"), dr("ORG_UID")))
        Next
        GetAdminGroup = adminGroupList
        db.Close()
    End Function

    '取得影印紀錄資料
    Private Function GetP_08(ByVal strPrint_DateBegin As String, ByVal strPrint_DateEnd As String, ByVal strORG_UID As String _
                             , ByVal strPrint_Name As String, ByVal strFile_Name As String, ByVal strSecurity_Status As String _
                             , ByVal strUse_For As String, ByVal strPrint_Num As String) As DataTable
        Dim strCondition As String = ""
        If Not String.IsNullOrEmpty(strPrint_DateBegin) Then '複印時間起始
            strCondition += " AND Print_Date > '" + strPrint_DateBegin + "'"
        End If
        If Not String.IsNullOrEmpty(strPrint_DateEnd) Then '複印時間終止
            strCondition += " AND Print_Date < '" + strPrint_DateEnd + " 23:59:59'"
        End If
        If Not String.IsNullOrEmpty(strORG_UID) Then '使用單位
            strCondition += " AND P.ORG_UID = '" + strORG_UID + "'"
        End If
        If Not String.IsNullOrEmpty(strPrint_Name) Then '姓名
            strCondition += " AND Print_Name LIKE '%" + strPrint_Name + "%'"
        End If
        If Not String.IsNullOrEmpty(strFile_Name) Then '複印資料名稱
            strCondition += " AND File_Name LIKE '%" + strFile_Name + "%'"
        End If
        If Not String.IsNullOrEmpty(strSecurity_Status) Then '密等
            strCondition += " AND Security_Status = '" + strSecurity_Status + "'"
        End If
        If Not String.IsNullOrEmpty(strUse_For) Then '用途
            strCondition += " AND Use_For LIKE '%" + strUse_For + "%'"
        End If
        If Not String.IsNullOrEmpty(strPrint_Num) Then '流水號
            strCondition += " AND Print_Num LIKE '%" + strPrint_Num + "%'"
        End If

        Dim dt As New DataTable
        db.Open()
        'Status:4 補登完畢資料
        Dim strSQL = "SELECT P.*, A.ORG_NAME FROM P_08 AS P "
        strSQL = strSQL + "LEFT JOIN ADMINGROUP AS A ON P.ORG_UID = A.ORG_UID "
        strSQL = strSQL + "WHERE P.Printer_No IN (" + GetPrintersString(Session("user_id").ToString(), ViewState("emploeeRole").ToString()) + ") "
        strSQL = strSQL + "AND (ApprovedByID IS NULL OR ApprovedByID = '') "
        If Not String.IsNullOrEmpty(strCondition) Then
            strSQL += strCondition
        End If
        strSQL += " ORDER BY P.Print_Date"
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        GetP_08 = dt
        db.Close()
    End Function

    '取得人員所屬影印機列表字串, 格式: Printer_No1,Printer_No2.....
    Private Function GetPrintersString(ByVal employee_id As String, ByVal emploeeRole As String) As String
        Dim strResult As String = ""
        Dim strSQLCom As String = ""

        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()

        Select Case emploeeRole '使用者影印機管理權限身分
            Case "1" '影印機管理者：可查詢該影印機的所有影印紀錄。
                strSQLCom = "SELECT * FROM P_0803 WHERE employee_id = '" + employee_id + "'"
            Case "3" '保密督導官：可查詢所屬2級單位影印機的所有影印紀錄。
                Dim organizeId2 As String = GetLevelOrganizeId(db, employee_id, "", 2) '成員所屬二級單位ID
                Dim orgTreeId2 As String = GetOrganizeIdTree(db, organizeId2) '二級單位樹狀結構底下所有之單位ID字串 格式: ORG_UID1,ORG_UID2.....
                strSQLCom = "SELECT * FROM P_0803 WHERE ORG_UID IN (" + orgTreeId2 + ")"
            Case Else
        End Select
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQLCom, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            strResult += dt.Rows(0)("Printer_No")
            For i As Integer = 1 To dt.Rows.Count - 1
                strResult += "," + dt.Rows(i)("Printer_No")
            Next
        End If

        db.Close()
        GetPrintersString = strResult
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

    Protected Sub ImgDate1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "70px"
        If txtPrint_DateBegin.Text.Trim() = "" Then
            Calendar1.SelectedDate = Date.Now.AddDays(-14).Date
        Else
            Calendar1.SelectedDate = txtPrint_DateBegin.Text
        End If
        txtPrint_DateBegin.Text = Calendar1.SelectedDate.Date
    End Sub

    Protected Sub ImgDate2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "130px"
        If (txtPrint_DateEnd.Text.Trim() = "") Then
            Calendar2.SelectedDate = Date.Now.Date
        Else
            Calendar2.SelectedDate = txtPrint_DateEnd.Text
        End If
        txtPrint_DateEnd.Text = Calendar2.SelectedDate.Date
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        txtPrint_DateBegin.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        txtPrint_DateEnd.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False
    End Sub

    Protected Sub btnClose1_Click(sender As Object, e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose2_Click(sender As Object, e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub

    Protected Sub gvP_08_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvP_08.RowCommand
        Dim currentCommand As String = e.CommandName
        Select Case currentCommand
            Case "Detail"
                Dim objCurButton As ImageButton = DirectCast(e.CommandSource, ImageButton)
                Dim curGridViewRow As GridViewRow = DirectCast(objCurButton.NamingContainer, GridViewRow)
                Dim HF_Status As HiddenField = CType(curGridViewRow.FindControl("HF_Status"), HiddenField)
                Server.Transfer("MOA08002.aspx?Log_Guid=" & e.CommandArgument.ToString() + "&Status=" + HF_Status.Value)
        End Select
    End Sub

    Protected Sub gvP_08_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvP_08.PageIndexChanging
        gvP_08.PageIndex = e.NewPageIndex
        Dim dtP_08 As DataTable = ViewState("DataTableP_08")
        gvP_08.DataSource = dtP_08
        gvP_08.DataBind()
    End Sub

    '轉換前端頁面GridView密等顯示文字
    Public Function displaySecurityStatus(ByVal statusNo As Int32) As String
        Select Case statusNo
            Case 1
                displaySecurityStatus = "普通"
            Case 2
                displaySecurityStatus = "密"
            Case 3
                displaySecurityStatus = "機密"
            Case 4
                displaySecurityStatus = "極機密"
            Case 5
                displaySecurityStatus = "絕對機密"
            Case Else
                displaySecurityStatus = ""
        End Select
    End Function

    '轉換前端頁面GridView資料張數顯示文字
    Public Function displayCopyDetail(ByVal copyA3M As String, ByVal copyA4M As String, ByVal copyA3C As String, ByVal copyA4C As String, ByVal scan As String) As String
        Dim str As String = ""
        If Not String.IsNullOrEmpty(copyA3M) And Not copyA3M.Equals("0") Then
            str = str + "A3黑白: " + copyA3M + "<br />"
        End If
        If Not String.IsNullOrEmpty(copyA4M) And Not copyA4M.Equals("0") Then
            str = str + "A4黑白: " + copyA4M + "<br />"
        End If
        If Not String.IsNullOrEmpty(copyA3C) And Not copyA3C.Equals("0") Then
            str = str + "A3彩色: " + copyA3C + "<br />"
        End If
        If Not String.IsNullOrEmpty(copyA4C) And Not copyA4C.Equals("0") Then
            str = str + "A4彩色: " + copyA4C + "<br />"
        End If
        If Not String.IsNullOrEmpty(scan) And Not scan.Equals("0") Then
            str = str + "掃描: " + scan
        End If
        displayCopyDetail = str
    End Function

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        Dim strPrint_DateBegin As String = txtPrint_DateBegin.Text '複印時間起始
        Dim strPrint_DateEnd As String = txtPrint_DateEnd.Text '複印時間終止
        Dim strORG_UID As String = ddlORG_UID.Text '使用單位
        Dim strPrint_Name As String = txtPrint_Name.Text '姓名
        Dim strFile_Name As String = txtFile_Name.Text '複印資料名稱
        Dim strSecurity_Status As String = ddlSecurity_Status.SelectedValue '密等
        Dim strUse_For As String = txtUse_For.Text '用途
        Dim strPrint_Num As String = txtPrint_Num.Text '流水號
        Dim dtP_08 As DataTable = GetP_08(strPrint_DateBegin, strPrint_DateEnd, strORG_UID, strPrint_Name, strFile_Name, strSecurity_Status, strUse_For, strPrint_Num)
        ViewState("DataTableP_08") = dtP_08
        gvP_08.DataSource = dtP_08
        gvP_08.DataBind()
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim tool As New C_Public
        Dim confidentialSupervisoryOfficerID As String = GetConfidentialSupervisoryOfficerID() '取得呈核者所屬保密督導官ID
        If String.IsNullOrEmpty(confidentialSupervisoryOfficerID) Then '檢查呈核者是否有所屬保密督導官
            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "Msg", "alert('找不到所屬保密督導官');", True)
            Exit Sub
        End If

        Dim Log_GuidArray() As String = selectedLog_Guid.Value.Split(",") '欲呈核影印紀錄資料ID
        If checkReported(selectedLog_Guid.Value) Then '檢查所選資料是否包含已呈核資料
            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "Msg", "alert('所選資料中不可包含已呈核資料,請重新選擇');", True)
            Exit Sub
        End If

        Dim eformsn As String = tool.CreateNewEFormSN() '建立唯一的eformsn (表單審核流程表單資料ID)
        Dim eformid As String = GetEFormId("影印紀錄呈核單") '取得影印紀錄呈核單ID
        Dim dtFlow As DataTable = GetFlow(eformid) '取得影印紀錄呈核單審核關卡流程資料
        Dim drEmployeeApplicant As DataRow = GetEmployee(Session("user_id").ToString()).Rows(0) '取得申請人員工資料
        Dim drApplicant As DataRow = dtFlow.Select("steps=0")(0) '取得初始關卡(即申請人關卡)資料
        Dim drEmployeeApprove As DataRow = GetEmployee(confidentialSupervisoryOfficerID).Rows(0) '取得保密督導官員工資料
        '取得審核關卡(即初始關卡之下一關卡)資料
        Dim drApprove As DataRow = dtFlow.Select("steps=" + drApplicant("nextstep").ToString())(0)

        Dim trans As SqlTransaction = Nothing
        Try
            db.Open()
            trans = db.BeginTransaction()
            '增加flowctl審核流程表單申請人資料
            Dim applicantResult As String = InsertFlowControl(db, trans, eformid, eformsn, drApplicant("stepsid").ToString(), _
                drApplicant("steps").ToString(), drEmployeeApplicant("employee_id").ToString(), drEmployeeApplicant("emp_chinese_name").ToString(), _
                "申請人", "-", "", drApplicant("nextstep").ToString(), "Now", drEmployeeApplicant("ORG_UID").ToString())
            If Not applicantResult.Equals("OK") Then
                Throw New Exception(applicantResult)
            End If

            '增加flowctl審核流程表單保密督導官資料
            Dim approveResult As String = InsertFlowControl(db, trans, eformid, eformsn, drApprove("stepsid").ToString(), _
                drApprove("steps").ToString(), drEmployeeApprove("employee_id").ToString(), drEmployeeApprove("emp_chinese_name").ToString(), _
                drApprove("object_name").ToString(), "?", "", drApprove("nextstep").ToString(), "", drEmployeeApprove("ORG_UID").ToString())
            If Not approveResult.Equals("OK") Then
                Throw New Exception(approveResult)
            End If

            For i As Integer = 0 To Log_GuidArray.Length - 1 '變更表單P_08之VerifyRequesterID欄位,即增加影印紀錄資料呈核送審者
                Dim P_08Result As String = UpdateP_08VerifyRequesterID(db, trans, Log_GuidArray(i), Session("user_id"))
                If Not P_08Result.Equals("OK") Then
                    Throw New Exception(P_08Result)
                End If
            Next

            '增加影印記錄呈核單資料
            Dim reportResult As String = InsertPrintRecordsReport(db, trans, eformsn, selectedLog_Guid.Value, Session("user_id"))
            If Not reportResult.Equals("OK") Then
                Throw New Exception(reportResult)
            End If
            trans.Commit()
            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "Msg", "alert('呈核已送出');location.href='MOA08012.aspx';", True)
        Catch ex As Exception
            trans.Rollback()
            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "Msg", "alert('呈核失敗:" + ex.Message + "');", True)
        Finally
            tool = Nothing
            db.Close()
        End Try
    End Sub

    ''' <summary>
    ''' 取得登入使用者所屬保密督導官
    ''' </summary>
    ''' <returns>保密督導官ID</returns>
    ''' <remarks></remarks>
    Private Function GetConfidentialSupervisoryOfficerID() As String
        Dim confidentialSupervisoryOfficerID As String = ""
        db.Open()
        Dim levelTwoOrgId As String = GetLevelOrganizeId(db, Session("user_id").ToString(), "", "2") '取得成員所屬二級單位ID
        '取得該二級單位及其樹狀結構底下所有之單位ID字串 格式: 'ORG_UID1','ORG_UID2'.....
        Dim strOrgIds As String = GetOrganizeIdTree(db, levelTwoOrgId)
        Dim strSQL As String = "SELECT RGI.Role_Num,RGI.Group_Uid,RGI.employee_id,E.ORG_UID FROM ROLEGROUPITEM AS RGI"
        strSQL = strSQL + " JOIN EMPLOYEE AS E ON RGI.employee_id = E.employee_id"
        strSQL = strSQL + " WHERE RGI.Group_Uid = '" + ViewState("ConfidentialSupervisoryOfficerID").ToString()
        strSQL = strSQL + "' AND E.ORG_UID IN (" + strOrgIds + ")"
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            confidentialSupervisoryOfficerID = ds.Tables(0).Rows(0)("employee_id")
        End If
        db.Close()
        GetConfidentialSupervisoryOfficerID = confidentialSupervisoryOfficerID
    End Function

    ''' <summary>
    ''' 檢查所選資料是否包含已呈核資料,VerifyRequesterID(資料送出呈核者)欄位有值則代表該資料已送出呈核
    ''' </summary>
    ''' <param name="logGuids">影印紀錄資料ID字串,格式:logGuid1,logGuid2...</param>
    ''' <returns>是否包含已呈核資料</returns>
    ''' <remarks></remarks>
    Private Function checkReported(ByVal logGuids As String) As Boolean
        db.Open()
        Dim strSQL = "SELECT * FROM P_08 WHERE Log_Guid IN (" + logGuids + ") AND NOT(VerifyRequesterID IS NULL OR VerifyRequesterID='')"
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            checkReported = True
        Else
            checkReported = False
        End If
        db.Close()
    End Function

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        db.Open()
        Dim sqlcomm As New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db)
        GetEFormId = sqlcomm.ExecuteScalar()
        db.Close()
    End Function

    '取得表單審核流程資料
    Private Function GetFlow(ByVal eformId As String) As DataTable
        db.Open()
        Dim strSQL = "SELECT f.*, S.object_name FROM flow AS f LEFT JOIN SYSTEMOBJ AS S ON f.group_id = S.object_uid WHERE eformid = '" + eformId + "'"
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        GetFlow = dt
    End Function

    '取得表單審核流程資料
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
    ''' 取得呈核表單下一個審核關卡
    ''' </summary>
    ''' <param name="eformSN">呈核表單ID</param>
    ''' <param name="userId">使用者ID</param>
    ''' <returns>呈核表單下一個審核關卡ID</returns>
    ''' <remarks></remarks>
    Private Function GetNextStep(ByVal eformSN As String, ByVal userId As String) As String
        Dim strNextStep As String = ""
        db.Open()
        Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" + eformSN + "' and empuid = '" + userId + "' and hddate is null", db)
        Dim RdrCheck = strComCheck.ExecuteReader()
        If RdrCheck.Read() Then
            strNextStep = RdrCheck.Item("nextstep")
        End If
        RdrCheck.Close()
        db.Close()
        GetNextStep = strNextStep
    End Function

    '更改影印紀錄資料VerifyRequesterID欄位,即申請呈核者
    Private Function UpdateP_08VerifyRequesterID(ByRef db As SqlConnection, ByRef trans As SqlTransaction, ByVal log_Guid As String, ByVal verifyRequesterID As String) As String
        Dim strResult As String = "OK"
        Try
            Dim strSQL = "UPDATE P_08 SET VerifyRequesterID = '" + verifyRequesterID + "' WHERE Log_Guid = '" + log_Guid + "'"
            Dim comm As New SqlCommand(strSQL, db, trans)
            comm.ExecuteNonQuery()
        Catch ex As Exception
            strResult = ex.Message
        End Try
        UpdateP_08VerifyRequesterID = strResult
    End Function

    '新增影印紀錄呈核單資料
    Private Function InsertPrintRecordsReport(ByRef db As SqlConnection, ByRef trans As SqlTransaction, ByVal eformsn As String, ByVal records As String, ByVal createBy As String) As String
        Dim strResult As String = "OK"
        Try
            Dim dateNow As DateTime = DateTime.Now
            Dim strSQL = "INSERT INTO PrintRecordsReport(EFORMSN,Status,CreateDate,CreateBy)"
            strSQL = strSQL + " VALUES(@EFORMSN,1,@CreateDate,@CreateBy)"
            Dim comm As New SqlCommand(strSQL, db, trans)
            comm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eformsn.Trim()
            comm.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = dateNow
            comm.Parameters.Add("@CreateBy", SqlDbType.VarChar, 10).Value = createBy.Trim()
            comm.ExecuteNonQuery()

            '新增影印紀錄呈核單及影印紀錄對應資料
            Dim recordsArray As String() = records.Split(",")
            For i As Integer = 0 To recordsArray.Length - 1
                Dim mappingSQL = "INSERT INTO ReportP_08Mapping(EFORMSN,Log_Guid,CreateDate,CreateBy)"
                mappingSQL = mappingSQL + " VALUES(@EFORMSN,@Log_Guid,@CreateDate,@CreateBy)"
                Dim mappingComm As New SqlCommand(mappingSQL, db, trans)
                mappingComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eformsn.Trim()
                mappingComm.Parameters.Add("@Log_Guid", SqlDbType.Int).Value = Convert.ToInt32(recordsArray(i))
                mappingComm.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = dateNow
                mappingComm.Parameters.Add("@CreateBy", SqlDbType.VarChar, 10).Value = createBy.Trim()
                mappingComm.ExecuteNonQuery()
            Next
        Catch ex As Exception
            strResult = ex.Message
        End Try
        InsertPrintRecordsReport = strResult
    End Function

    '新增flowctl資料
    Private Function InsertFlowControl(ByRef db As SqlConnection, ByRef trans As SqlTransaction, ByVal eformid As String _
                                       , ByVal eformsn As String, ByVal stepsid As String, ByVal steps As String _
                                       , ByVal empuid As String, ByVal emp_chinese_name As String, ByVal group_name As String _
                                       , ByVal gonogo As String, ByVal comment As String, ByVal nextstep As String _
                                       , ByVal recdate As String, ByVal deptcode As String) As String
        Dim strResult As String = "OK"
        Try
            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = db
            sqlComm.Transaction = trans
            Dim strSQL = "INSERT INTO flowctl(eformid,eformrole,eformsn,stepsid,steps,empuid,emp_chinese_name,group_name"
            strSQL = strSQL + ",hddate,gonogo,comment,nextstep,important,recdate,appdate,deptcode,createdate) "
            strSQL = strSQL + "VALUES(@eformid,1,@eformsn,@stepsid,@steps,@empuid,@emp_chinese_name,@group_name"
            strSQL = strSQL + ",@hddate,@gonogo,@comment,@nextstep,'1',@recdate,@appdate,@deptcode,@createdate)"
            sqlComm.CommandText = strSQL
            sqlComm.Parameters.Add("@eformid", SqlDbType.VarChar, 10).Value = eformid
            sqlComm.Parameters.Add("@eformsn", SqlDbType.VarChar, 16).Value = eformsn
            sqlComm.Parameters.Add("@stepsid", SqlDbType.Int).Value = Convert.ToInt32(stepsid)
            sqlComm.Parameters.Add("@steps", SqlDbType.Int).Value = Convert.ToInt32(steps)
            sqlComm.Parameters.Add("@empuid", SqlDbType.VarChar, 10).Value = empuid
            sqlComm.Parameters.Add("@emp_chinese_name", SqlDbType.NVarChar, 50).Value = emp_chinese_name
            sqlComm.Parameters.Add("@group_name", SqlDbType.VarChar, 50).Value = group_name
            Dim dateTimeNow As DateTime = DateTime.Now
            sqlComm.Parameters.Add("@hddate", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(recdate), DBNull.Value, dateTimeNow)
            sqlComm.Parameters.Add("@gonogo", SqlDbType.VarChar, 1).Value = gonogo
            sqlComm.Parameters.Add("@comment", SqlDbType.VarChar, 1200).Value = comment
            sqlComm.Parameters.Add("@nextstep", SqlDbType.Int).Value = Convert.ToInt32(nextstep)
            sqlComm.Parameters.Add("@recdate", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(recdate), DBNull.Value, dateTimeNow)
            sqlComm.Parameters.Add("@appdate", SqlDbType.DateTime).Value = dateTimeNow
            sqlComm.Parameters.Add("@deptcode", SqlDbType.VarChar, 10).Value = deptcode
            sqlComm.Parameters.Add("@createdate", SqlDbType.DateTime).Value = dateTimeNow
            sqlComm.ExecuteNonQuery()
        Catch ex As Exception
            strResult = ex.Message
        End Try
        InsertFlowControl = strResult
    End Function

    '前端頁面GridView批核欄位顯示名稱
    Public Function ShowApproveStatus(ByVal verifyRequesterID As Object, ByVal approvedByID As Object) As String
        Dim strVerifyRequesterID As String = IIf(verifyRequesterID Is Nothing, "", verifyRequesterID.ToString)
        Dim strApprovedByID As String = IIf(approvedByID Is Nothing, "", approvedByID.ToString)

        If Not String.IsNullOrEmpty(strApprovedByID) Then '若ApprovedByID欄位已有紀錄批核者,則為批核已完成
            ShowApproveStatus = "完成"
        ElseIf Not String.IsNullOrEmpty(strVerifyRequesterID) Then '若VerifyRequesterID欄位已有紀錄呈核者,則為已送出呈核,等待批核
            ShowApproveStatus = "待批核"
        Else
            ShowApproveStatus = "未完成"
        End If
    End Function
End Class
