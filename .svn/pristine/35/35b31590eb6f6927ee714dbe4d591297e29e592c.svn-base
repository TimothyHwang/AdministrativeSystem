Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Collections.Generic

Partial Class M_Source_08_MOA08003
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Public aa As New C_CheckFun
    Dim CF As New CFlowSend
    Dim dt As Date = Now()
    Dim user_id, org_uid, Roleid As String
    Dim sQuery As String = "select * from (select distinct Log_Guid,a.Printer_No,c.Printer_Name,PAIDNO,Security_Status,File_Name,Print_Name,b.ORG_Name,a.ORG_UID,a.LogTime,Print_Date,a.[Status],EFORMSN,isnull(Copy_A3M,0) as Copy_A3M,isnull(Copy_A4M,0) as Copy_A4M,isnull(Copy_A3C,0) as Copy_A3C,isnull(Copy_A4C,0) as Copy_A4C,isnull(Scan,0) as Scan, a.VerifyRequesterID, a.ApprovedByID from P_08 a with(nolock) left join [ADMINGROUP] b with(nolock) on a.org_uid = b.org_uid left join P_0803 as c on a.Printer_No = c.Printer_No"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '取得登入者帳號資訊
        If Not Session("ORG_UID") Is Nothing And Not Session("user_id") Is Nothing And Not Session("Role") Is Nothing Then
            user_id = Session("user_id").ToString()
            org_uid = Session("ORG_UID").ToString()
            Roleid = Session("Role").ToString()
        End If

        ErrMsg.Text = ""
        'session被清空回首頁
        If user_id = "" And org_uid = "" And Roleid = "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            If Not IsPostBack Then
                Dim dtRoleGroup As DataTable = GetROLEGROUPGroupUid() '權限群組資料
                ViewState("EndUsersID") = dtRoleGroup.Select("Group_Name='一般使用者'")(0)("Group_Uid")
                ViewState("ManagersOfCopierID") = dtRoleGroup.Select("Group_Name='影印機管理者'")(0)("Group_Uid")
                ViewState("UnitSafeguardsOfficerID") = dtRoleGroup.Select("Group_Name='單位保防官'")(0)("Group_Uid")
                ViewState("ConfidentialSupervisoryOfficerID") = dtRoleGroup.Select("Group_Name='保密督導官'")(0)("Group_Uid")
                ViewState("DepartmentSafeguardsOfficerID") = dtRoleGroup.Select("Group_Name='國防部保防官'")(0)("Group_Uid")

                Dim sPrinter_No As String = PrinterManager(user_id)
                'default Load 未登記的資料 , Roleid = 1系統管理員  2單位管理者 3一般使用者或其他權限
                If Roleid = "2" Then
                    ' 單位管理者只能查自己及其對應下的單位與機器設備的所有人的影印記錄
                    'Return = "0":非管理人員 "-1":SQL Error 其他:管理的影印機機器號碼
                    If sPrinter_No = "0" Then
                    ElseIf sPrinter_No = "-1" Then
                        ErrMsg.Text = "查詢登入者影印機與單位是否有管理對應作業失敗，請重新操作或聯絡資訊人員!!"
                    Else
                        lbORG.Visible = False
                        ddl_ORG_NAME.Visible = False
                        ViewState("ManagerPrinter_No") = sPrinter_No
                    End If
                End If

                If ErrMsg.Text = "" Then
                    PrinterDataRecords_DataRebinding(0, sPrinter_No)
                    LogExportDataRecords(sPrinter_No, "")
                End If
                'Response.Write(Roleid)
            End If
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

    Private Sub PrinterDataRecords_DataRebinding(ByVal SearchStatus As Int16, ByVal ManagerPrinter_No As String)
        If ManagerPrinter_No <> "0" And ManagerPrinter_No <> "-1" And Roleid = "2" Then
            sQuery += " where a.Printer_No in (" + ManagerPrinter_No + ")) as d where d.Status <> 0"
            'ElseIf Roleid = "3" Then
            '    sQuery += ") as d where PAIDNO ='" + user_id + "' and d.Status <> 0" '只予許查詢自己的申請資料
        Else
            sQuery += ") as d where d.Status <> 0" '1:未印(申請完) 2:已印 3:印列失敗 0:清除此sn 
        End If

        Dim ListEmployeeGroup = GetEmployeeRoleGroup(user_id) '使用者所組群組ID
        Dim printerString As String = ""
        'GetPrintersString 取得使用者權限所屬影印機字串,格式: Printer_No1,Printer_No2.....
        '函數第二參數為權限群組 1: 影印機管理者 2: 單位保防官 3: 保密督導官 4: OO部保防官
        If ListEmployeeGroup.Contains(ViewState("DepartmentSafeguardsOfficerID").ToString()) Then 'OO部保防官
            printerString = GetPrintersString(user_id, "4")
        ElseIf ListEmployeeGroup.Contains(ViewState("ConfidentialSupervisoryOfficerID").ToString()) Then '保密督導官
            printerString = GetPrintersString(user_id, "3")
        ElseIf ListEmployeeGroup.Contains(ViewState("UnitSafeguardsOfficerID").ToString()) Then '單位保防官
            printerString = GetPrintersString(user_id, "2")
        ElseIf ListEmployeeGroup.Contains(ViewState("ManagersOfCopierID").ToString()) Then '影印機管理者
            printerString = GetPrintersString(user_id, "1")
        End If
        If Not String.IsNullOrEmpty(printerString) Then '
            sQuery += " AND ( Printer_No IN (" + printerString + ") OR PAIDNO = '" + user_id + "' )"
        ElseIf ListEmployeeGroup.Contains(ViewState("EndUsersID").ToString()) Then '一般使用者
            sQuery += " AND PAIDNO = '" + user_id + "'"
        End If

        'SearchStatus 代查是否為查詢 0:Page_Load 1:ImgSearch_Click
        sqlPrintLog.SelectParameters.Clear()
        Dim bl_Search As Boolean = False
        Dim strSearch As String = String.Empty
        If (SearchStatus = 1) Then
            Dim sSearchString As String = String.Empty
            Dim PrintStartDate As String = Sdate.Text.Trim()
            Dim PrintEndDate As String = Edate.Text.Trim()
            Dim File_Name As String = tbFile_Name.Text.Trim()
            Dim Printer_No As String = tbPrinter_No.Text.Trim()
            Dim Print_Name As String = tb_Print_Name.Text.Trim()
            Dim Printer_Name As String = tbPrinter_Name.Text.Trim()
            Dim PID As String = tb_PID.Text.Trim()
            Dim Security_Status As String = ddl_Security_Status.SelectedValue.ToString()
            Dim Org_Uid As String = ddl_ORG_NAME.SelectedValue.ToString()
            If (PrintStartDate = "" And PrintEndDate <> "") Or (PrintStartDate <> "" And PrintEndDate = "") Then
                ErrMsg.Text = "查詢日期區間必需二者都輸入哦~"
            ElseIf (PrintStartDate <> "" And PrintEndDate <> "") Then
                bl_Search = True
                sSearchString += " and (d.[Print_Date] between @SDate and @EDate)"
                sqlPrintLog.SelectParameters.Add(New Parameter("SDate", DbType.String, PrintStartDate))
                sqlPrintLog.SelectParameters.Add(New Parameter("EDate", DbType.String, PrintEndDate + " 23:59:59"))
                strSearch += " and (d.[Print_Date] between '" + PrintStartDate + "' and '" + PrintEndDate + " 23:59:59')"
            End If

            If (Printer_No.Length > 0) Then
                bl_Search = True
                sSearchString += " and Printer_No like '%' + @Printer_No + '%'"
                sqlPrintLog.SelectParameters.Add(New Parameter("Printer_No", DbType.String, Printer_No))
                strSearch += " and Printer_No like '%" + Printer_No + "%'"
            End If

            If File_Name.Length > 0 Then
                bl_Search = True
                sSearchString += " and File_Name like '%' +@File_Name+'%'"
                sqlPrintLog.SelectParameters.Add(New Parameter("File_Name", DbType.String, File_Name))
                strSearch += " and File_Name like '%" + File_Name + "%'"
            End If

            If Printer_Name.Length > 0 Then
                bl_Search = True
                sSearchString += " and Printer_Name like '%' + @Printer_Name+'%'"
                sqlPrintLog.SelectParameters.Add(New Parameter("Printer_Name", DbType.String, Printer_Name))
                strSearch += " and Printer_Name like '%" + Printer_Name + "%'"
            End If

            If (Security_Status <> "0") Then
                bl_Search = True
                sSearchString += " and d.[Security_Status] = @Security_Status"
                sqlPrintLog.SelectParameters.Add(New Parameter("Security_Status", DbType.String, Security_Status))
                strSearch += " and d.[Security_Status] ='" + Security_Status + "'"
            Else
                bl_Search = True
                strSearch += " and d.[Security_Status] <>'" + Security_Status + "'"
            End If

            If (Print_Name.Length > 0) Then
                bl_Search = True
                sSearchString += " and Print_Name like '%' + @Print_Name + '%'"
                sqlPrintLog.SelectParameters.Add(New Parameter("Print_Name", DbType.String, Print_Name))
                strSearch += " and Print_Name like '%" + Print_Name + "%'"
            End If

            If (PID.Length > 0) Then
                bl_Search = True
                sSearchString += " and PAIDNO like '%' + @employee_id +'%'"
                sqlPrintLog.SelectParameters.Add(New Parameter("employee_id", DbType.String, PID))
                strSearch += " and PAIDNO like '%" + PID + "%'"
            End If

            'Role=2單位管理者就算不是影印機單位管理者也要可以看到同單位下其他同仁的影印記錄
            If (Org_Uid <> "0") Then
                bl_Search = True
                sSearchString += " and d.Org_Uid = @Org_Uid"
                sqlPrintLog.SelectParameters.Add(New Parameter("Org_Uid", DbType.String, Org_Uid))
                strSearch += " and d.Org_Uid ='" + Org_Uid + "'"
            Else
                bl_Search = True
                strSearch += " and d.Org_Uid <> '" + Org_Uid + "'"
            End If

            If (bl_Search) Then
                ViewState("SearchString") = strSearch
                'sQuery += sSearchString
                sQuery += strSearch
            End If
        End If
        sQuery = sQuery + " order by  Log_Guid desc"
        ViewState("QueryString") = sQuery
        sqlPrintLog.SelectCommand = sQuery
        GV_NewLog.DataBind()

    End Sub

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
            Case "2" '單位保防官：可查詢所屬1級單位影印機的所有影印紀錄。
                Dim organizeId1 As String = GetLevelOrganizeId(db, employee_id, "", 1) '成員所屬一級單位ID
                Dim orgTreeId1 As String = GetOrganizeIdTree(db, organizeId1) '一級單位樹狀結構底下所有之單位ID字串 格式: ORG_UID1,ORG_UID2.....
                strSQLCom = "SELECT * FROM P_0803 WHERE ORG_UID IN (" + orgTreeId1 + ")"
            Case "3" '保密督導官：可查詢所屬2級單位影印機的所有影印紀錄。
                Dim organizeId2 As String = GetLevelOrganizeId(db, employee_id, "", 2) '成員所屬二級單位ID
                Dim orgTreeId2 As String = GetOrganizeIdTree(db, organizeId2) '二級單位樹狀結構底下所有之單位ID字串 格式: ORG_UID1,ORG_UID2.....
                strSQLCom = "SELECT * FROM P_0803 WHERE ORG_UID IN (" + orgTreeId2 + ")"
            Case "4" 'OO部保防官：可查詢各單位影印機的所有影印紀錄。
                strSQLCom = "SELECT * FROM P_0803"
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

    '判斷該登入者是不是屬於影印設備管理人員之一 return: "0":非管理人員 "-1":SQL Error 其他:管理的影印機機器號碼
    Private Function PrinterManager(ByVal user_id As String) As String
        Dim sPrint_No As String = "0"
        Dim command As New SqlCommand(String.Empty, New SqlConnection(do_sql.G_conn_string))
        Dim sQuery As String = "select Printer_No from P_0803 with(nolock) where employee_id=@employee_id"

        command.CommandType = CommandType.Text
        command.CommandText = sQuery

        command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = user_id
        Try
            command.Connection.Open()
            Dim drPrinterNo As SqlDataReader = command.ExecuteReader()
            If drPrinterNo.HasRows Then
                While drPrinterNo.Read()
                    sPrint_No += CStr(drPrinterNo("Printer_No")) + ","
                End While
                sPrint_No = sPrint_No.Substring(1, sPrint_No.Length - 2)
            End If
            drPrinterNo.Close()
        Catch ex As Exception
            sPrint_No = "-1"
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try

        Return sPrint_No
    End Function

    Public Function ShowSecurity_Status(ByVal Security_Status As String) As String
        Dim sResult As String = String.Empty
        If Security_Status.Trim().Length <> 0 Then '1:普 2:密 3:機密 4:極機密 5:絕對機密
            Select Case Security_Status
                Case "1"
                    sResult = "普"
                Case "2"
                    sResult = "密"
                Case "3"
                    sResult = "機密"
                Case "4"
                    sResult = "極機密"
                Case "5"
                    sResult = "絕對機密"
                Case Else
                    sResult = "區分未明:" + Security_Status
            End Select
        End If
        Return sResult
    End Function

    Protected Sub ImgDate1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "70px"
        If Sdate.Text.Trim() = "" Then
            Calendar1.SelectedDate = dt.AddDays(-14).Date
        Else
            Calendar1.SelectedDate = Sdate.Text
        End If
        Sdate.Text = Calendar1.SelectedDate.Date
    End Sub

    Protected Sub ImgDate2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "130px"
        If (Edate.Text.Trim() = "") Then
            Calendar2.SelectedDate = dt.Date
        Else
            Calendar2.SelectedDate = Edate.Text
        End If
        Edate.Text = Calendar2.SelectedDate.Date
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Sdate.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        Edate.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False
    End Sub

    Protected Sub btnClose1_Click(sender As Object, e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose2_Click(sender As Object, e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub

    Protected Sub ImgSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        PrinterDataRecords_DataRebinding(1, IIf(IsNothing(ViewState("ManagerPrinter_No")), "0", ViewState("ManagerPrinter_No")))
    End Sub

    Protected Sub ImgClearSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgClearSearch.Click
        tbPrinter_No.Text = ""
        tb_Print_Name.Text = ""
        tbPrinter_Name.Text = ""
        tb_PID.Text = ""
        Sdate.Text = ""
        Edate.Text = ""
        tbFile_Name.Text = ""
        ddl_Security_Status.SelectedValue = "0"
        ddl_ORG_NAME.SelectedValue = "0"
    End Sub

    Protected Sub GV_NewLog_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GV_NewLog.PageIndexChanging
        PrinterDataRecords_DataRebinding(1, IIf(IsNothing(ViewState("ManagerPrinter_No")), "0", ViewState("ManagerPrinter_No")))
    End Sub

    Protected Sub GV_NewLog_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GV_NewLog.RowCommand
        Dim currentCommand As String = e.CommandName
        Select Case currentCommand
            Case "Detail"
                Dim objCurButton As ImageButton = DirectCast(e.CommandSource, ImageButton)
                Dim curGridViewRow As GridViewRow = DirectCast(objCurButton.NamingContainer, GridViewRow)
                Dim HF_Status As HiddenField = CType(curGridViewRow.FindControl("HF_Status"), HiddenField)
                Server.Transfer("MOA08002.aspx?Log_Guid=" & e.CommandArgument.ToString() + "&Status=" + HF_Status.Value)
        End Select
    End Sub
    '必須有此方法，否則RenderControl()方法會出錯 
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Protected Sub ImgExport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgExport.Click
        GV_Export.Visible = True
        '寫入歷程movement=4 另存新檔/列印
        Dim cp As New C_Public
        cp.ActionReWrite(0, user_id, 4, ViewState("QueryString"))
        GV_Export.AllowPaging = False
        LogExportDataRecords(IIf(IsNothing(ViewState("ManagerPrinter_No")), "0", ViewState("ManagerPrinter_No")), ViewState("SearchString"))
        If (GV_Export.Rows.Count = 0) Then
            ErrMsg.Text = "目前查無資料，無法為您匯出檔案哦~"
        Else
            Try
                Dim name As String = "attachment;filename=PrinterLog.xls"
                name = Server.UrlPathEncode(name)
                Dim ExcelHeader As String = "<html><head><meta http-equiv=Content-Type content=text/html; charset=UTF-8><style>td{mso-number-format:\\@}</style></head><body>"
                Dim ExcelTitle As String = "<br/><center><font size=3 color=blue>複(影)印使用管制登記記錄</font></center>"
                Dim ExcelFooter As String = "</body></html>"
                Response.Clear()
                Response.AddHeader("content-disposition", name)
                Response.Charset = ""
                Response.ContentType = "application/excel"

                Dim stringWrite As New StringWriter()
                Dim htmlWrite As New HtmlTextWriter(stringWrite)
                GV_Export.RenderControl(htmlWrite)
                Response.Write(ExcelHeader + ExcelTitle + stringWrite.ToString() + ExcelFooter)
                Response.End()
            Catch ex As Exception
                ErrMsg.Text = "另存Excel失敗：" + ex.Message
            Finally
                GV_Export.AllowPaging = False
            End Try
            GV_Export.Visible = False
        End If
    End Sub

    Public Sub LogExportDataRecords(ByVal ManagerPrinter_No As String, ByVal QueryString As String)
        Dim sPrintQuery As String = "select * from (select distinct Log_Guid,a.Printer_No,PAIDNO,Print_Name,a.Print_Date,isnull(Copy_A3M,0) as Copy_A3M,isnull(Copy_A4M,0) as Copy_A4M,isnull(Copy_A3C,0) as Copy_A3C,isnull(Copy_A4C,0) as Copy_A4C,isnull(SCan,0) as SCan,"
        sPrintQuery += "[Status],EFORMSN,c.Printer_Name,b.[ORG_NAME],a.Org_Uid,a.TU_ID,File_No,[File_Name],Security_Status, "
        sPrintQuery += "PrintTotalCnt,Use_For,Useless from P_08 a with(nolock)  "
        sPrintQuery += "left join [ADMINGROUP] b with(nolock) on a.org_uid = b.org_uid Left join P_0803 c on a.Printer_No=c.Printer_No"
        If ManagerPrinter_No.Length > 0 And Roleid = "2" Then
            sPrintQuery += " where a.Printer_No in (" + ManagerPrinter_No + ")) as d where (Status <> 0)"
        ElseIf Roleid = "3" Then
            sPrintQuery += ") as d where (Status  <> 0) and PAIDNO ='" + user_id + "'"
        Else
            sPrintQuery += ") as d where (Status  <> 0)"
        End If
        sPrintQuery += QueryString
        Dim Logdt As New DataTable("LogData")
        Dim da As SqlDataAdapter = New SqlDataAdapter(sPrintQuery, do_sql.G_conn_string)
        da.Fill(Logdt)
        GV_Export.DataSource = Logdt
        GV_Export.DataBind()
    End Sub

    '組合5種列印張數並回傳顯示給前台的字串
    Public Function ShowPrint(ByVal sA3C As String, ByVal sA4C As String, ByVal sA3M As String, ByVal sA4M As String, ByVal sScan As String) As String
        ShowPrint = "  "
        Dim iPrintTatol As Integer = 0
        If Not aa.isNumeric(sA3C) Or Not aa.isNumeric(sA4C) Or Not aa.isNumeric(sA3M) Or Not aa.isNumeric(sA4M) Or Not aa.isNumeric(sScan) Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('該影(複)印資料表單查詢張數有錯誤，請重新操作或聯絡資訊人員！');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            Dim ayPrintTypeCnt As String() = {sA3C, sA4C, sA3M, sA4M, sScan}
            Dim ayPrintTypeName As String() = {"A3彩色", "A3黑白", "A4彩色", "A4黑白", "掃瞄"}
            For i As Int16 = 0 To ayPrintTypeCnt.Length - 1 Step 1
                If Int16.Parse(ayPrintTypeCnt(i).ToString()) <> 0 Then
                    ShowPrint += ayPrintTypeName(i).ToString() + " : " + ayPrintTypeCnt(i).ToString() + "張 / "
                    iPrintTatol += Int16.Parse(ayPrintTypeCnt(i).ToString())
                End If
            Next i
        End If
        ShowPrint = ShowPrint.Substring(0, ShowPrint.Length - 2)
    End Function

    '查詢送印人級職 
    Public Function ShowTU_Nam(ByVal sTU_ID As String) As String
        ShowTU_Nam = String.Empty
        Dim sTU_Name As String = CF.getTU_Name(sTU_ID, do_sql.G_conn_string)
        If sTU_Name.Length > 0 And sTU_Name <> "error" Then
            ShowTU_Nam = sTU_Name
        End If
    End Function

    '顯示此筆資料是掃瞄或列印 
    Public Function ShowPrintType(ByVal sA3C As String, ByVal sA4C As String, ByVal sA3M As String, ByVal sA4M As String, ByVal sScan As String) As String
        ShowPrintType = ""
        Dim numPrint As Integer
        If Not aa.isNumeric(sA3C) Or Not aa.isNumeric(sA4C) Or Not aa.isNumeric(sA3M) Or Not aa.isNumeric(sA4M) Or Not aa.isNumeric(sScan) Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('該影(複)印資料表單查詢張數有錯誤，請重新操作或聯絡資訊人員！');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            numPrint = Integer.Parse(sA3C) + Integer.Parse(sA4C) + Integer.Parse(sA3M) + Integer.Parse(sA4M)
            If (Integer.Parse(sA3C) > 0 Or Integer.Parse(sA4C) > 0 Or Integer.Parse(sA3M) > 0 Or Integer.Parse(sA4M) > 0 Or Integer.Parse(sScan) > 0) Then
                If (numPrint > 0 And Integer.Parse(sScan) > 0) Then
                    ShowPrintType = "影印&掃瞄"
                ElseIf (numPrint > 0) Then
                    ShowPrintType = "影印"
                ElseIf (Integer.Parse(sScan) > 0) Then
                    ShowPrintType = "掃瞄"
                End If
            End If
        End If
        Return ShowPrintType
    End Function
    Public Function ShowCheckStatus(ByVal sStatus As String) As Boolean
        'ShowCheckStatus = False
        ShowCheckStatus = True
        If Not aa.isNumeric(sStatus) Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('影(複)印資料表單查詢狀態有錯誤，請重新操作或聯絡資訊人員！');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            'If (Integer.Parse(sStatus) = 2 Or Integer.Parse(sStatus) = 4) Then
            '    ShowCheckStatus = True
            'End If
        End If
        Return ShowCheckStatus
    End Function
    Public Function ShowStatus(ByVal Status As String) As String
        Dim sResult As String = String.Empty
        If Status.Trim().Length <> 0 Then '1:未印 2:已印 3:印列失敗 4:補登完畢 5:申請單駁回 6:申請人取消 0:清除此sn
            Select Case Status
                Case "0"
                    sResult = "清除影印申請"
                Case "1"
                    sResult = "申請完成未列印"
                Case "2"
                    sResult = "已列印未回登資料"
                Case "3"
                    sResult = "印列失敗"
                Case "4"
                    sResult = "補登完畢"
                Case "5"
                    sResult = "審核不通過"
                Case "6"
                    sResult = "申請人取消"
                Case Else
                    sResult = "區分未明:" + Status
            End Select
        End If
        Return sResult
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
