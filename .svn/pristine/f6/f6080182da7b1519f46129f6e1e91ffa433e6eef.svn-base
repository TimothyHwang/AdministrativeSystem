Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.Common
Imports System.Web.UI
Imports System.Collections.Generic

Partial Class M_Source_08_MOA08014
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
            Dim strEFORMSN As String = Request.QueryString("EFORMSN")
            Dim strReadOnly As String = Request.QueryString("Read_Only") '表單開啟模式 1:檢視 2:審核
            Dim dt As DataTable = GetPrintRecordsReport(strEFORMSN)
            If dt.Rows.Count > 0 Then '取得呈核單資料
                '若狀態為已批核過 1:待批核 2:核准 3:駁回 或 表單開啟為檢視模式
                If Not dt.Rows(0)("Status").Equals(1) Or strReadOnly.Equals("1") Then
                    btnApprove.Visible = False '則隱藏核准及駁回功能,批核意見隱藏TextBox,顯示Lable
                    btnReject.Visible = False
                    txtOpinion.Visible = False
                    lbOpinion.Visible = True
                    lbOpinion.Text = IIf(dt.Rows(0)("Opinion").Equals(DBNull.Value), "", dt.Rows(0)("Opinion").ToString())
                End If
            End If
            ViewState("EFORMSN") = strEFORMSN
            Dim dtReports As DataTable = GetP_08(strEFORMSN) '取得呈核單所屬影印紀錄資料
            ViewState("DataTablePrintReports") = dtReports
            gvPrintReports.DataSource = dtReports
            gvPrintReports.DataBind()
        End If
    End Sub

    '取得呈核單資料
    Private Function GetPrintRecordsReport(ByVal strEFORMSN As String) As DataTable
        Dim dt As New DataTable
        db.Open()
        Dim strSQL = "SELECT * FROM PrintRecordsReport WHERE EFORMSN = '" + strEFORMSN + "'"
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        GetPrintRecordsReport = dt
    End Function

    '取得呈核單所屬影印紀錄資料
    Private Function GetP_08(ByVal strEFORMSN As String) As DataTable
        Dim dtM As New DataTable
        Dim dt As New DataTable
        db.Open()
        Dim strSQL = "SELECT Log_Guid FROM ReportP_08Mapping WHERE EFORMSN = '" + strEFORMSN + "'"
        Dim dsM As New DataSet
        Dim daM As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        daM.Fill(dsM)
        dtM = dsM.Tables(0)
        Dim strRecords As String = "" '組串該申請單包含影印紀錄資料ID
        For Each drM As DataRow In dtM.Rows
            strRecords = strRecords + drM("Log_Guid").ToString() + ","
        Next
        If Not String.IsNullOrEmpty(strRecords) Then
            strRecords = strRecords.Remove(strRecords.Length - 1) '刪除字串尾逗號
            Dim strP_08 = "SELECT P.*, A.ORG_NAME FROM P_08 AS P LEFT JOIN ADMINGROUP AS A ON P.ORG_UID = A.ORG_UID "
            strP_08 = strP_08 + "WHERE Log_Guid IN (" + strRecords + ")"
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(strP_08, db)
            da.Fill(ds)
            dt = ds.Tables(0)
        End If
        db.Close()
        GetP_08 = dt
    End Function

    Protected Sub gvPrintReports_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvPrintReports.PageIndexChanging
        gvPrintReports.PageIndex = e.NewPageIndex
        Dim dtReports As DataTable = ViewState("DataTablePrintReports")
        gvPrintReports.DataSource = dtReports
        gvPrintReports.DataBind()
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

    Protected Sub gvPrintReports_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvPrintReports.RowCommand
        Dim currentCommand As String = e.CommandName
        Select Case currentCommand
            Case "Detail"
                Dim objCurButton As ImageButton = DirectCast(e.CommandSource, ImageButton)
                Dim curGridViewRow As GridViewRow = DirectCast(objCurButton.NamingContainer, GridViewRow)
                Dim HF_Status As HiddenField = CType(curGridViewRow.FindControl("HF_Status"), HiddenField)
                Server.Transfer("MOA08002.aspx?Log_Guid=" & e.CommandArgument.ToString() + "&Status=" + HF_Status.Value)
        End Select
    End Sub

    Protected Sub btnApprove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        Dim strEFORMSN As String = ViewState("EFORMSN").ToString()
        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "Msg", updatePrintRecordsReportApprove(strEFORMSN, txtOpinion.Text, Session("user_id")), True)
    End Sub

    Protected Sub btnReject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReject.Click
        Dim strEFORMSN As String = ViewState("EFORMSN").ToString()
        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "Msg", updatePrintRecordsReportReject(strEFORMSN, txtOpinion.Text, Session("user_id")), True)
    End Sub

    '核准呈核單
    Private Function updatePrintRecordsReportApprove(ByVal EFORMSN As String, ByVal opinion As String, ByVal verifyBy As String) As String
        Dim resultScript As String = ""
        Dim trans As SqlTransaction = Nothing
        Try
            db.Open()
            Dim strSQL = "SELECT Log_Guid FROM ReportP_08Mapping WHERE EFORMSN = '" + EFORMSN + "'"
            Dim dt As New DataTable
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
            da.Fill(ds)
            dt = ds.Tables(0)
            trans = db.BeginTransaction()
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows '依照對應表內該呈核單所屬Log_Guid Update P_08 影印紀錄資料
                    Dim p08Result As String = UpdateP_08ApprovedByID(db, trans, dr("Log_Guid").ToString(), verifyBy)
                    If Not p08Result.Equals("OK") Then
                        Throw New Exception(p08Result)
                    End If
                Next
                '更新呈核單資料,狀態: 1:待批核 2:核准 3:駁回
                Dim updateSQL = "UPDATE PrintRecordsReport SET Opinion = '" + opinion + "' ,Status = 2 ,verifyDate = GETDATE() ,"
                updateSQL = updateSQL + "verifyBy = '" + Session("user_id") + "' WHERE EFORMSN = '" + EFORMSN + "'"
                Dim updateComm As New SqlCommand(updateSQL, db, trans)
                updateComm.ExecuteNonQuery()
                '更新審核流程資料為核准狀態
                Dim flowSQL = "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'E',comment = '" + opinion + "'"
                flowSQL = flowSQL + " WHERE eformsn = '" + EFORMSN + "' AND gonogo = '?' AND nextstep = '-1'"
                Dim flowComm As New SqlCommand(flowSQL, db, trans)
                flowComm.ExecuteNonQuery()
            Else
                Throw New Exception("Get PrintRecordsReport Fail")
            End If
            trans.Commit()
            resultScript = "alert('資料已核准'); window.dialogArguments.document.location.href = window.dialogArguments.document.location; window.close();"
        Catch ex As Exception
            trans.Rollback()
            resultScript = "alert(""資料核准失敗:" + ex.Message + """);"
        Finally
            db.Close()
        End Try
        updatePrintRecordsReportApprove = resultScript
    End Function

    '駁回呈核單
    Private Function updatePrintRecordsReportReject(ByVal EFORMSN As String, ByVal opinion As String, ByVal verifyBy As String) As String
        Dim resultScript As String = ""
        Dim trans As SqlTransaction = Nothing
        Try
            db.Open()
            Dim strSQL = "SELECT Log_Guid FROM ReportP_08Mapping WHERE EFORMSN = '" + EFORMSN + "'"
            Dim dt As New DataTable
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
            da.Fill(ds)
            dt = ds.Tables(0)
            trans = db.BeginTransaction()
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows '依照對應表內該呈核單所屬Log_Guid Update P_08 影印紀錄資料
                    Dim p08Result As String = ClearP_08VerifyRequesterID(db, trans, dr("Log_Guid").ToString())
                    If Not p08Result.Equals("OK") Then
                        Throw New Exception(p08Result)
                    End If
                Next
                '呈核單狀態: 1:待批核 2:核准 3:駁回
                Dim updateSQL = "UPDATE PrintRecordsReport SET Opinion = '" + opinion + "' ,Status = 3 ,verifyDate = GETDATE() ,"
                updateSQL = updateSQL + "verifyBy = '" + verifyBy + "' WHERE EFORMSN = '" + EFORMSN + "'"
                Dim updateComm As New SqlCommand(updateSQL, db, trans)
                updateComm.ExecuteNonQuery()
                '更新審核流程資料為駁回狀態
                Dim flowSQL = "UPDATE flowctl SET hddate = GETDATE(),gonogo = '0',comment = '" + opinion + "'"
                flowSQL = flowSQL + " WHERE eformsn = '" + EFORMSN + "' AND gonogo = '?' AND nextstep = '-1'"
                Dim flowComm As New SqlCommand(flowSQL, db, trans)
                flowComm.ExecuteNonQuery()
            Else
                Throw New Exception("Get PrintRecordsReport Fail")
            End If
            trans.Commit()
            resultScript = "alert('資料已駁回'); window.dialogArguments.document.location.href = window.dialogArguments.document.location; window.close();"
        Catch ex As Exception
            trans.Rollback()
            resultScript = "alert(""資料駁回失敗:" + ex.Message + """);"
        Finally
            db.Close()
        End Try
        updatePrintRecordsReportReject = resultScript
    End Function

    '更改影印紀錄資料ApprovedByID欄位,即批核者
    Private Function UpdateP_08ApprovedByID(ByRef db As SqlConnection, ByRef trans As SqlTransaction, ByVal log_Guid As String, ByVal ApprovedByID As String) As String
        Dim strResult As String = "OK"
        Try
            Dim strSQL = "UPDATE P_08 SET ApprovedByID = '" + ApprovedByID + "' WHERE Log_Guid = '" + log_Guid + "'"
            Dim comm As New SqlCommand(strSQL, db, trans)
            comm.ExecuteNonQuery()
        Catch ex As Exception
            strResult = ex.Message
        End Try
        UpdateP_08ApprovedByID = strResult
    End Function

    '清除影印紀錄資料VerifyRequesterID欄位,即呈核申請者,亦即將該筆資料設定為未呈核狀態
    Private Function ClearP_08VerifyRequesterID(ByRef db As SqlConnection, ByRef trans As SqlTransaction, ByVal log_Guid As String) As String
        Dim strResult As String = "OK"
        Try
            Dim strSQL = "UPDATE P_08 SET VerifyRequesterID = '' WHERE Log_Guid = '" + log_Guid + "'"
            Dim comm As New SqlCommand(strSQL, db, trans)
            comm.ExecuteNonQuery()
        Catch ex As Exception
            strResult = ex.Message
        End Try
        ClearP_08VerifyRequesterID = strResult
    End Function
End Class
