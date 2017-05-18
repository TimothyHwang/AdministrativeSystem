Imports System.Data.SqlClient
Imports System.Data
Imports System.IO

Partial Class M_Source_08_MOA08005
    Inherits System.Web.UI.Page
    Dim dt As Date = Now()
    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        'Session被清空回首頁
        If user_id = "" Or org_uid = "" And Not Session("Role") Is Nothing Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")
        Else
            '判斷登入者權限()
            Dim LoginCheck As New C_Public
            If LoginCheck.LoginCheck(user_id, "MOA08005") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA08005.aspx")
                Response.End()
            End If

            Dim Roleid As String = Session("Role").ToString()
            If Roleid <> "1" Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('您無權限進入本頁面觀看功能!!');")
                Response.Write(" window.parent.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            End If
        End If

        '計算目前報修單總數
        Dim iSumRecordCnt As Integer = SumTotalCnt()
        If iSumRecordCnt = -1 Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('查詢資料庫出現異常，請您稍候再試!!');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")
        Else
            ViewState("iSumRecordCnt") = iSumRecordCnt.ToString()
        End If

        If Not Page.IsPostBack And Session("PrintTypeID") Is Nothing Then
            'default先顯示[依登記狀態]的統計表 
            Dim dtErrCause As New DataTable
            Dim sAnyName As String = String.Empty
            dtErrCause = GetAnyData(1, sAnyName, Nothing)
            GVDataBind(1, dtErrCause, sAnyName)
        ElseIf Not Session("PrintTypeID") Is Nothing Then
            Dim iType As Integer = Int16.Parse(Session("PrintTypeID").ToString())
            Session("PrintTypeID") = Nothing
            ddl_QueryType.SelectedIndex = (iType - 1)
            ddl_QueryType_SelectedIndexChanged(Nothing, Nothing)
        End If
        ErrMsg.Text = ""
    End Sub

    Private Function GetAnyData(ByVal iType As Integer, ByRef TypeName As String, ByVal sQueryStr As String) As DataTable
        GetAnyData = New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        Select Case iType
            Case 1
                TypeName = "依影印申請狀態"
                scripts.Append("select [Status],case [status] ") _
                .Append("when 0 then '已清除' when 1 then '申請中未列印' when 2 then '申請已列印未補登' when 3 then '列印失敗' ") _
                .Append("when '4' then '列印補登完成' when '5' then '審核未通過' else '其他' end as nStatusName ") _
                .Append(",count([status]) as Anycnt from P_08 with(nolock)") _
                .Append(sQueryStr) _
                .Append(" group by [status] order by [status] ")
            Case 2
                TypeName = "影印機機型號碼"
                scripts.Append("select Printer_No,count([Printer_No]) as Anycnt from P_08 with(nolock) where Printer_No is not null ") _
                .Append(sQueryStr) _
                .Append(" group by Printer_No order by Printer_No ")
            Case 3
                TypeName = "保密區分"
                scripts.Append("select Security_Status, case Security_Status when 1 then '普' when 2 then '密' ") _
                .Append("when 3 then '機密' when 4 then '極機密' when 5 then '絕對機密' else '未知' end as Security_Status_Name,") _
                .Append("count(Security_Status) as Anycnt from P_08 with(nolock) ") _
                .Append(sQueryStr) _
                .Append(" group by Security_Status order by Security_Status ")
            Case 4
                TypeName = "歷程查詢"
                scripts.Append("select movement,case movement when 1 then '查看' when 2 then '新增' when 3 then '修改' ") _
                 .Append("when 4 then '列印' else '其他' end as movementName ,count([movement]) as Anycnt ") _
                 .Append("from P_0802 with(nolock) ") _
                 .Append(sQueryStr) _
                 .Append(" group by movement order by movement ")
        End Select
        ViewState("TypeName") = TypeName
        ViewState("QueryStr") = scripts.ToString()
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            GetAnyData.Load(dr)

        Catch ex As Exception
            GetAnyData = Nothing

        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

    Public Sub GVDataBind(ByVal iType As Integer, ByVal dt As DataTable, ByVal sAnyName As String)
        Img_Export.Enabled = False
        '先把所有GV顯示先false掉，以避免查詢的結果是無資料時，原GV還SHOW著！
        Dim cControl As Control
        For Each cControl In Me.Form.Controls
            If (TypeOf cControl Is GridView) Then
                cControl.Visible = False
            End If
        Next cControl
        If Not dt Is Nothing Then
            If dt.Rows.Count > 0 Then
                Dim iSumCnt As Integer = 0
                For i As Integer = 0 To dt.Rows.Count - 1
                    iSumCnt += CInt(dt.Rows(i)("Anycnt").ToString())
                Next
                Img_Export.Enabled = True
                ViewState("AnySumCnt") = iSumCnt.ToString()
                ViewState("iType") = iType.ToString()
                Select Case iType
                    Case 1
                        gvPrintStatus.Visible = True
                        gvPrintStatus.DataSource = dt
                        gvPrintStatus.DataBind()
                    Case 2
                        gvPrinterNO.Visible = True
                        gvPrinterNO.DataSource = dt
                        gvPrinterNO.DataBind()
                    Case 3
                        gvSecurityStatus.Visible = True
                        gvSecurityStatus.DataSource = dt
                        gvSecurityStatus.DataBind()
                    Case 4
                        gvHistory.Visible = True
                        gvHistory.DataSource = dt
                        gvHistory.DataBind()
                End Select
            Else
                ErrMsg.Text = "查無" + sAnyName + "的列印記錄!"
            End If
        Else
            ErrMsg.Text = "目前" + sAnyName + ",依您查詢條件時，無任何列印記錄!"
        End If

    End Sub

    Private Function SumTotalCnt() As Integer
        SumTotalCnt = 0
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        scripts.Append("select count(Log_Guid) from P_08 with(nolock)")
        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)

        Try
            command.Connection.Open()
            Dim exobject As New Object
            exobject = command.ExecuteScalar()
            If Not exobject Is Nothing Then
                SumTotalCnt = DirectCast(command.ExecuteScalar(), Integer)
            End If
        Catch ex As Exception
            SumTotalCnt = -1
        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

            command.Dispose()
            command = Nothing

        End Try

    End Function

#Region "日期區間使用之日曆"
    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
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

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
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

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub
#End Region

    Protected Sub ddl_QueryType_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_QueryType.SelectedIndexChanged
        '先將所有的gv都隱藏，再依下拉式選單選到的顯示出來即可
        Dim cControl As Control
        For Each cControl In Me.Form.Controls
            If (TypeOf cControl Is GridView) Then
                cControl.Visible = False
            End If
        Next cControl
        Sdate.Text = ""
        Edate.Text = ""
        Dim iType = Convert.ToInt16(ddl_QueryType.SelectedValue)
        Dim dtAny As New DataTable
        Dim sAnyName As String = String.Empty
        Select Case iType
            Case 1
                dtAny = GetAnyData(iType, sAnyName, Nothing)
                GVDataBind(iType, dtAny, sAnyName)
            Case 2
                dtAny = GetAnyData(iType, sAnyName, Nothing)
                GVDataBind(iType, dtAny, sAnyName)
            Case 3
                dtAny = GetAnyData(iType, sAnyName, Nothing)
                GVDataBind(iType, dtAny, sAnyName)
            Case 4
                dtAny = GetAnyData(iType, sAnyName, Nothing)
                GVDataBind(iType, dtAny, sAnyName)
        End Select
    End Sub

    Protected Sub gvPrintStatus_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvPrintStatus.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("統計件數/總影(複)印件數"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("[" + ViewState("AnySumCnt").ToString() + " ] / [" + ViewState("iSumRecordCnt").ToString() + "]"))
        End If
    End Sub

    Protected Sub Img_Search_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles Img_Search.Click
        If Not ViewState("iType") Is Nothing And Not ViewState("TypeName") Is Nothing Then
            Dim iType As Integer = Convert.ToInt16(ViewState("iType").ToString())
            Dim dtAny As New DataTable
            Dim sAnyName As String = String.Empty
            Dim sDateQuery As String = String.Empty
            If (Sdate.Text.Trim <> "" And Edate.Text.Trim() <> "") Then
                Select Case iType
                    Case 1, 3
                        sDateQuery = " where LogTime between '" + Sdate.Text.Trim() + "' and '" + Edate.Text.Trim() + " 23:59:59'"
                    Case 2
                        sDateQuery = " and LogTime between '" + Sdate.Text.Trim() + "' and '" + Edate.Text.Trim() + " 23:59:59'"
                    Case 4
                        sDateQuery = " where History_Date between '" + Sdate.Text.Trim() + "' and '" + Edate.Text.Trim() + " 23:59:59'"
                End Select
                dtAny = GetAnyData(iType, sAnyName, sDateQuery)
                GVDataBind(iType, dtAny, sAnyName)
            End If
        End If
    End Sub

    Protected Sub gvPrinterNO_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvPrinterNO.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("統計件數/總影(複)印件數"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("[" + ViewState("AnySumCnt").ToString() + " ] / [" + ViewState("iSumRecordCnt").ToString() + "]"))
        End If
    End Sub

    Protected Sub gvSecurityStatus_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvSecurityStatus.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("統計件數/總影(複)印件數"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("[" + ViewState("AnySumCnt").ToString() + " ] / [" + ViewState("iSumRecordCnt").ToString() + "]"))
        End If
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
        '必須有此方法，否則RenderControl()方法會出錯 
    End Sub

    Private Function ExportData(ByRef TypeName As String, ByVal sQueryStr As String) As DataTable
        ExportData = New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        command.CommandText = sQueryStr
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            ExportData.Load(dr)

        Catch ex As Exception
            ExportData = Nothing

        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

    Protected Sub Img_Export_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim iType As Integer = 0
        If Not ViewState("iType") Is Nothing And Not ViewState("TypeName") Is Nothing Then
            iType = Convert.ToInt16(ViewState("iType").ToString())
            Dim GV As New GridView
            Select Case iType
                Case 1
                    GV = CType(FindControl("gvPrintStatus"), GridView)
                Case 2
                    GV = CType(FindControl("gvPrinterNO"), GridView)
                Case 3
                    GV = CType(FindControl("gvSecurityStatus"), GridView)
                Case 4
                    GV = CType(FindControl("gvHistory"), GridView)
            End Select
            Dim sAnyName As String = String.Empty
            Dim dt As New DataTable
            Dim sQuery As String
            If Not ViewState("QueryStr") Is Nothing Then
                sQuery = ViewState("QueryStr").ToString()
                dt = ExportData(sAnyName, sQuery)
            Else
                dt = ExportData(sAnyName, Nothing)
            End If

            GV.DataSource = dt
            GV.AllowPaging = False
            '寫入歷程movement=4 另存新檔/列印
            Dim cp As New C_Public
            cp.ActionReWrite(0, user_id, 4, ViewState("QueryStr"))
            Try
                Dim name As String = "attachment;filename=" + ViewState("TypeName").ToString() + ".xls"
                name = Server.UrlPathEncode(name)
                Dim ExcelHeader As String = "<html><head><meta http-equiv=Content-Type content=text/html; charset=UTF-8><style>td{mso-number-format:\\@}</style></head><body>"
                Dim ExcelTitle As String = "<br/><center><font size=3 color=blue>" + ViewState("TypeName").ToString() + "</font></center>"
                Dim ExcelFooter As String = "</body></html>"
                Response.Clear()
                Response.AddHeader("content-disposition", name)
                Response.Charset = ""
                Response.ContentType = "application/excel"
                GV.DataBind()
                Dim stringWrite As New StringWriter()
                Dim htmlWrite As New HtmlTextWriter(stringWrite)
                GV.Columns(2).Visible = False
                GV.RenderControl(htmlWrite)
                Response.Write(ExcelHeader + ExcelTitle + stringWrite.ToString() + ExcelFooter)
                Response.End()
            Catch ex As Exception
                ErrMsg.Visible = True
                ErrMsg.Text = "另存Excel失敗：" + ex.Message
            Finally
                GV.AllowPaging = False
            End Try

        End If
    End Sub

    Protected Sub gvHistory_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvHistory.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("統計件數/總影(複)印件數"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("[" + ViewState("AnySumCnt").ToString() + " ] / [" + SumTotalHistoryCnt().ToString() + "]"))
        End If
    End Sub

    Private Function SumTotalHistoryCnt() As Integer
        SumTotalHistoryCnt = 0
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        command.CommandText = "select count(History_ID) from P_0802 with(nolock)"
        Try
            command.Connection.Open()
            Dim exobject As New Object
            exobject = command.ExecuteScalar()
            If Not exobject Is Nothing Then
                SumTotalHistoryCnt = DirectCast(command.ExecuteScalar(), Integer)
            End If
        Catch ex As Exception
            SumTotalHistoryCnt = -1
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

#Region "換頁導頁至明細資料頁"
    Protected Sub gvPrintStatus_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvPrintStatus.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA08009.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub


    Protected Sub gvPrinterNO_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvPrinterNO.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA08009.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub

    Protected Sub gvSecurityStatus_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvSecurityStatus.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA08009.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub

    Protected Sub gvHistory_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvHistory.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA08009.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub
#End Region
 
End Class
