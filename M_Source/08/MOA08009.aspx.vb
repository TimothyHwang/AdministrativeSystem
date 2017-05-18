Imports WebUtilities.Functions
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Web.UI.HtmlControls

Partial Class M_Source_08_MOA08009
    Inherits System.Web.UI.Page
    Dim user_id, TypeName As String
    Dim scripts As New StringBuilder
    Dim CP As New C_Public
    Dim CCFun As New C_CheckFun
    Public sTypeID As String = String.Empty
    Public sSearchID As String = String.Empty
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        'Session被清空回首頁
        If user_id = "" Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")
        Else
            '判斷登入者權限
            If CP.LoginCheck(user_id, "MOA08005") <> "" Then
                CP.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA08005.aspx")
                Response.End()
            End If
        End If
        sTypeID = Request.QueryString("TypeID")
        sSearchID = Request.QueryString("SearchID")

        If CCFun.isNumeric(sTypeID) = False Or CCFun.isNumeric(sSearchID) = False Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('參數已遺失，請您重新操作!!');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")
        End If
        If Not IsPostBack Then
            Select Case sTypeID
                Case 1, 2, 3
                    gvDataPrinter_DataRebinding(sTypeID, sSearchID, TypeName)
                Case 4
                    gvHistory_DataRebinding(sTypeID, sSearchID, TypeName)
            End Select

        End If
    End Sub

    Private Function GetPrintLogData(ByVal iType As Integer, ByVal sSearchID As String, ByRef TypeName As String) As DataTable
        GetPrintLogData = New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        Dim sPrintQuery As String = "select distinct Log_Guid,x.*,z.Printer_Name,case security_status "
        sPrintQuery += "when '1' then '普' when '2' then '密' when '3' then '機密' when '4' then '極機密' when '5' then '絕對機密' else convert(nvarchar,security_status,10)"
        sPrintQuery += " end as security_statusName,isnull(y.ORG_Name,'尚未登記') as ORG_Name from"
        sPrintQuery += "(select Log_Guid,Printer_No,PAIDNO,isnull(Print_Name,'尚未列印') as emp_chinese_name,[status],ORG_UID,UPdate_Date,"
        sPrintQuery += " security_status, case when PrintLogDate is null then '尚未列印' else convert(nvarchar,PrintLogDate,120) end as PrintLogDate"
        sPrintQuery += " from P_08 with(nolock)) x "
        sPrintQuery += " left join Admingroup y with(nolock) on x.ORG_UID=y.ORG_UID"
        sPrintQuery += " left join P_0803 z with(nolock) on x.Printer_No=z.Printer_No"
        command.CommandType = CommandType.Text
        Select Case iType
            Case 1
                TypeName = "申請影印狀態" 'searchID = status (分5種) 1:未印 2:已印 3:印列失敗 4:補登完畢 5:不通過 0:清除此sn
                sPrintQuery += " where status = @status"
                command.Parameters.Add("@status", SqlDbType.VarChar, 1).Value = sSearchID
            Case 2
                TypeName = "印表機機型號碼" 'searchID = Printer_No
                sPrintQuery += " where x.Printer_No = @Printer_No"
                command.Parameters.Add("@Printer_No", SqlDbType.VarChar, 6).Value = sSearchID
            Case 3
                TypeName = "保密區分" 'searchID = security_status (分5種)
                sPrintQuery += " where security_status = @security_status"
                command.Parameters.Add("@security_status", SqlDbType.VarChar, 1).Value = sSearchID
        End Select
        'sPrintQuery += " order by Log_Guid"
        Session("PrintTypeID") = iType.ToString()
        ViewState("iType") = iType.ToString()
        ViewState("TypeName") = TypeName
        ViewState("SearchID") = sSearchID
        ViewState("QueryString") = sPrintQuery
        command.CommandText = sPrintQuery
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader
            dr = command.ExecuteReader()
            GetPrintLogData.Load(dr)
            dr.Close()
        Catch ex As Exception
            GetPrintLogData = Nothing
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing

        End Try
    End Function
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
        '必須有此方法，否則RenderControl()方法會出錯 
    End Sub

    Private Function GetBuildingName(ByVal map_code As String, ByRef map_name As String) As Boolean
        Dim bl_execResult = False
        Dim Building As New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        command.CommandText = "SELECT map_name FROM P_0409 with(nolock) where map_code = @map_code"
        command.Parameters.Add("@map_code", SqlDbType.VarChar, 20).Value = map_code
        Try
            command.Connection.Open()
            Dim ob As Object = command.ExecuteScalar()
            If Not ob Is Nothing Then
                map_name = ob.ToString()
                bl_execResult = True
            End If
        Catch ex As Exception
            bl_execResult = False
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
        Return bl_execResult
    End Function

    Private Function GetFacilityName(ByVal it_code As String, ByRef it_name As String) As Boolean
        Dim bl_execResult = False
        Dim Building As New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        command.CommandText = "SELECT it_name FROM P_0407 with(nolock) where it_code = @it_code"
        command.Parameters.Add("@it_code", SqlDbType.VarChar, 6).Value = it_code
        Try
            command.Connection.Open()
            Dim ob As Object = command.ExecuteScalar()
            If Not ob Is Nothing Then
                it_name = ob.ToString()
                bl_execResult = True
            End If
        Catch ex As Exception
            bl_execResult = False
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
        Return bl_execResult
    End Function

    Protected Sub ibtnPrevious_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnPrevious.Click
        Dim parameters As New SortedList : With parameters
            .Add("Action", "MOA08005.aspx")
        End With
        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()
    End Sub

    Public Shared Sub PrintWebControl(ByVal ctrl As Control, ByVal Script As String)
        Dim stringWrite As StringWriter = New StringWriter()
        Dim htmlWrite As System.Web.UI.HtmlTextWriter = New System.Web.UI.HtmlTextWriter(stringWrite)
        If TypeOf ctrl Is WebControl Then
            Dim w As Unit = New Unit(100, UnitType.Percentage)
            CType(ctrl, WebControl).Width = w
        End If
        Dim pg As Page = New Page()
        pg.EnableEventValidation = False
        If Script <> String.Empty Then
            pg.ClientScript.RegisterStartupScript(pg.GetType(), "PrintJavaScript", Script)
        End If
        Dim frm As HtmlForm = New HtmlForm()
        pg.Controls.Add(frm)
        frm.Attributes.Add("runat", "server")
        frm.Controls.Add(ctrl)
        pg.DesignerInitialize()
        pg.RenderControl(htmlWrite)
        Dim strHTML As String = stringWrite.ToString()
        HttpContext.Current.Response.Clear()
        HttpContext.Current.Response.Write(strHTML)
        HttpContext.Current.Response.Write("<script>window.print();</script>")
        HttpContext.Current.Response.End()
    End Sub

    Protected Sub Img_Export_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim iType As Integer = 0
        If Not ViewState("iType") Is Nothing And Not ViewState("TypeName") Is Nothing And Not ViewState("SearchID") Is Nothing Then
            iType = Convert.ToInt16(ViewState("iType").ToString())
            Dim GV As New GridView
            Dim TypeName As String = String.Empty
            Dim dt As New DataTable
            If iType <> 4 Then
                GV = CType(FindControl("gvPrinterData"), GridView)
                dt = GetPrintLogData(Int16.Parse(ViewState("iType").ToString()), ViewState("SearchID").ToString(), TypeName)
                GV.Columns(0).Visible = False
            Else
                GV = CType(FindControl("gvPrintHistory"), GridView)
                dt = GetHistoryLogData(Int16.Parse(ViewState("iType").ToString()), ViewState("SearchID").ToString(), TypeName)
            End If

            GV.DataSource = dt
            GV.AllowPaging = False
            '寫入歷程movement=4 另存新檔/列印
            CP.ActionReWrite(0, user_id, 4, ViewState("QueryString"))
            Try
                Dim name As String = "attachment;filename=" + ViewState("TypeName").ToString() + ".xls"
                name = Server.UrlPathEncode(name)
                Dim ExcelHeader As String = "<html><head><meta http-equiv=Content-Type content=text/html; charset=UTF-8><style>td{mso-number-format:\\@}</style></head><body>"
                Dim ExcelTitle As String = "<br/><center><font size=3 color=blue>依" + ViewState("TypeName").ToString() + "統計明細</font></center>"
                Dim ExcelFooter As String = "</body></html>"
                Response.Clear()
                Response.AddHeader("content-disposition", name)
                Response.Charset = ""
                Response.ContentType = "application/excel"
                GV.DataBind()
                Dim stringWrite As New StringWriter()
                Dim htmlWrite As New HtmlTextWriter(stringWrite)
                GV.Columns(5).Visible = False '人員帳號=身份証字號於匯出excel檔時隱藏不匯出
                GV.RenderControl(htmlWrite)
                Response.Write(ExcelHeader + ExcelTitle + stringWrite.ToString() + ExcelFooter)
                Response.End()
            Catch ex As Exception
                lbEmptyData.Visible = True
                lbEmptyData.Text = "另存Excel失敗：" + ex.Message
            Finally
                GV.AllowPaging = False
            End Try

        End If
    End Sub

    Private Sub gvDataPrinter_DataRebinding(ByVal sTypeID As String, ByVal sSearchID As String, ByRef TypeName As String)
        Dim dt As New DataTable
        dt = GetPrintLogData(Int16.Parse(sTypeID), sSearchID, TypeName)
        lbTypeName.Text = "[ " + TypeName + " ] 單項統計明細資料"
        If Not dt Is Nothing And dt.Rows.Count > 0 Then
            gvPrinterData.Visible = True
            gvPrinterData.DataSource = dt
            gvPrinterData.DataBind()
            gvPrinterData.Enabled = True
            Img_Export.Enabled = True
        End If
    End Sub
    Protected Sub gvPrinterData_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvPrinterData.PageIndexChanging
        gvPrinterData.PageIndex = e.NewPageIndex
        gvDataPrinter_DataRebinding(ViewState("iType").ToString(), ViewState("SearchID").ToString(), TypeName)
    End Sub

    Protected Sub gvPrinterData_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvPrinterData.RowCommand
        Select Case e.CommandName
            Case "GetDetail"
                Server.Transfer("MOA08004.aspx?Log_ID=" + e.CommandArgument.ToString() + "&TypeID=" + sTypeID + "&SearchID=" + sSearchID)
        End Select
    End Sub

    Private Sub gvHistory_DataRebinding(ByVal sTypeID As String, ByVal sSearchID As String, ByRef TypeName As String)
        Dim dt As New DataTable
        dt = GetHistoryLogData(Int16.Parse(sTypeID), sSearchID, TypeName)
        lbTypeName.Text = "[" + TypeName + "] 單項統計明細資料"
        If Not dt Is Nothing And dt.Rows.Count > 0 Then
            gvPrintHistory.Visible = True
            gvPrintHistory.DataSource = dt
            gvPrintHistory.DataBind()
            gvPrintHistory.Enabled = True
            Img_Export.Enabled = True
        End If
    End Sub

    Private Function GetHistoryLogData(ByVal iType As Integer, ByVal sSearchID As String, ByRef TypeName As String) As DataTable
        GetHistoryLogData = New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        Dim sPrintQuery As String = "select x.*,y.ORG_Name from "
        sPrintQuery += "(select History_ID,PrintLog_ID,b.employee_id,b.emp_chinese_name,b.ORG_UID,History_Date, movement,"
        sPrintQuery += "case movement when 1 then '查看' when 2 then '登記/新增' when 3 then '修改' when 4 then '列印/另存' end as movementName"
        sPrintQuery += " from P_0802 a with(nolock) left join employee b on a.employee_id = b.employee_id"
        sPrintQuery += ") as x left join admingroup y with(nolock) on x.ORG_UID = y.ORG_UID"
        TypeName = "操作歷程"
        command.CommandType = CommandType.Text
        sPrintQuery += " where movement = @movement"
        command.Parameters.Add("@movement", SqlDbType.VarChar, 1).Value = sSearchID
        sPrintQuery += " order by History_ID"
        Session("PrintTypeID") = iType.ToString()
        ViewState("TypeName") = TypeName
        ViewState("iType") = iType.ToString()
        ViewState("SearchID") = sSearchID
        ViewState("QueryString") = sPrintQuery
        command.CommandText = sPrintQuery
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader
            dr = command.ExecuteReader()
            GetHistoryLogData.Load(dr)
            dr.Close()
        Catch ex As Exception
            GetHistoryLogData = Nothing
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

    Protected Sub gvPrintHistory_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvPrintHistory.PageIndexChanging
        gvPrintHistory.PageIndex = e.NewPageIndex
        gvHistory_DataRebinding(ViewState("iType").ToString(), ViewState("SearchID").ToString(), TypeName)
    End Sub

    Protected Sub gvPrinterData_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvPrinterData.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim HFPrintLogStatus As HiddenField = CType(e.Row.FindControl("HFPrintLogStatus"), HiddenField)
            If HFPrintLogStatus.Value = "4" Then '4:補登完畢
                Dim ImgPrinterDetail As ImageButton = CType(e.Row.FindControl("ImgPrinterDetail"), ImageButton)
                ImgPrinterDetail.Visible = True
            End If
        End If
    End Sub

    Protected Sub gvPrintHistory_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvPrintHistory.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sType As String = ViewState("iType")
            Dim SearchID As String = ViewState("SearchID")
            If sType = 4 And SearchID = 4 Then
                Dim lbPrintLog_ID As Label = CType(e.Row.FindControl("lbPrintLog_ID"), Label)
                lbPrintLog_ID.Text = "多筆查詢編號"
            End If
        End If
    End Sub
End Class
