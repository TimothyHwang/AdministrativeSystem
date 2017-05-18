Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Web.UI.HtmlControls

Partial Class M_Source_04_MOA04105
    Inherits System.Web.UI.Page
    Dim user_id, org_uid As String
    Public print_file As String = ""
    Public sStardDate As String = String.Empty
    Public sEndDate As String = String.Empty
    Dim dt As Date = Now()
    Dim scripts As New StringBuilder

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        'Session被清空回首頁
        If user_id = "" Or org_uid = "" Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")
        Else
            '判斷登入者權限
            Dim LoginCheck As New C_Public
            If LoginCheck.LoginCheck(user_id, "MOA04105") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04105.aspx")
                Response.End()
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

        If Not Page.IsPostBack And Session("TypeID") Is Nothing Then
            'default先顯示[依故障原因]的統計表 
            Dim dtErrCause As New DataTable
            Dim sAnyName As String = String.Empty
            dtErrCause = GetAnyData(1, sAnyName, Nothing)
            GVDataBind(1, dtErrCause, sAnyName)
        ElseIf Not Session("TypeID") Is Nothing Then
            Dim iType = Int16.Parse(Session("TypeID").ToString())
            Session("TypeID") = Nothing
            ddl_QueryType.SelectedIndex = (iType - 1)
            ddl_QueryType_SelectedIndexChanged(Nothing, Nothing)
        End If
        lbEmptyData.Text = ""
    End Sub
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
                        gvErrCause.Visible = True
                        gvErrCause.DataSource = dt
                        gvErrCause.DataBind()
                    Case 2
                        gvBuilding.Visible = True
                        gvBuilding.DataSource = dt
                        gvBuilding.DataBind()
                    Case 3
                        gvFacility.Visible = True
                        gvFacility.DataSource = dt
                        gvFacility.DataBind()
                    Case 4
                        gvItName.Visible = True
                        gvItName.DataSource = dt
                        gvItName.DataBind()
                    Case 5
                        gvitCodeName.Visible = True
                        gvitCodeName.DataSource = dt
                        gvitCodeName.DataBind()
                    Case 6
                        gvBgFlRome.Visible = True
                        gvBgFlRome.DataSource = dt
                        gvBgFlRome.DataBind()
                End Select
            Else
                lbEmptyData.Text = "查無" + sAnyName + "已完工的報修單!"
            End If
        Else
            lbEmptyData.Text = "目前" + sAnyName + ",依您查詢修件時，無任何已完工的報修單!"
        End If

    End Sub
    Private Function SumTotalCnt() As Integer
        SumTotalCnt = 0
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        scripts.Append("select count(P_NUM) from P_0415 with(nolock)")
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

    Private Function GetAnyData(ByVal iType As Integer, ByRef TypeName As String, ByVal sQueryStr As String) As DataTable
        GetAnyData = New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        Select Case iType
            Case 1
                TypeName = "依故障原因統計"
                sQueryStr = "select nErrCause,case nErrCause "
                sQueryStr += " when '1' then '人為因素' when '2' then '自然因素' when 3 then '維護查報' else '其他' end as nErrCauseName "
                sQueryStr += " ,count(nErrCause) as Anycnt from P_0415 with(nolock)"
                sQueryStr += " where(flowstatus = 4) "
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    sQueryStr += " and convert(varchar,nFinalDate,111) between @StartDate and @EndDate"
                End If
                sQueryStr += "  group by nErrCause order by nErrCause"
                scripts.Append(sQueryStr)
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    If InStr(sQueryStr, "@StartDate") >= 1 And InStr(sQueryStr, "@EndDate") >= 1 Then
                        command.Parameters.Add(New SqlParameter("StartDate", SqlDbType.NVarChar)).Value = sStardDate
                        command.Parameters.Add(New SqlParameter("EndDate", SqlDbType.NVarChar)).Value = sEndDate
                    End If
                End If
            Case 2
                TypeName = "依建築物統計"
                sQueryStr = "select b.bd_code,b.bd_name,count(b.bd_code) as anycnt from P_0415 a left join P_0404 b on a.nbd_code=b.bd_code where flowstatus =4"
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    sQueryStr += " and convert(varchar,nFinalDate,111) between @StartDate and @EndDate"
                End If
                sQueryStr += " group by b.bd_code,b.bd_name "
                'scripts.Append("select b.map_code, b.map_name ,count(b.map_code) as anycnt from P_0415 a with(nolock) left join ( ") _
                '        .Append("select rnum_code,rnum_name,y.map_code,y.map_name from P_0411 x left join P_0409 y on x.map_code = y.map_code ") _
                '        .Append(" ) b on a.nrnum_code = b.rnum_code where flowstatus =4 group by map_code, map_name ")
                scripts.Append(sQueryStr)
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    If InStr(sQueryStr, "@StartDate") >= 1 And InStr(sQueryStr, "@EndDate") >= 1 Then
                        command.Parameters.Add(New SqlParameter("StartDate", SqlDbType.NVarChar)).Value = sStardDate
                        command.Parameters.Add(New SqlParameter("EndDate", SqlDbType.NVarChar)).Value = sEndDate
                    End If
                End If
            Case 3
                TypeName = "依維修項目統計"
                scripts.Append(sQueryStr)
                If InStr(sQueryStr, "@nFacilityNo") >= 1 Then
                    command.Parameters.Add(New SqlParameter("nFacilityNo", SqlDbType.NVarChar, 50)).Value = tbnFacilityNo.Text.Trim()
                End If
                If InStr(sQueryStr, "@it_name") >= 1 Then
                    command.Parameters.Add(New SqlParameter("@it_name", SqlDbType.NVarChar, 255)).Value = tbit_Name.Text.Trim()
                End If
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    If InStr(sQueryStr, "@StartDate") >= 1 And InStr(sQueryStr, "@EndDate") >= 1 Then
                        command.Parameters.Add(New SqlParameter("StartDate", SqlDbType.NVarChar)).Value = sStardDate
                        command.Parameters.Add(New SqlParameter("EndDate", SqlDbType.NVarChar)).Value = sEndDate
                    End If
                End If
            Case 4
                TypeName = "依報修類別統計"
                scripts.Append(sQueryStr)
                If InStr(sQueryStr, "@type1name") >= 1 Then
                    command.Parameters.Add(New SqlParameter("type1name", SqlDbType.NVarChar, 255)).Value = tbtype1name.Text.Trim()
                End If
                If InStr(sQueryStr, "@type2name") >= 1 Then
                    command.Parameters.Add(New SqlParameter("@type2name", SqlDbType.NVarChar, 255)).Value = tbtype2name.Text.Trim()
                End If
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    If InStr(sQueryStr, "@StartDate") >= 1 And InStr(sQueryStr, "@EndDate") >= 1 Then
                        command.Parameters.Add(New SqlParameter("StartDate", SqlDbType.NVarChar)).Value = sStardDate
                        command.Parameters.Add(New SqlParameter("EndDate", SqlDbType.NVarChar)).Value = sEndDate
                    End If
                End If
            Case 5
                TypeName = "依用料統計"
                scripts.Append(sQueryStr)
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    If InStr(sQueryStr, "@StartDate") >= 1 And InStr(sQueryStr, "@EndDate") >= 1 Then
                        command.Parameters.Add(New SqlParameter("StartDate", SqlDbType.NVarChar)).Value = sStardDate
                        command.Parameters.Add(New SqlParameter("EndDate", SqlDbType.NVarChar)).Value = sEndDate
                    End If
                End If
            Case 6
                TypeName = "依設備分佈統計"
                scripts.Append(sQueryStr)
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    If InStr(sQueryStr, "@StartDate") >= 1 And InStr(sQueryStr, "@EndDate") >= 1 Then
                        command.Parameters.Add(New SqlParameter("StartDate", SqlDbType.NVarChar)).Value = sStardDate
                        command.Parameters.Add(New SqlParameter("EndDate", SqlDbType.NVarChar)).Value = sEndDate
                    End If
                End If
        End Select
        ViewState("TypeName") = TypeName

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
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
        '必須有此方法，否則RenderControl()方法會出錯 
    End Sub

    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim iType As Integer = 0
        If Not ViewState("iType") Is Nothing And Not ViewState("TypeName") Is Nothing Then
            iType = Convert.ToInt16(ViewState("iType").ToString())
            Dim GV As New GridView
            Select Case iType
                Case 1
                    GV = CType(FindControl("gvErrCause"), GridView)
                Case 2
                    GV = CType(FindControl("gvBuilding"), GridView)
                Case 3
                    GV = CType(FindControl("gvFacility"), GridView)
                Case 4
                    GV = CType(FindControl("gvItName"), GridView)
                Case 5
                    GV = CType(FindControl("gvitCodeName"), GridView)
                Case 6
                    GV = CType(FindControl("gvBgFlRome"), GridView)
            End Select
            Dim sAnyName As String = String.Empty
            Dim dt As New DataTable
            Dim sQuery As String
            If Not ViewState("QueryStr") Is Nothing Then
                sQuery = ViewState("QueryStr").ToString()
                dt = GetAnyData(iType, sAnyName, sQuery)
            Else
                dt = GetAnyData(iType, sAnyName, Nothing)
            End If

            GV.DataSource = dt
            GV.AllowPaging = False

            Try
                Dim name As String = "attachment;filename=" + ViewState("TypeName").ToString() + ".xls"
                name = Server.UrlPathEncode(name)
                Dim ExcelHeader As String = "<html><head><meta http-equiv=Content-Type content=text/html; charset=UTF-8></head><body>"
                Dim ExcelTitle As String = "<br/><center><font size=3 color=blue>" + ViewState("TypeName").ToString() + "</font></center>"
                Dim ExcelFooter As String = "</body></html>"
                Response.Clear()
                Response.AddHeader("content-disposition", name)
                Response.Charset = ""
                Response.ContentType = "application/excel"
                GV.DataBind()
                Dim stringWrite As New StringWriter()
                Dim htmlWrite As New HtmlTextWriter(stringWrite)
                GV.Columns(3).Visible = False
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

    Protected Sub gvErrCause_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvErrCause.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("總計"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("統計數/總單數"))
            e.Row.Cells(2).Controls.Add(New LiteralControl("　" + ViewState("AnySumCnt").ToString() + " / " + ViewState("iSumRecordCnt").ToString()))
        End If
    End Sub

    Protected Sub ddl_QueryType_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_QueryType.SelectedIndexChanged
        '先將所有的gv都隱藏，再依下拉式選單選到的顯示出來即可
        Dim cControl As Control
        For Each cControl In Me.Form.Controls
            If (TypeOf cControl Is GridView) Then
                cControl.Visible = False
            End If
        Next cControl
        dvFacifity.Visible = False
        dvitCode.Visible = False
        nStartDATE1.Text = ""
        nStartDATE2.Text = ""
        btQueryYear.Visible = True
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
                btQueryYear.Visible = False
                dvFacifity.Visible = True
                btQueryFacifity_Click(Nothing, Nothing)
            Case 4
                btQueryYear.Visible = False
                dvitCode.Visible = True
                btQueryitCode_Click(Nothing, Nothing)
            Case 5
                btQueryYear_Click(Nothing, Nothing)
            Case 6
                btQueryYear_Click(Nothing, Nothing)
        End Select
    End Sub

    Protected Sub gvBuilding_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvBuilding.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("總計"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("統計數/總單數"))
            e.Row.Cells(2).Controls.Add(New LiteralControl("　" + ViewState("AnySumCnt").ToString() + " / " + ViewState("iSumRecordCnt").ToString()))
        End If
    End Sub

    Protected Sub btQueryFacifity_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btQueryFacifity.Click
        Dim bl_DTError As Boolean = False
        If nStartDATE1.Text.Trim() <> "" And nStartDATE2.Text.Trim() <> "" Then
            Try
                Dim dt_StartDT As Date = Date.Parse(nStartDATE1.Text)
                Dim dt_EndDT As Date = Date.Parse(nStartDATE2.Text)
                If (dt_EndDT < dt_StartDT) Then
                    bl_DTError = True
                Else
                    sStardDate = dt_StartDT.ToString("yyyy/MM/dd")
                    sEndDate = dt_EndDT.ToString("yyyy/MM/dd")
                End If

            Catch ex As Exception
                bl_DTError = True
            End Try

            If (bl_DTError) Then
                lbDTMsg.Text = "您查詢的日期有誤!!"
                lbDTMsg.ForeColor = Drawing.Color.Red
                Return
            Else
                lbDTMsg.Text = ""
            End If
        End If

        gvFacility.Visible = False
        Dim sQuery As String = "select d.nFacilityNo, it_name , count(d.nFacilityNo) as Anycnt from P_0407 c right join "
        sQuery += " (select it_code ,nFacilityNo,eformsn"
        sQuery += " from P_0415 a left join P_0405 b with(nolock) on a.nFacilityNo = b.element_code"
        sQuery += " where flowstatus = 4 "
        If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
            sQuery += " and convert(varchar,nFinalDate,111) between @StartDate and @EndDate"
        End If
        sQuery += ") d on c.it_code = d.it_code where (1=1)"

        If tbnFacilityNo.Text.Trim() <> "" Then
            sQuery += " and nFacilityNo like ('%' +@nFacilityNo + '%')"
        End If
        If tbit_Name.Text.Trim() <> "" Then
            sQuery += " and it_name like ('%' +@it_name + '%')"
        End If

        sQuery += " group by nFacilityNo, it_name "
        ViewState("QueryStr") = sQuery
        Dim dtAny As New DataTable
        Dim sAnyName As String = String.Empty
        dtAny = GetAnyData(3, sAnyName, sQuery)
        GVDataBind(3, dtAny, sAnyName)
    End Sub

    Protected Sub gvFacility_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvFacility.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("總計"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("統計數/總單數"))
            e.Row.Cells(2).Controls.Add(New LiteralControl("　" + ViewState("AnySumCnt").ToString() + " / " + ViewState("iSumRecordCnt").ToString()))
        End If
    End Sub

    Protected Sub btQueryitCode_Click(sender As Object, e As System.EventArgs) Handles btQueryitCode.Click
        gvItName.Visible = False
        Dim bl_DTError As Boolean = False
        If nStartDATE1.Text.Trim() <> "" And nStartDATE2.Text.Trim() <> "" Then
            Try
                Dim dt_StartDT As Date = Date.Parse(nStartDATE1.Text)
                Dim dt_EndDT As Date = Date.Parse(nStartDATE2.Text)
                If (dt_EndDT < dt_StartDT) Then
                    bl_DTError = True
                Else
                    sStardDate = dt_StartDT.ToString("yyyy/MM/dd")
                    sEndDate = dt_EndDT.ToString("yyyy/MM/dd")
                End If

            Catch ex As Exception
                bl_DTError = True
            End Try

            If (bl_DTError) Then
                lbDTMsg.Text = "您查詢的日期有誤!!"
                lbDTMsg.ForeColor = Drawing.Color.Red
                Return
            Else
                lbDTMsg.Text = ""
            End If
        End If

        Dim sQuery As String = "select type1name,d.it_name as type2name,itcode,count(itcode) as Anycnt from View_P0407_ItType2 d right join "
        sQuery += " (select it_name as type1name,itcode from View_P0407_ItType1 c right join "
        sQuery += " (select substring(it_code,1,2) as itcode from P_0405 a right join "
        sQuery += " (select nFacilityNo from P_0415 with(nolock) where flowstatus = 4"
        If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
            sQuery += " and (convert(varchar,nFinalDate,111) between @StartDate and @EndDate)"
        End If
        sQuery += ") b  on a.element_code = b.nFacilityNo"
        sQuery += " ) as maindata on substring(c.it_code,1,1) = substring(maindata.itcode,1,1)"
        sQuery += " ) e on e.itcode = substring(d.it_code,1,2)"
        sQuery += " where (1=1)"

        If tbtype1name.Text.Trim() <> "" Then
            sQuery += " and type1name like ('%' +@type1name + '%')"
        End If
        If tbtype2name.Text.Trim() <> "" Then
            sQuery += " and d.it_name like ('%' +@type2name + '%')"
        End If
      
        sQuery += " group by type1name,d.it_name,itcode "
        ViewState("QueryStr") = sQuery
        Dim dtAny As New DataTable
        Dim sAnyName As String = String.Empty
        dtAny = GetAnyData(4, sAnyName, sQuery)
        GVDataBind(4, dtAny, sAnyName)
    End Sub

    Protected Sub gvItName_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvItName.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("總計"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("統計數/總單數"))
            e.Row.Cells(2).Controls.Add(New LiteralControl("　" + ViewState("AnySumCnt").ToString() + " / " + ViewState("iSumRecordCnt").ToString()))
        End If
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

    Protected Sub gvitCodeName_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvitCodeName.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("總計"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("統計數/總單數"))
            e.Row.Cells(2).Controls.Add(New LiteralControl("　" + ViewState("AnySumCnt").ToString() + " / " + ViewState("iSumRecordCnt").ToString()))
        End If
    End Sub

    Protected Sub gvBgFlRome_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvBgFlRome.RowCreated
        If e.Row.RowType = DataControlRowType.Footer Then
            'Cells[0]表示第一個欄位,以此類推
            e.Row.Cells(0).Controls.Add(New LiteralControl("總計"))
            e.Row.Cells(1).Controls.Add(New LiteralControl("統計數/總單數"))
            e.Row.Cells(2).Controls.Add(New LiteralControl("　" + ViewState("AnySumCnt").ToString() + " / " + ViewState("iSumRecordCnt").ToString()))
        End If
    End Sub

#Region "換頁導頁至明細資料頁"
    Protected Sub gvErrCause_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvErrCause.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA04104.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub
    Protected Sub gvBuilding_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvBuilding.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA04104.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub
    Protected Sub gvFacility_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvFacility.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA04104.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub
    Protected Sub gvItName_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvItName.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA04104.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
        End Select
    End Sub
    Protected Sub gvBgFlRome_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvBgFlRome.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA04104.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
                'SearchID= "c1c2e0," 以相同物料查詢其所有報修單，所以有可能統計頁查詢出1筆，但按下明細卻有多筆報修單資料
        End Select
    End Sub
    Protected Sub gvitCodeName_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvitCodeName.RowCommand
        Select Case e.CommandName
            Case "AnyDetail"
                Server.Transfer("MOA04104.aspx?TypeID=" + ViewState("iType").ToString() + "&SearchID=" + e.CommandArgument.ToString())
                '以 "c1f1b5,主體大樓地下二樓A2梯" 以相同物料c1f1b5查詢其所有報修單，所以有可能統計頁查詢出1筆，但按下明細卻有多筆報修單資料
        End Select
    End Sub
#End Region
   
    Protected Sub btQueryYear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btQueryYear.Click
        Dim iType = Convert.ToInt16(ddl_QueryType.SelectedValue)
        Dim bl_DTError As Boolean = False
        If iType = 1 Or iType = 2 Or iType = 5 Or iType = 6 Then

            If nStartDATE1.Text.Trim() <> "" And nStartDATE2.Text.Trim() <> "" Then
                Try
                    Dim dt_StartDT As Date = Date.Parse(nStartDATE1.Text)
                    Dim dt_EndDT As Date = Date.Parse(nStartDATE2.Text)
                    If (dt_EndDT < dt_StartDT) Then
                        bl_DTError = True
                    Else
                        sStardDate = dt_StartDT.ToString("yyyy/MM/dd")
                        sEndDate = dt_EndDT.ToString("yyyy/MM/dd")
                    End If

                Catch ex As Exception
                    bl_DTError = True
                End Try

                If (bl_DTError) Then
                    lbDTMsg.Text = "您查詢的日期有誤!!"
                    lbDTMsg.ForeColor = Drawing.Color.Red
                    Return
                Else
                    lbDTMsg.Text = ""
                End If
            End If
            Dim dtAny As New DataTable
            Dim sAnyName As String = String.Empty

            If iType = 1 Or iType = 2 Then
                dtAny = GetAnyData(iType, sAnyName, Nothing)
                GVDataBind(iType, dtAny, sAnyName)
            End If
            If iType = 5 Then
                gvItName.Visible = False
                Dim sQuery As String = "select c.it_code,c.it_name,count(it_code) as Anycnt from P_0407 c with(nolock) right join"
                sQuery += " (select EFORMSN,nFinalDate,substring(b.shcode,1,6) as P0414Code,UseDate"
                sQuery += " from P_0415 a with(nolock) left join P_0414 b with(nolock) on a.EFORMSN = b.Job_num"
                sQuery += " where flowstatus = 4 and b.UseCheck = 2) d on c.it_code = d.P0414Code"
                sQuery += " where (it_code Is Not null) and substring(it_code,3,3) != '000'" '有些報修單沒有使用到領料用料

                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    sQuery += " and convert(varchar,nFinalDate,111) between @StartDate and @EndDate"
                End If
                sQuery += " group by it_code,it_name"
                ViewState("QueryStr") = sQuery

                dtAny = GetAnyData(5, sAnyName, sQuery)
                GVDataBind(5, dtAny, sAnyName)
            End If
            If iType = 6 Then
                gvItName.Visible = False
                Dim sQuery As String = "select element_code,d.bd_name,e.fl_name,f.rnum_name,d.bd_name+e.fl_name+f.rnum_name as TypeName,count(element_code) as Anycnt from"
                sQuery += " (select element_code,bd_code,fl_code,rnum_code from P_0405 a"
                sQuery += " right join P_0415 b on a.element_code = b.nFacilityNo where b.FlowStatus = 4"
                If (sStardDate <> String.Empty And sEndDate <> String.Empty) Then
                    sQuery += " and convert(varchar,b.nFinalDate,111) between @StartDate and @EndDate"
                End If
                sQuery += ") c"
                sQuery += " left join P_0404 d with(nolock) on c.bd_code = d.bd_code"
                sQuery += " left join P_0406 e with(nolock) on c.fl_code = e.fl_code"
                sQuery += " left join P_0410 f with(nolock) on c.rnum_code = f.rnum_code"
              
                sQuery += " group by element_code,d.bd_name,e.fl_name,f.rnum_name"
                ViewState("QueryStr") = sQuery

                dtAny = GetAnyData(6, sAnyName, sQuery)
                GVDataBind(6, dtAny, sAnyName)
            End If
        End If
    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "170px"
        If nStartDATE1.Text.Trim() = "" Then
            Calendar1.SelectedDate = dt.AddDays(-14).Date
        Else
            Calendar1.SelectedDate = nStartDATE1.Text
        End If
        nStartDATE1.Text = Calendar1.SelectedDate.Date
    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "230px"
        If (nStartDATE2.Text.Trim() = "") Then
            Calendar2.SelectedDate = dt.Date
        Else
            Calendar2.SelectedDate = nStartDATE2.Text
        End If
        nStartDATE2.Text = Calendar2.SelectedDate.Date
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        nStartDATE1.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        nStartDATE2.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False
    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub
End Class
