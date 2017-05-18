Imports WebUtilities.Functions
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Web.UI.HtmlControls

Partial Class M_Source_04_MOA04104
    Inherits System.Web.UI.Page
    Dim user_id, org_uid, TypeName As String
    Dim scripts As New StringBuilder
    Dim CCFun As New C_CheckFun
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
        Dim sTypeID As String = Request.QueryString("TypeID")
        Dim sSearchID As String = Request.QueryString("SearchID")

        If CCFun.isNumeric(sTypeID) = False Or sSearchID.Length < 1 Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('參數已遺失，請您重新操作!!');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")
        End If
        If Not IsPostBack Then
            gvData_DataRebinding(sTypeID, sSearchID, TypeName)
        End If
    End Sub

    Private Function GetAnyData(ByVal iType As Integer, ByVal sSearchID As String, ByRef TypeName As String) As DataTable
        GetAnyData = New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        Select Case iType
            Case 1
                If sSearchID = "1" Then
                    TypeName = "人為因素"
                ElseIf sSearchID = "2" Then
                    TypeName = "自然因素"
                Else
                    TypeName = "維護查報"
                End If
                'Type = 1:依故障原因統計的單項明細; searchID = ErrCause (分3種，如上三種)
                scripts.Append("select P_Num,EFORMSN,PAUNIT,PANAME,nAPPTIME,nFIXITEM,nCause,nFinalDate,nResult,") _
                        .Append("b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as location") _
                        .Append(" from P_0415 a with(nolock) ") _
                        .Append(" left join P_0404 b on a.nbd_code = b.bd_code") _
                        .Append(" left join P_0406 c on a.nfl_code = c.fl_code") _
                        .Append(" left join P_0411 d on a.nrnum_code = d.rnum_code") _
                        .Append(" where FlowStatus = 4 and nErrCause = '") _
                        .Append(sSearchID) _
                        .Append("'")
            Case 2
                'Type = 2:依建築物統計的單項明細; searchID = map_code (Data from P_0409)
                GetP_0404BuildingName(sSearchID, TypeName)
                scripts.Append("select P_Num,EFORMSN,nAPPTIME,nFIXITEM,nCause,nFinalDate,nResult,") _
                    .Append("b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as location,PAUNIT,PANAME") _
                    .Append(" from P_0415 a with(nolock) ") _
                    .Append(" left join P_0404 b on a.nbd_code = b.bd_code") _
                    .Append(" left join P_0406 c on a.nfl_code = c.fl_code") _
                    .Append(" left join P_0411 d on a.nrnum_code = d.rnum_code") _
                    .Append(" left join (select rnum_code,rnum_name,y.map_code,y.map_name from P_0411 x ") _
                    .Append(" left join P_0409 y on x.map_code = y.map_code )") _
                    .Append(" z on a.nrnum_code = z.rnum_code where flowstatus =4 AND b.bd_code = '") _
                    .Append(sSearchID) _
                    .Append("'")

            Case 3
                'Type = 3:依維修項目統計的單項明細; searchID = nFacilityNo,it_name (Data from P_0415) 
                Dim DateArray As String() = sSearchID.Split(",")
                Dim snFacilityNo As String = DateArray(0)
                TypeName = DateArray(1)
                scripts.Append("select P_Num,EFORMSN,nAPPTIME,EFORMSN,PAUNIT,PANAME,nFIXITEM,nCause,nFinalDate,nResult,") _
                        .Append(" b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as location,nAPPTIME") _
                        .Append(" from P_0415 a with(nolock) ") _
                        .Append(" left join P_0404 b on a.nbd_code = b.bd_code") _
                        .Append(" left join P_0406 c on a.nfl_code = c.fl_code") _
                        .Append(" left join P_0411 d on a.nrnum_code = d.rnum_code") _
                        .Append(" where flowstatus = 4 and nFacilityNo = '") _
                        .Append(snFacilityNo) _
                        .Append("'")
            Case 4
                'Type = 4:依報修類別統計的單項明細; searchID =(2碼)itcode,type1name,type2name 
                Dim DateArray As String() = sSearchID.Split(",")
                Dim s2itcode As String = DateArray(0)
                TypeName = DateArray(1) + " - " + DateArray(2)
                scripts.Append("select P_Num,substring(nFacilityNo,12,6) as it_code,EFORMSN,PAUNIT,PANAME,") _
                    .Append(" b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as location,nAPPTIME,nFIXITEM,nCause,nFinalDate,nResult") _
                    .Append(" from P_0415 a with(nolock) ") _
                    .Append(" left join P_0404 b on a.nbd_code = b.bd_code") _
                    .Append(" left join P_0406 c on a.nfl_code = c.fl_code") _
                    .Append(" left join P_0411 d on a.nrnum_code = d.rnum_code") _
                    .Append(" where flowstatus = 4 ") _
                    .Append(" and substring(substring(nFacilityNo,12,6),1,2) = '") _
                    .Append(s2itcode) _
                    .Append("'")
            Case 5
                'Type = 5:依用料統計的單項明細; searchID = it_code,it_name  (Data from P_0414's substring(a.shcode,1,6))
                Dim DateArray As String() = sSearchID.Split(",")
                Dim sit_code As String = DateArray(0)
                TypeName = DateArray(1)
                scripts.Append("select distinct P_Num,EFORMSN,PAUNIT,PANAME,nAPPTIME,nFIXITEM,nCause,") _
                        .Append(" b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as location,nFinalDate,nResult") _
                        .Append(" from P_0415 a with(nolock) ") _
                        .Append(" left join P_0404 b on a.nbd_code = b.bd_code") _
                        .Append(" left join P_0406 c on a.nfl_code = c.fl_code") _
                        .Append(" left join P_0411 d on a.nrnum_code = d.rnum_code") _
                        .Append(" left join P_0414 e on a.EFORMSN = e.Job_num") _
                        .Append(" where flowstatus = 4 And (substring(e.shcode, 1, 6) = '") _
                        .Append(sit_code) _
                        .Append("')")
            Case 6
                'Type = 6:依element_code的單項明細; searchID = element_code
                Dim DateArray As String() = sSearchID.Split(",")
                Dim selement_code As String = DateArray(0)
                TypeName = DateArray(1)
                scripts.Append("select distinct P_Num,EFORMSN,PAUNIT,PANAME,nAPPTIME,nFIXITEM,nCause,") _
                        .Append(" b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as location,nFinalDate,nResult") _
                        .Append(" from P_0415 a with(nolock) ") _
                        .Append(" left join P_0404 b on a.nbd_code = b.bd_code") _
                        .Append(" left join P_0406 c on a.nfl_code = c.fl_code") _
                        .Append(" left join P_0411 d on a.nrnum_code = d.rnum_code") _
                        .Append(" left join P_0414 e on a.EFORMSN = e.Job_num") _
                        .Append(" where flowstatus = 4 And a.nFacilityNo = '") _
                        .Append(selement_code) _
                        .Append("'")
        End Select
        Session("TypeID") = iType.ToString()
        ViewState("iType") = iType.ToString()
        ViewState("TypeName") = TypeName
        ViewState("SearchID") = sSearchID
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader
            dr = command.ExecuteReader()
            GetAnyData.Load(dr)
            dr.Close()
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

    Private Function GetP_0404BuildingName(ByVal bd_code As String, ByRef bd_name As String) As Boolean
        Dim bl_execResult = False
        Dim Building As New DataTable("DetailData")
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        command.CommandType = CommandType.Text
        command.CommandText = "SELECT bd_name FROM P_0404 with(nolock) where bd_code = @bd_code"
        command.Parameters.Add("@bd_code", SqlDbType.VarChar, 20).Value = bd_code
        Try
            command.Connection.Open()
            Dim ob As Object = command.ExecuteScalar()
            If Not ob Is Nothing Then
                bd_name = ob.ToString()
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
            .Add("Action", "MOA04105.aspx")
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
        If Not ViewState("iType") Is Nothing And Not ViewState("TypeName") Is Nothing Then
            iType = Convert.ToInt16(ViewState("iType").ToString())
            Dim GV As New GridView
            GV = CType(FindControl("gvData"), GridView)
            Dim TypeName As String = String.Empty
            Dim dt As New DataTable
            If Not ViewState("iType") Is Nothing And Not ViewState("SearchID") Is Nothing Then
                dt = GetAnyData(Int16.Parse(ViewState("iType").ToString()), ViewState("SearchID").ToString(), TypeName)
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
                GV.Columns(0).Visible = False
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

    Private Sub gvData_DataRebinding(ByVal sTypeID As String, ByVal sSearchID As String, ByRef TypeName As String)
        Dim dt As New DataTable
        dt = GetAnyData(Int16.Parse(sTypeID), sSearchID, TypeName)
        lbTypeName.Text = "[ " + TypeName + " ] 單項統計明細資料"
        If Not dt Is Nothing And dt.Rows.Count > 0 Then
            gvData.Visible = True
            gvData.DataSource = dt
            gvData.DataBind()
            gvData.Enabled = True
            Img_Export.Enabled = True
        End If
    End Sub
    Protected Sub gvData_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvData.PageIndexChanging
        gvData.PageIndex = e.NewPageIndex
        gvData_DataRebinding(ViewState("iType").ToString(), ViewState("SearchID").ToString(), TypeName)
    End Sub

    Protected Sub gvData_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvData.RowCommand
        Select Case e.CommandName
            Case "GetDetail"
                Server.Transfer("MOA04107.aspx?P_Num=" + e.CommandArgument.ToString())
        End Select
    End Sub
End Class
