Imports System.Web.UI
Imports System.ComponentModel
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports WebUtilities.Functions
Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_04_MOA04103
    Inherits System.Web.UI.Page
    Public printdate As String = DateTime.Now.ToString("yyyy/MM/dd")
    Dim user_id, org_uid As String
    Public EFORMSN As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'Session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將回首頁重新整理');")
            Response.Write("window.close();")
            Response.Write("</script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA04100") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04100.aspx")
                Response.End()
            End If

        End If


        If CType(Request.QueryString("streformsn"), String) Is Nothing Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('您的參數已遺失，請重新操作，感謝您！');")
            Response.Write("window.close();")
            Response.Write("</script>")
        Else
            EFORMSN = CType(Request.QueryString("streformsn"), String)
            If EFORMSN.Length <> 16 Then
                Response.Write("<script language='javascript'>")
                Response.Write("alert('您的報修單號錯誤:<" + EFORMSN + ">，請重新操作，感謝您！');")
                Response.Write("window.close();")
                Response.Write("</script>")
            Else
                lb_EFORMSN.Text = EFORMSN
                Dim snAppStockDate As String = GetnAppStockDate(EFORMSN)
                lb_nAppStockDate.Text = snAppStockDate
                Dim dt As DataTable = GetItemP_0414DetailData(EFORMSN)
                AddRepeaterNullData(dt)
                ExportM04103Word(EFORMSN)
            End If
        End If
    End Sub
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)


    End Sub

    Public Sub AddRepeaterNullData(ByVal dt As DataTable)
        Dim iaddCellCnt As Integer
        If dt Is Nothing Or dt.Rows.Count < 1 Then
            iaddCellCnt = 20
        Else
            iaddCellCnt = 20 - dt.Rows.Count
        End If
        For i As Integer = 1 To iaddCellCnt
            Dim dar As DataRow = dt.NewRow()
            dar("shcode") = ""
            dar("it_name") = ""
            dar("it_unit") = ""
            dar("seatnum") = ""
            dt.Rows.Add(dar)
        Next
        RptitemList.DataSource = dt
        RptitemList.DataBind()
    End Sub

    Public Sub ExportM04103Word(ByVal EFORMSN As String)
        Dim name As String = "attachment;filename=房舍水電派工單與領料單_" + EFORMSN + ".doc"
        name = Server.UrlPathEncode(name)
        Response.Clear()
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Response.AddHeader("content-disposition", name)
        Response.Charset = "UTF-8"
        Response.ContentType = "application/vnd.ms-word"
        Me.EnableViewState = False
        RptitemList.DataBind()
        Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter
        Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)
        Dim hf As New HtmlForm
        Controls.Add(hf)
        hf.Controls.Add(PanelContent)
        hf.RenderControl(htmlWrite)
        Response.Write(stringWrite.ToString)
        Response.End()
    End Sub

    Private Function GetItemP_0414DetailData(ByVal EFORMSN As String) As DataTable
        GetItemP_0414DetailData = New DataTable("Detail")
        Dim scripts As New StringBuilder
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        scripts.Append("select shcode, b.seat_num+'_'+b.seat_name as seatnum,c.it_name,c.it_unit ") _
               .Append("from P_0414 a with(nolock) left join P_0417 b with(nolock) on a.seat_num = b.seat_num ") _
                .Append("left join P_0407 c with(nolock) on substring(a.shcode,1,6) = c.it_code ") _
                .Append("where job_Num =@EFORMSN;")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("EFORMSN", SqlDbType.NVarChar, 16)).Value = EFORMSN
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            GetItemP_0414DetailData.Load(dr)

        Catch ex As Exception
            GetItemP_0414DetailData = Nothing

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

            command.Dispose()
            command = Nothing

        End Try

    End Function

    Private Function GetnAppStockDate(ByVal EFORMSN As String) As String
        GetnAppStockDate = String.Empty
        Dim scripts As New StringBuilder
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        scripts.Append("select nAppStockDate from P_0415 with(nolock)") _
                .Append("  where eformsn=@EFORMSN;")

        command.CommandType = CommandType.Text
        command.CommandText = scripts.ToString()
        scripts.Remove(0, scripts.Length)
        command.Parameters.Add(New SqlParameter("EFORMSN", SqlDbType.NVarChar, 16)).Value = EFORMSN
        Try
            command.Connection.Open()
            Dim obj As New Object
            obj = command.ExecuteScalar()
            If Not obj Is Nothing Then
                GetnAppStockDate = obj.ToString()
            Else
                GetnAppStockDate = "查無此報修單號之申請日期"
            End If

        Catch ex As Exception
            GetnAppStockDate = "查詢報修領料單申請時間失敗"

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

            command.Dispose()
            command = Nothing

        End Try

    End Function
End Class
