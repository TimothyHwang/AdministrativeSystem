Imports System.Data.SqlClient
Imports System.Data
Imports WebUtilities.Functions
Partial Class M_Source_08_MOA08004
    Inherits System.Web.UI.Page
    Dim user_id, org_uid, sTU_ID, sLog_Guid, sType As String
    Dim bl_Check As Boolean = False
    Dim iPrintTatol As Int32
    Dim CCFun As New C_CheckFun
    Dim CF As New CFlowSend
    Dim CP As New C_Public
    Dim sql_function As New C_SQLFUN
    Public sTypeID As String = String.Empty
    Public sSearchID As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        sTU_ID = Session("TU_ID")
        'session被清空回首頁
        If user_id = "" Or org_uid = "" Or sTU_ID = "" Then
            bl_Check = True
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            '判斷登入者權限
            If CP.LoginCheck(user_id, "MOA08005") <> "" Then
                CP.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA08005.aspx")
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('您無權限查看此頁面!!');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            End If
            Dim sLog_ID As String = CType(Request.QueryString("Log_ID"), String)
            sTypeID = Request.QueryString("TypeID")
            sSearchID = Request.QueryString("SearchID")
            '判斷導址值是否正確 Type=1新增2編輯 
            If (sLog_ID.Trim().Length < 1 Or (Not CCFun.isNumeric(sLog_ID)) Or CCFun.isNumeric(sTypeID) = False Or CCFun.isNumeric(sSearchID) = False) Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('影(複)印表單序號錯誤，請重新操作！');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            Else

                '查詢截取單一派工單詳細資料
                Dim iLog_ID As Integer = Integer.Parse(sLog_ID)
                Dim sQueryPrintLog As String = String.Empty
                Dim Printerdt As DataTable = DataRecords_DataRebinding(iLog_ID, sQueryPrintLog)
                If Not Printerdt Is Nothing And Printerdt.Rows.Count > 0 Then
                    'Insert歷程入DB Table:[P_0802]-movement=3 查看 
                    CP.ActionReWrite(iLog_ID, user_id, 1, sQueryPrintLog)
                    DV_PrintLog.DataSource = Printerdt
                    DV_PrintLog.DataBind()
                Else
                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('影(複)印表單序號錯誤，請重新操作！！');")
                    Response.Write(" window.parent.location='../../index.aspx';")
                    Response.Write(" </script>")
                End If
            End If
        End If
    End Sub
    '查詢該筆原始列印資料
    Private Function DataRecords_DataRebinding(ByVal Log_ID As Integer, ByRef sQueryPrintLog As String) As DataTable
        DataRecords_DataRebinding = New DataTable("Detail")
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        sQueryPrintLog = "select a.*,b.TU_Name,c.org_name,isnull(Copy_A3M,0) as Copy_A3M,isnull(Copy_A4M,0) as Copy_A4M,isnull(Copy_A3C,0) as Copy_A3C, "
        sQueryPrintLog += "isnull(Copy_A4C,0) as Copy_A4C,isnull(SCan,0) as SCan from P_08 a with(nolock) left join TITLES_U b with(nolock) on a.TU_ID = b.TU_ID "
        sQueryPrintLog += "left join [ADMINGROUP] c with(nolock) on a.org_uid = c.org_uid  "
        sQueryPrintLog += "where Log_Guid =@Log_ID "
        command.CommandType = CommandType.Text
        command.CommandText = sQueryPrintLog
        command.Parameters.Add(New SqlParameter("Log_ID", SqlDbType.Int)).Value = Log_ID
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            DataRecords_DataRebinding.Load(dr)

        Catch ex As Exception
            DataRecords_DataRebinding = Nothing
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

     '組合5種列印張數並回傳顯示給前台的字串
    Public Function ShowPrint(ByVal sA3C As String, ByVal sA4C As String, ByVal sA3M As String, ByVal sA4M As String, ByVal sScan As String) As String
        ShowPrint = String.Empty
        Dim iPrintTatol As Integer = 0
        If Not CCFun.isNumeric(sA3C) Or Not CCFun.isNumeric(sA4C) Or Not CCFun.isNumeric(sA3M) Or Not CCFun.isNumeric(sA4M) Or Not CCFun.isNumeric(sScan) Then
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

        If iPrintTatol = 0 Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('該影(複)印資料表單查詢張數有錯誤(0)，請重新操作或聯絡資訊人員！');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        End If
        ShowPrint = ShowPrint.Substring(0, ShowPrint.Length - 2)
    End Function


    Public Function ShowSecurity_Status(ByVal Security_Status As String, ByVal Security_Guid As String) As String
        Dim sResult As String = String.Empty
        Dim Security As Boolean = False
        If Security_Status.Trim().Length <> 0 Then
            Select Case Security_Status
                Case "1"
                    sResult = "普通件"
                Case "2"
                    Security = True
                    sResult = "<font color='Red'>密件</font>"
                Case "3"
                    Security = True
                    sResult = "<font color='Red'>機密件</font>"
                Case "4"
                    Security = True
                    sResult = "<font color='Red'>極機密件</font>"
                Case "5"
                    Security = True
                    sResult = "<font color='Red'>絕對機密件</font>"
                Case Else
                    sResult = "區分未明:" + Security_Status
            End Select
        End If

        If Security = True Then
            sResult += "&nbsp;<a href = " + "../08/MOA08010.aspx?Security_GuidID=" + Security_Guid + "&Security_Level=" + Security_Status + "&ViewStatusCode=1" + " target = '_blank'>檢視</a>"
        End If
        Return sResult
    End Function

    Protected Sub ibtnPrevious_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ibtnPrevious.Click
        If Not Session("PrintTypeID") Is Nothing Then
            Server.Transfer("MOA08009.aspx?TypeID=" + sTypeID + "&SearchID=" + sSearchID)
        Else
            Dim parameters As New SortedList : With parameters
                .Add("Action", "MOA08003.aspx")
            End With
            Response.Write(Utilities.SubmitFormGeneration(parameters))
            Response.End()
        End If
      
    End Sub
End Class
