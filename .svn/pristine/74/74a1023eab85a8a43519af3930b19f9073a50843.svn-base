Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Partial Class M_Source_08_MOA08007
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Public aa As New C_CheckFun
    Dim CP As New C_Public
    Dim dt As Date = Now()
    Dim user_id, org_uid, Roleid As String
    Public sQuery As String = "select Guid_ID,Top1unitName,Subject,Security_No,Security_Level,Security_Type,Purpose,Print_UserID,isnull(Purpose_Other,'') as Purpose_Other,emp_chinese_name,Security_DateTime,Security_Status,PAIDNO from P_0804 with(nolock) left join employee with(nolock) on P_0804.Print_UserID= employee.employee_id where (1=1)"
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
        End If

        If Not IsPostBack Then
            'default Load 全部的資料 , Roleid = 1系統管理員  2單位管理者 3一般使用者或其他權限
            'Roleid = "1"=系統管理員即所有的都可以看，其他則只能看自己的申請登記的機密資訊申請單
            If Roleid <> "1" Then '一般使用者只能看到自己的list，無法查詢其他人的機密資訊申請記錄
                lbUserName.Visible = False
                tbUserName.Visible = False
            End If

            PrinterDataRecords_DataRebinding(0)
        End If
    End Sub
    Private Sub PrinterDataRecords_DataRebinding(ByVal SearchStatus As Int16)
        sqlSecurityLog.SelectParameters.Clear()
        Dim bl_Search As Boolean = False
        'SearchStatus 代查是否為查詢 0:Page_Load 1:ImgSearch_Click
        If (SearchStatus = 1) Then
            Dim sSearchString As String = String.Empty
            Dim PrintStartDate As String = Sdate.Text.Trim()
            Dim PrintEndDate As String = Edate.Text.Trim()
            Dim Print_Name As String = tbUserName.Text.Trim()
            If (Print_Name <> "") Then
                bl_Search = True
                sSearchString += " and emp_chinese_name = @Print_Name"
                sqlSecurityLog.SelectParameters.Add(New Parameter("Print_Name", DbType.String, Print_Name))
            End If

            Dim Security_Status As String = ddl_Security_Status.SelectedValue.ToString()
            If (PrintStartDate = "" And PrintEndDate <> "") Or (PrintStartDate <> "" And PrintEndDate = "") Then
                ErrMsg.Text = "查詢日期區間必需二者都輸入哦~"
            ElseIf (PrintStartDate <> "" And PrintEndDate <> "") Then
                bl_Search = True
                sSearchString += " and(Security_DateTime between @SDate and @EDate)"
                sqlSecurityLog.SelectParameters.Add(New Parameter("SDate", DbType.String, PrintStartDate))
                sqlSecurityLog.SelectParameters.Add(New Parameter("EDate", DbType.String, PrintEndDate + " 23:59:59"))
            End If

            If (Security_Status <> "-1") Then
                bl_Search = True
                sSearchString += " and Security_Status = @Security_Status"
                sqlSecurityLog.SelectParameters.Add(New Parameter("Security_Status", DbType.String, Security_Status))
            End If

            If (bl_Search) Then
                sQuery += sSearchString
            End If
        End If

        If Roleid <> "1" Then '只能看到單一帳號(個人)的機密資訊申請記錄資料
            sQuery += " and Print_UserID = @Print_UserID"
            sqlSecurityLog.SelectParameters.Add(New Parameter("Print_UserID", DbType.String, user_id))
        End If

        sQuery += " order by Security_DateTime Desc"
        ViewState("QueryString") = sQuery
        sqlSecurityLog.SelectCommand = sQuery
        GV_Security.DataBind()
    End Sub

    '將機密等級:[2、3、4、5]的狀態代碼轉為中文字於前台顯示
    Public Function ShowSecurity_Level(ByVal nSecurity_Level As String) As String
        Dim sResult As String = String.Empty
        If nSecurity_Level.Trim().Length <> 0 Then
            Select Case nSecurity_Level
                Case "2"
                    sResult = "密"
                Case "3"
                    sResult = "機密"
                Case "4"
                    sResult = "極機密"
                Case "5"
                    sResult = "絕對機密"
                Case Else
                    sResult = "機密等級未明:" + nSecurity_Level
            End Select
        End If
        Return sResult
    End Function

    '將機密屬性:[1~7]的狀態代碼轉為中文字於前台顯示
    Public Function ShowSecurity_Type(ByVal nSecurity_Level As String, ByVal nSecurity_Type As String) As String
        Dim sResult As String = String.Empty
        If nSecurity_Type.Trim().Length <> 0 And nSecurity_Level.Trim().Length <> 0 Then

            If (nSecurity_Level = "2") Then 'nSecurity_Level = 2 [密]時屬性不可自選一定為[一般公務機密]
                sResult = "一般公務機密"
            Else
                Select Case nSecurity_Type
                    Case "1"
                        sResult = "國家機密"
                    Case "2"
                        sResult = "軍事機密"
                    Case "3"
                        sResult = "國防秘密"
                    Case "4"
                        sResult = "國家機密亦屬軍事機密"
                    Case "5"
                        sResult = "國家機密亦屬國防秘密"
                    Case Else
                        sResult = "機密屬性未明:" + nSecurity_Type
                End Select
            End If
        End If
        Return sResult
    End Function

    '將申請單狀態:[0 or 1 or 2 or 3]的狀態代碼轉為中文字於前台顯示
    Public Function ShowSecurity_Status(ByVal nSecurity_Guid As String, ByVal nSecurity_Status As String) As String
        Dim sResult As String = String.Empty
        If nSecurity_Status.Trim().Length <> 0 Then
            Select Case nSecurity_Status
                Case "0"
                    If getSecurity_Status(nSecurity_Guid) = "1" Then
                        sResult = "審核中"
                    Else
                        sResult = "未送審"
                    End If
                Case "1"
                    sResult = "審核通過"
                Case "2"
                    sResult = "審核不通過"
                Case "3"
                    sResult = "申請人取消"
                Case Else
                    sResult = "申請單狀態未明:" + nSecurity_Status
            End Select
        End If
        Return sResult
    End Function
    '傳回機密件是否已申請影印
    Private Function getSecurity_Status(ByVal nSecurity_Guid As String) As String
        Dim connstr As String
        Dim dt As DataTable = New DataTable("securitydt")
        connstr = do_sql.G_conn_string
        Try
            Using conn As New SqlConnection(connstr)
                Dim da As New SqlDataAdapter("select * from P_08 a join P_0804 b on b.Guid_ID=a.Security_Guid where b.Guid_ID=" + nSecurity_Guid, conn)
                da.Fill(dt)
            End Using
        Catch ex As Exception
            dt = Nothing
        End Try
        If Not dt Is Nothing And dt.Rows.Count > 0 Then
            Return "1" '已申請
        Else
            Return "0"
        End If
    End Function


    '將申請用途:[1~7]的狀態代碼轉為中文字於前台顯示,7:其他會含文字敘述
    Public Function ShowPurpose(ByVal nPurpose As String, ByVal nPurpose_Other As String) As String
        Dim sResult As String = String.Empty
        If nPurpose.Trim().Length <> 0 Then
            Select Case nPurpose
                Case "1"
                    sResult = "呈閱"
                Case "2"
                    sResult = "分會、辦"
                Case "3"
                    sResult = "作業用"
                Case "4"
                    sResult = "歸檔"
                Case "5"
                    sResult = "隨文分發"
                Case "6"
                    sResult = "會議分發"
                Case "7"
                    sResult = "其他：" + nPurpose_Other
                Case Else
                    sResult = "申請單狀態未明:" + nPurpose
            End Select
        End If
        Return sResult
    End Function

    Protected Sub GV_Security_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GV_Security.PageIndexChanging
        PrinterDataRecords_DataRebinding(1)
    End Sub
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

    Protected Sub ImBt_Clear_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImBt_Clear.Click
        tbUserName.Text = ""
        Sdate.Text = ""
        Edate.Text = ""
        ddl_Security_Status.SelectedValue = "-1"
    End Sub

    Protected Sub ImgSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        PrinterDataRecords_DataRebinding(1)
    End Sub

    Protected Sub GV_Security_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GV_Security.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim _LinkBtnDetail As HyperLink = e.Row.FindControl("LinkBtnDetail")
            Dim _hid_guid As HiddenField = e.Row.FindControl("hid_guid")
            Dim _hid_Security_Level As HiddenField = e.Row.FindControl("hid_Security_Level")
            Dim _hid_PAIDNO As HiddenField = e.Row.FindControl("hid_PAIDNO")
            Dim _LinkBtnDel As ImageButton = e.Row.FindControl("LinkBtnDel")
            Dim connstr As String
            _LinkBtnDel.Visible = False

            If _hid_guid.Value <> Nothing Then
                Dim dt As DataTable = New DataTable("P_0804")
                connstr = do_sql.G_conn_string
                Try
                    Using conn As New SqlConnection(connstr)
                        Dim da As New SqlDataAdapter("select Guid_ID from P_0804 where PAIDNO='" + user_id + "' and Guid_ID=" + Trim(_hid_guid.Value), conn)                        '
                        da.Fill(dt)
                    End Using
                Catch ex As Exception
                    dt = Nothing
                End Try
                If Not dt Is Nothing And dt.Rows.Count > 0 Then
                    If getSecurity_Status(Trim(_hid_guid.Value)) = "0" Then
                        _LinkBtnDel.Visible = True
                    End If
                End If
                '_LinkBtnDetail.NavigateUrl = "../08/MOA08010.aspx?Security_GuidID=" + _hid_guid.Value + "&Security_Level=" + _hid_Security_Level.Value
            End If
        End If
    End Sub


    Protected Sub GV_Security_RowDeleting(sender As Object, e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GV_Security.RowDeleting
        PrinterDataRecords_DataRebinding(1)
    End Sub
End Class
