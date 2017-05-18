Imports System.Data.SqlClient
Imports System.Data
Partial Class M_Source_04_MOA04107
    Inherits System.Web.UI.Page
    Dim user_id, org_uid, sP_Num As String
    Dim CCFun As New C_CheckFun
    Dim conn As New C_SQLFUN
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            '判斷導址值是否正確
            sP_Num = Request.QueryString("P_Num")
            If (sP_Num.Trim().Length < 1 Or (Not CCFun.isNumeric(sP_Num))) Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('派工單表單序號錯誤，請重新操作！');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            End If
        End If
        '查詢截取單一派工單詳細資料
        Dim sEFORMSN As String = DataRecords_DataRebinding(sP_Num)
        If sEFORMSN = "Error" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('派工單表單資料查詢失敗，請重新操作！');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        End If

        If sEFORMSN <> String.Empty Then
            '查詢此派工單之領料詳細資料
            ItCodeDataRecords_DataRebinding(sEFORMSN)
        End If
    End Sub
    Private Function DataRecords_DataRebinding(ByVal Job_num As String) As String
        Dim sEFORMSN As String = String.Empty
        Try
            Dim db As New SqlConnection(conn.G_conn_string)
            Dim dt As New DataTable("detail")
            Dim sQueryString As String = "SELECT a.*,b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as nPLACEName ,e.House_Name"
            sQueryString += " FROM P_0415 a left join P_0404 b on a.nbd_code = b.bd_code"
            sQueryString += " left join P_0406 c on a.nfl_code = c.fl_code"
            sQueryString += " left join P_0411 d on a.nrnum_code = d.rnum_code"
            sQueryString += " left join P_0416 e on a.nViewPer = e.House_Num"
            sQueryString += " WHERE a.P_NUM =@P_NUM"
            db.Open()
            Dim detailCom As New SqlCommand(sQueryString, db)
            Dim ds As New DataSet()
            Dim detailAdapter As New SqlDataAdapter(sQueryString, db)
            detailAdapter.SelectCommand.Parameters.Add("@P_NUM", SqlDbType.VarChar, 20).Value = Job_num
            detailAdapter.Fill(ds, "detail")
            If Not ds Is Nothing And ds.Tables.Count > 0 Then
                DetailsView1.DataSource = ds
                DetailsView1.DataBind()
                dt = ds.Tables(0)
                If (dt.Rows.Count > 0) Then
                    sEFORMSN = dt.Rows(0)("EFORMSN").ToString()
                End If
            End If
            db.Close()
        Catch ex As Exception
            sEFORMSN = "Error"
        End Try
        Return sEFORMSN
    End Function

    Private Sub ItCodeDataRecords_DataRebinding(ByVal sEFORMSN As String)
        Dim sErrorMsg As String = String.Empty
        Try
            Dim db As New SqlConnection(conn.G_conn_string)
            Dim dt As New DataTable("detail")
        
            Dim sQueryString As String = "select distinct P_0407.it_code,P_0407.it_name,P_0407.it_unit,b.cnt"
            sQueryString += " from P_0414 join P_0407 on P_0407.it_code = substring(P_0414.shcode,0,7) Left Join "
            sQueryString += " (select count(substring(P_0414.shcode,0,7)) as cnt, Job_num from P_0414 "
            sQueryString += " where Job_num = @Job_Num and shtype='0' and UseCheck ='2' group by substring(P_0414.shcode,0,7),Job_num) b"
            sQueryString += " on P_0414.Job_num = b.Job_num where P_0414.Job_Num =@Job_Num and shtype='0'"

            db.Open()
            Dim detailCom As New SqlCommand(sQueryString, db)
            Dim ds As New DataSet()
            Dim detailAdapter As New SqlDataAdapter(sQueryString, db)
            detailAdapter.SelectCommand.Parameters.Add("@Job_Num", SqlDbType.VarChar, 20).Value = sEFORMSN
            detailAdapter.Fill(ds, "detail")
            If Not ds Is Nothing And ds.Tables.Count > 0 Then
                If ds.Tables(0).Rows.Count > 0 Then
                    gv_itData.DataSource = ds
                    gv_itData.DataBind()
                    Pnl_ItCodeDetail.Visible = True
                End If
            End If

            db.Close()
        Catch ex As Exception
            sErrorMsg = ex.Message
        End Try
    End Sub


    Public Function StatusName(ByVal FlowStatus As String) As String
        Select Case FlowStatus
            Case "1"
                Return "新送件"
            Case "2"
                Return "處理中"
            Case "3"
                Return "待料中"
            Case "4"
                Return "完工"
            Case Else
                Return String.Empty
        End Select
    End Function

    Public Function ErrorWord(ByVal nErrCause As Object) As String
        Dim sErrorWord = String.Empty
        If Not nErrCause Is Nothing Then
            Dim sErrorCase As String = nErrCause.ToString()
            Select Case sErrorCase
                Case "1"
                    sErrorWord = "人為因素"
                Case "2"
                    sErrorWord = "自然因素"
                Case "3"
                    sErrorWord = "維護查報"
            End Select
        End If
        Return sErrorWord
    End Function
    Public Function CheckDate(ByVal nCheckDate As Object) As String
        Dim sResult = String.Empty
        If Not nCheckDate Is Nothing Then
            If nCheckDate.ToString.Trim().Length <> 0 Then
                Try
                    Dim dt As New DateTime
                    dt = DateTime.Parse(nCheckDate)
                    Dim sDataDT As String = dt.ToString("yyyy/MM/dd")
                    If DateTime.Parse(sDataDT) >= DateTime.Parse("2000/01/01") Then
                        sResult = DateTime.Parse(sDataDT).ToString("yyyy/MM/dd")
                    End If
                Catch ex As Exception
                    sResult = "error"
                End Try
            End If
        End If
        Return sResult
    End Function
End Class
