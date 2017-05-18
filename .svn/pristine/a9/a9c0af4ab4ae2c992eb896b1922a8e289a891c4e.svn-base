Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_09_MOA09005
    Inherits System.Web.UI.Page
    Dim org_uid As String
    Dim sql_function As New C_SQLFUN
    Dim connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Session("files") = Request.Files                 '附件檔案
            hidEFormID.Value = Request.QueryString("eFormID")     '表單種類ID
            hidEFormSN.Value = Request.QueryString("eFormSN")     '表單ID
            hidFillOutBy.Value = Request.QueryString("FillOutBy") '填表人員工ID
            hidCreateBy.Value = Request.Form("ddlCreatorName")    '申請人員工ID
            hidSubject.Value = Request.Form("txtSubject")         '開會事由
            hidLocation.Value = Request.Form("txtLocation")       '開會地點
            hidSponsor.Value = Request.QueryString("Sponsor")     '承辦單位
            hidModerator.Value = Request.Form("txtModerator")     '主持人
            hidPhoneNumber.Value = Request.Form("txtPhoneNumber") '聯絡人電話
            hidMeetingDate.Value = Request.Form("txtMeetingDate") '開會日期
            hidHour.Value = Request.Form("ddlHour")               '開會時間(時)
            hidMinute.Value = Request.Form("ddlMinute")           '開會時間(分)
            hidDocumentNo.Value = Request.Form("txtDocumentWord") + "字第" & Request.Form("txtDocumentNo") + "號" '發文字號
            hidEnteringPeopleNumber.Value = Request.Form("txtEnteringPeopleNumber") '進出人員(部外)
            '進出營門 0:博愛營區一號門 1:博愛營區二號門 2:博愛營區三號門 3:博愛營區四號門 4:採購中心 5:博一大樓 6:博二大樓
            hidEnteringGate.Value = Request.Form("hidEnteringGate") '資料字串紀錄格式: 營門編號1,營門編號2,...
        End If
    End Sub

    Protected Sub btnSend_Click(sender As Object, e As System.EventArgs) Handles btnSend.Click
        Dim files As HttpFileCollection = Session("files") '附件檔案
        Session.Remove("files")

        Dim redirectFlag As Boolean = True                 '執行是否成功,若成功則轉址
        Dim Val_P As String = ""                           '表單送件回傳值用變數
        '取得欲呈送之門禁會議管制群組成員資料(審核者)
        Dim drEmployeeApprove As DataRow = GetEmployee(DDLUser.SelectedValue).Rows(0)
        Try
            db.Open()
            '新增UPLOAD附件檔案上傳資料,並儲存檔案
            Dim saveFileResult As String = SavingUploadFiles(db, hidEFormSN.Value, files)
            If Not saveFileResult.Equals("OK") Then
                Throw New Exception(saveFileResult)
            End If

            '表單審核,FCC.SendVal用參數,格式: 表單種類ID,申請者ID,eFormID,1,審核者ID
            Dim SendVal = hidEFormID.Value & "," & hidCreateBy.Value & "," & hidEFormSN.Value & "," & "1" & "," & drEmployeeApprove("employee_id").ToString()

            '表單送件,寄信及增加flowctl資料表內申請人及其下一關審核者資料
            Dim FCC As New CFlowSend
            Val_P = FCC.F_Send(SendVal, connstr)

            '新增P_09門禁會議管制表單資料,insertP_09函數若執行成功則回傳該筆Insert資料之PK,失敗則回傳錯誤訊息
            Dim insertP_09Result As String = InserP_09(db, hidEFormSN.Value, hidCreateBy.Value, hidSubject.Value, hidLocation.Value _
                                                       , hidSponsor.Value, hidModerator.Value, hidPhoneNumber.Value, hidMeetingDate.Value _
                                                       , hidHour.Value, hidMinute.Value, hidDocumentNo.Value, hidEnteringPeopleNumber.Value _
                                                       , hidEnteringGate.Value)
            If Not insertP_09Result.Equals("OK") Then
                Throw New Exception(insertP_09Result)
            End If
        Catch ex As Exception
            redirectFlag = False
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('申請單送出失敗:" + ex.Message + "');")
            Response.Write(" </script>")
            Response.End()
        Finally
            db.Close()
        End Try
        If redirectFlag Then '若程式執行成功,則將頁面轉址
            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
        End If
    End Sub

    ''' <summary>
    ''' 取得員工資料
    ''' </summary>
    ''' <param name="employeeId">員工ID</param>
    ''' <returns>員工資料</returns>
    ''' <remarks></remarks>
    Private Function GetEmployee(ByVal employeeId As String) As DataTable
        db.Open()
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM EMPLOYEE WHERE employee_id = '" + employeeId + "'", db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        GetEmployee = dt
    End Function

    ''' <summary>
    ''' 儲存附件檔案並新增資料庫資料
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eformsn">表單審核流程表單資料ID</param>
    ''' <param name="files">HttpFileCollection</param>
    ''' <returns>執行成功訊息OK或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function SavingUploadFiles(ByRef db As SqlConnection, ByVal eformsn As String, ByVal files As HttpFileCollection) As String
        Dim result As String = "OK"
        Try
            For i As Integer = 0 To files.Count - 1
                If Not String.IsNullOrEmpty(files(i).FileName) Then
                    '儲存檔案名稱加上eformsn產生唯一性
                    Dim fileNameArray As String() = files(i).FileName.Split("\")
                    Dim fileName As String = fileNameArray(fileNameArray.Length - 1)
                    Dim saveFileName = eformsn & "-" & fileName
                    Dim uploadPath = Server.MapPath("/MOA/M_Source/99/" & saveFileName)
                    files(i).SaveAs(uploadPath)

                    Dim insertComm = New SqlCommand("INSERT INTO UPLOAD(eformsn,FileName,FilePath,Upload_Time) VALUES(@eformsn,@FileName,@FilePath,GETDATE())", db)
                    insertComm.Parameters.Add("@eformsn", SqlDbType.VarChar, 16).Value = eformsn.Trim()
                    insertComm.Parameters.Add("@FileName", SqlDbType.VarChar, 100).Value = saveFileName.Trim()
                    insertComm.Parameters.Add("@FilePath", SqlDbType.VarChar, 255).Value = "/MOA/M_Source/99/"
                    insertComm.ExecuteNonQuery()
                    insertComm.Dispose()
                End If
            Next
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 新增P_09門禁管制會議申請表資料
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eformsn">表單審核流程表單資料ID</param>
    ''' <param name="strCreateBy">申請人員工ID</param>
    ''' <param name="strSubject">開會事由</param>
    ''' <param name="strLocation">開會地點</param>
    ''' <param name="strSponsor">承辦單位</param>
    ''' <param name="strModerator">主持人</param>
    ''' <param name="strPhoneNumber">聯絡人電話</param>
    ''' <param name="strMeetingDate">開會日期</param>
    ''' <param name="strHour">開會時間(時)</param>
    ''' <param name="strMinute">開會時間(分)</param>
    ''' <param name="strDocumentNo">發文字號</param>
    ''' <param name="strEnteringPeopleNumber">進出人員數量</param>
    ''' <param name="strEnteringGate">進出營門編號</param>
    ''' <returns>執行成功訊息OK或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function InserP_09(ByRef db As SqlConnection, ByVal eformsn As String, ByVal strCreateBy As String _
                               , ByVal strSubject As String, ByVal strLocation As String, ByVal strSponsor As String _
                               , ByVal strModerator As String, ByVal strPhoneNumber As String, ByVal strMeetingDate As String _
                               , ByVal strHour As String, ByVal strMinute As String, ByVal strDocumentNo As String _
                               , ByVal strEnteringPeopleNumber As String, ByVal strEnteringGate As String) As String
        Dim result As String = "OK"
        Dim strSQL As String = "INSERT INTO P_09(EFORMSN,FillOutBy,CreateBy,Subject,MeetingDate,MeetingHour,MeetingMinute,Location"
        strSQL = strSQL + ",Sponsor,Moderator,PhoneNumber,DocumentNo,EnteringPeopleNumber,Status,CreateDate,PENDFLAG) "
        strSQL = strSQL + "VALUES(@EFORMSN,@FillOutBy,@CreateBy,@Subject,@MeetingDate,@MeetingHour,@MeetingMinute,@Location"
        'Status:表單狀態 [ 0：未登管 1：已登管 2：退件 3：刪除 ]
        strSQL = strSQL + ",@Sponsor,@Moderator,@PhoneNumber,@DocumentNo,@EnteringPeopleNumber,0,GETDATE(),'-')"
        Try
            Dim sqlComm = New SqlCommand(strSQL, db)
            sqlComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eformsn                  '表單審核流程表單資料ID
            sqlComm.Parameters.Add("@FillOutBy", SqlDbType.VarChar, 10).Value = hidFillOutBy.Value     '填表人員工編號
            sqlComm.Parameters.Add("@CreateBy", SqlDbType.VarChar, 10).Value = strCreateBy             '申請人員工編號
            sqlComm.Parameters.Add("@Subject", SqlDbType.NVarChar, 100).Value = strSubject             '開會事由
            sqlComm.Parameters.Add("@MeetingDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strMeetingDate) '開會日期
            sqlComm.Parameters.Add("@MeetingHour", SqlDbType.Int).Value = Convert.ToInt32(strHour)     '開會時間(時)
            sqlComm.Parameters.Add("@MeetingMinute", SqlDbType.Int).Value = Convert.ToInt32(strMinute) '開會時間(時)
            sqlComm.Parameters.Add("@Location", SqlDbType.NVarChar, 100).Value = strLocation           '開會地點
            sqlComm.Parameters.Add("@Sponsor", SqlDbType.NVarChar, 50).Value = strSponsor              '承辦單位
            sqlComm.Parameters.Add("@Moderator", SqlDbType.NVarChar, 50).Value = strModerator          '主持人
            sqlComm.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar, 50).Value = strPhoneNumber      '聯絡人電話
            sqlComm.Parameters.Add("@DocumentNo", SqlDbType.NVarChar, 50).Value = strDocumentNo        '發文字號
            sqlComm.Parameters.Add("@EnteringPeopleNumber", SqlDbType.Int).Value = Convert.ToInt32(strEnteringPeopleNumber) '進出人員數量
            sqlComm.ExecuteNonQuery()
            Dim gateNoArray As String() = strEnteringGate.Split(",")
            For i As Integer = 0 To gateNoArray.Length - 1
                '進出營門 0:博愛營區一號門 1:博愛營區二號門 2:博愛營區三號門 3:博愛營區四號門 4:採購中心 5:博一大樓 6:博二大樓
                Dim insertP_0901Result As String = InserP_0901(db, eformsn, gateNoArray(i))
                If Not insertP_0901Result.Equals("OK") Then
                    Throw New Exception(insertP_0901Result)
                End If
            Next
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 新增P_0901門禁管制會議申請表資料
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eFormSN">表單審核流程表單資料ID</param>
    ''' <param name="gateNo">出入營門編號</param>
    ''' <returns>執行成功訊息OK或錯誤訊息</returns>
    ''' <remarks></remarks>
    Private Function InserP_0901(ByRef db As SqlConnection, ByVal eFormSN As String, ByVal gateNo As String) As String
        Dim result As String = "OK"
        Dim strSQL As String = "INSERT INTO P_0901(EFORMSN,GateNumber) VALUES(@EFORMSN,@GateNumber)"
        Try
            Dim sqlComm = New SqlCommand(strSQL, db)
            sqlComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = eFormSN '表單審核流程表單資料ID
            '進出營門 0:博愛營區一號門 1:博愛營區二號門 2:博愛營區三號門 3:博愛營區四號門 4:採購中心 5:博一大樓 6:博二大樓
            sqlComm.Parameters.Add("@GateNumber", SqlDbType.Int).Value = Convert.ToInt32(gateNo)
            sqlComm.ExecuteNonQuery()
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

End Class
