Imports System.Data.SqlClient
Imports System.Data

Partial Class M_Source_03_MOA03011
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim dt As DataTable = New DataTable()
        Dim eFormID As String = GetEFormId("派車申請單") '取得派車申請單eForm ID
        Dim strEFORMSN As String = Request("EFORMSN").ToString()
        'Dim strEFORMSN As String = "SYZO5C5D95AY9M22"
        Try
            Using conn As New SqlConnection(connstr)
                Dim da As New SqlDataAdapter("select * from P_03 where EFORMSN = '" + strEFORMSN + "'", conn)
                da.Fill(dt)

                If dt.Rows.Count > 0 Then
                    'nAPPLYTIME 申請時間
                    Dim applyTime As DateTime = Convert.ToDateTime(dt.Rows(0).Item("nAPPLYTIME"))
                    lbApplyYear.Text = (applyTime.Year - 1911).ToString() '接受申請時間(年)
                    lbApplyMonth.Text = applyTime.Month.ToString()        '接受申請時間(月)
                    lbApplyDay.Text = applyTime.Day.ToString()            '接受申請時間(日)
                    lbApplyHour.Text = applyTime.Hour.ToString()          '接受申請時間(時)
                    lbApplyMinute.Text = applyTime.Minute.ToString()      '接受申請時間(分)
                    'nReason 任務理由
                    lbReason.Text = dt.Rows(0).Item("nReason").ToString() '任務理由
                    '任務使用時間
                    lbUseStartMonth.Text = Convert.ToDateTime(dt.Rows(0).Item("nUSEDATE")).Month.ToString() '任務使用起始時間(月)
                    lbUseStartDay.Text = Convert.ToDateTime(dt.Rows(0).Item("nUSEDATE")).Day.ToString() '任務使用起始時間(日)
                    lbUseStartHour.Text = dt.Rows(0).Item("nSTUSEHOUR").ToString()  '任務使用起始時間(時)
                    lbUseStartMinute.Text = dt.Rows(0).Item("nSTUSEMIN").ToString() '任務使用起始時間(分)
                    lbUseEndMonth.Text = Convert.ToDateTime(dt.Rows(0).Item("nEDUSEDATE")).Month.ToString() '任務使用終止時間(月)
                    lbUseEndDay.Text = Convert.ToDateTime(dt.Rows(0).Item("nEDUSEDATE")).Day.ToString() '任務使用終止時間(日)
                    lbUseEndHour.Text = dt.Rows(0).Item("nEDUSEHOUR").ToString()    '任務使用終止時間(時)
                    lbUseEndMinute.Text = dt.Rows(0).Item("nEDUSEMIN").ToString()   '任務使用終止時間(分)
                    'nITEM 人員項目
                    lbItem.Text = dt.Rows(0).Item("nITEM").ToString()     '人員項目
                    'nARRIVEPLACE 車輛報到地點
                    lbArrivePlace.Text = dt.Rows(0).Item("nARRIVEPLACE").ToString() '車輛報到地點
                    'nARRIVETO 向何人報到
                    lbnArriveTo.Text = dt.Rows(0).Item("nARRIVETO").ToString()   '向何人報到
                    '車輛報到日期
                    lbArriveMonth.Text = Convert.ToDateTime(dt.Rows(0).Item("nARRDATE")).Month.ToString() '車輛報到日期(月)
                    lbArriveDay.Text = Convert.ToDateTime(dt.Rows(0).Item("nARRDATE")).Day.ToString() '車輛報到日期(日)
                    lbArriveHour.Text = dt.Rows(0).Item("nSTHOUR").ToString()   '車輛報到日期(時)
                    lbArriveMinute.Text = dt.Rows(0).Item("nEDHOUR").ToString() '車輛報到日期(分)
                    'nSTYLE 車輛品名型式
                    lbStyle.Text = dt.Rows(0).Item("nSTYLE").ToString() '使用車輛品名及型式
                    'nSTARTPOINT 起點
                    lbStartPoint.Text = dt.Rows(0).Item("nSTARTPOINT").ToString() '地點
                    'nENDPOINT 目的地
                    lbEndPoint.Text = dt.Rows(0).Item("nENDPOINT").ToString()     '目的地
                    '預定行程
                    Dim dtRoute As DataTable = GetP_0305(strEFORMSN) '取得預定行程資料
                    If dtRoute.Rows.Count > 0 Then
                        lbGoLocal1.Text = dtRoute.Rows(0).Item("GoLocal").ToString()
                        lbEndLocal1.Text = dtRoute.Rows(0).Item("EndLocal").ToString()
                    End If
                    If dtRoute.Rows.Count > 1 Then
                        lbGoLocal2.Text = dtRoute.Rows(1).Item("GoLocal").ToString()
                        lbEndLocal2.Text = dtRoute.Rows(1).Item("EndLocal").ToString()
                    End If
                    If dtRoute.Rows.Count > 2 Then
                        lbGoLocal3.Text = dtRoute.Rows(2).Item("GoLocal").ToString()
                        lbEndLocal3.Text = dtRoute.Rows(2).Item("EndLocal").ToString()
                    End If
                    If dtRoute.Rows.Count > 3 Then
                        lbGoLocal4.Text = dtRoute.Rows(3).Item("GoLocal").ToString()
                        lbEndLocal4.Text = dtRoute.Rows(3).Item("EndLocal").ToString()
                    End If
                    If dtRoute.Rows.Count > 4 Then
                        lbGoLocal5.Text = dtRoute.Rows(4).Item("GoLocal").ToString()
                        lbEndLocal5.Text = dtRoute.Rows(4).Item("EndLocal").ToString()
                    End If
                    '申請單位主管級職姓名
                    Dim dtFirstSuperior As DataTable = GetFirstSuperiorData(strEFORMSN, eFormID) '取得申請人上一級主管資料
                    lbSuperiorDept1.Text = dtFirstSuperior.Rows(0).Item("ORG_NAME").ToString()         '主管單位
                    lbSuperiorName1.Text = dtFirstSuperior.Rows(0).Item("emp_chinese_name").ToString() '主管姓名
                    lbNow1.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtFirstSuperior.Rows(0).Item("hddate"))) '批核時間,顯示格式: yyyMMddhhmm
                    '申請人單位級職姓名
                    Dim dtApplicant As DataTable = GetApplicantData(strEFORMSN, eFormID) '取得申請人資料
                    lbPADept.Text = dtApplicant.Rows(0).Item("ORG_NAME").ToString()         '申請人單位
                    lbPAName.Text = dtApplicant.Rows(0).Item("emp_chinese_name").ToString() '申請人姓名
                    lbNow2.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtApplicant.Rows(0).Item("hddate"))) '批核時間,顯示格式: yyyMMddhhmm
                    '部隊長批示
                    Dim dtSecendSuperior As DataTable = GetSecondSuperiorData(strEFORMSN, eFormID) '取得派車管制單位最後一位審核之上一級主管資料
                    lbSuperiorDept2.Text = dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 1).Item("ORG_NAME").ToString()         '部隊長單位
                    lbSuperiorName2.Text = dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 1).Item("emp_chinese_name").ToString() '部隊長姓名
                    lbNow3.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 1).Item("hddate"))) '批核時間,顯示格式: yyyMMddhhmm
                    '調派軍管審核意見
                    If dtSecendSuperior.Rows.Count = 2 Then '部隊長若有呈轉至上一級一次
                        lbTransfereeDept1.Text = dtSecendSuperior.Rows(0).Item("ORG_NAME").ToString()         '部隊長單位
                        lbTransfereeName1.Text = dtSecendSuperior.Rows(0).Item("emp_chinese_name").ToString() '部隊長姓名
                        lbTransfereeOpnion1.Text = "擬派" '此欄位顯示擬派字樣
                        lbTransfereeTime1.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtSecendSuperior.Rows(0).Item("hddate"))) '批核時間,顯示格式: yyyMMddhhmm
                    End If
                    If dtSecendSuperior.Rows.Count > 2 Then '部隊長上一級若呈轉兩次以上
                        '顯示倒數第二位審核之部隊長單位
                        lbTransfereeDept1.Text = dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 3).Item("ORG_NAME").ToString()
                        '顯示倒數第二位審核之部隊長姓名
                        lbTransfereeName1.Text = dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 3).Item("emp_chinese_name").ToString()
                        '顯示倒數第二位審核之部隊長批核時間,顯示格式: yyyMMddhhmm
                        lbTransfereeTime1.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 3).Item("hddate")))
                        '顯示倒數第三位審核之部隊長單位
                        lbTransfereeDept2.Text = dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 2).Item("ORG_NAME").ToString()
                        '顯示倒數第三位審核之部隊長姓名
                        lbTransfereeName2.Text = dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 2).Item("emp_chinese_name").ToString()
                        lbTransfereeOpnion2.Text = "擬派" '此欄位顯示擬派字樣
                        '顯示倒數第三位審核之部隊長批核時間,顯示格式: yyyMMddhhmm
                        lbTransfereeTime2.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtSecendSuperior.Rows(dtSecendSuperior.Rows.Count - 2).Item("hddate")))
                    End If
                    Dim dtControl As DataTable = GetControlAndDispatch(strEFORMSN, "派車管制單位") '取得派車管制單位批核者資料
                    If dtSecendSuperior.Rows.Count = 1 Then '若部隊長審核關卡無呈轉
                        lbControlOpnion.Text = "擬派" '則此欄位顯示擬派字樣
                    End If
                    lbControlDept.Text = dtControl.Rows(0).Item("ORG_NAME").ToString()           '批核者單位
                    lbControlName.Text = dtControl.Rows(0).Item("emp_chinese_name").ToString()   '批核者姓名
                    lbControlTime.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtControl.Rows(0).Item("hddate"))) '批核時間,顯示格式: yyyMMddhhmm
                    Dim dtDispatch As DataTable = GetControlAndDispatch(strEFORMSN, "派車調度單位") '取得派車管制單位批核者資料
                    lbDispatchDept.Text = dtDispatch.Rows(0).Item("ORG_NAME").ToString()          '批核者單位
                    lbDispatchName.Text = dtDispatch.Rows(0).Item("emp_chinese_name").ToString()  '批核者姓名
                    lbDispatchTime.Text = ConvertDateTimeFormat(Convert.ToDateTime(dtDispatch.Rows(0).Item("hddate"))) '批核時間,顯示格式: yyyMMddhhmm
                End If
            End Using
        Catch ex As Exception
            dt = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 轉換批核時間為民國年顯示格式字串
    ''' </summary>
    ''' <param name="varDateTime">批核時間</param>
    ''' <returns>格式yyyMMddhhmm之民國年時間顯示字串</returns>
    ''' <remarks></remarks>
    Private Function ConvertDateTimeFormat(ByVal varDateTime As DateTime) As String
        Return (varDateTime.Year - 1911).ToString() + varDateTime.Month.ToString("00") _
            + varDateTime.Day.ToString("00") + varDateTime.Hour.ToString("00") + varDateTime.Minute.ToString("00")
    End Function

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()
        Dim sqlcomm As New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db)
        GetEFormId = sqlcomm.ExecuteScalar().ToString()
        db.Close()
    End Function

    ''' <summary>
    ''' 取得派車申請單Flow中,第一個(即申請人之)上一級主管資料
    ''' </summary>
    ''' <param name="eformSN">表單ID</param>
    ''' <param name="eformID">派車申請單eForm ID</param>
    ''' <returns>上一級主管資料DataTable</returns>
    ''' <remarks></remarks>
    Private Function GetFirstSuperiorData(ByVal eformSN As String, ByVal eformID As String) As DataTable
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim dt As DataTable = New DataTable()
        Using conn As New SqlConnection(connstr)
            'Dim strSql As String = "SELECT VE.ORG_NAME, VE.emp_chinese_name, f.hddate FROM flowctl AS f "
            'strSql = strSql + "JOIN V_EmpInfo AS VE ON f.empuid = VE.employee_id "
            ''strSql = strSql + "WHERE VE.leave='Y' AND f.eformsn = '" + eformSN + "' AND f.gonogo = '1' "
            'strSql = strSql + "WHERE f.eformsn = '" + eformSN + "' AND f.gonogo = '1' "
            'strSql = strSql + "AND f.stepsid = ( SELECT nextstep FROM flow WHERE eformid = '" + eformSN + "' AND stepsid = 1 )"
            ''人員單位需顯示申請時的單位 2013/6/25 paul
            Dim strSql As String = "SELECT GP.ORG_NAME, f.emp_chinese_name,f.eformid, f.hddate FROM flowctl AS f"
            strSql = strSql & " JOIN ADMINGROUP AS GP ON f.DEPTCODE = GP.ORG_UID"
            strSql = strSql & " WHERE f.eformsn = '" + eformSN + "' AND f.gonogo = '1'"
            strSql = strSql & " AND f.stepsid = ( SELECT nextstep FROM flow WHERE eformid = '" + eformID + "' AND stepsid = 1 )"
            Dim da As New SqlDataAdapter(strSql, conn)
            da.Fill(dt)
        End Using
        Return dt
    End Function

    ''' <summary>
    ''' 取得派車申請單Flow中,第二個(即派車管制單位之)上一級主管資料
    ''' </summary>
    ''' <param name="eformSN">表單ID</param>
    ''' <param name="eformID">派車申請單eForm ID</param>
    ''' <returns>上一級主管資料DataTable</returns>
    ''' <remarks></remarks>
    Private Function GetSecondSuperiorData(ByVal eformSN As String, ByVal eformID As String) As DataTable
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim dt As DataTable = New DataTable()
        Using conn As New SqlConnection(connstr)
            'Dim strSql As String = "SELECT E.ORG_NAME,E.AD_TITLE,E.emp_chinese_name,f.hddate,f.gonogo FROM flowctl AS f "
            'strSql = strSql + "JOIN V_EmpInfo AS E ON f.empuid = E.employee_id WHERE eformsn = '" + eformSN + "' "
            'strSql = strSql + "AND nextstep = ( SELECT stepsid FROM flow WHERE eformid = '" + eformID + "' "
            'strSql = strSql + "AND nextstep = -1 AND group_id <> '-1' ) ORDER BY f.flowsn"
            ''人員單位需顯示申請時的單位 2013/6/25 paul
            Dim strSql As String = "SELECT E.AD_TITLE,GP.ORG_NAME,F.EMP_CHINESE_NAME,F.HDDATE,F.GONOGO FROM FLOWCTL AS F"
            strSql = strSql & " LEFT JOIN ADMINGROUP GP ON F.DEPTCODE=GP.ORG_UID"
            strSql = strSql & " LEFT JOIN V_EMPINFO E ON F.EMPUID=E.EMPLOYEE_ID"
            strSql = strSql & " WHERE F.EFORMSN='" & eformSN & "' AND F.NEXTSTEP=(SELECT STEPSID FROM FLOW WHERE EFORMID='" & eformID & "' AND NEXTSTEP= -1 AND GROUP_ID<> '-1')"
            strSql = strSql & " ORDER BY F.FLOWSN"
            Dim da As New SqlDataAdapter(strSql, conn)
            da.Fill(dt)
        End Using
        Return dt
    End Function

    ''' <summary>
    ''' 取得派車申請單Flow中,申請人資料
    ''' </summary>
    ''' <param name="eformSN">表單ID</param>
    ''' <param name="eformID">派車申請單eForm ID</param>
    ''' <returns>上一級主管資料DataTable</returns>
    ''' <remarks></remarks>
    Private Function GetApplicantData(ByVal eformSN As String, ByVal eformID As String) As DataTable
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim dt As DataTable = New DataTable()
        Using conn As New SqlConnection(connstr)
            'Dim strSql As String = "SELECT VE.ORG_NAME, VE.emp_chinese_name, f.hddate FROM flowctl AS f "
            'strSql = strSql + "JOIN V_EmpInfo AS VE ON f.empuid = VE.employee_id "
            'strSql = strSql + "WHERE f.eformsn = '" + eformSN + "' AND f.gonogo = '-'"
            '' strSql = strSql + "WHERE VE.leave='Y' AND f.eformsn = '" + eformSN + "' AND f.gonogo = '-'"
            ''人員單位需顯示申請時的單位 2013/6/25 paul
            Dim strSql As String = "SELECT GP.ORG_NAME,GP.ORG_UID, f.emp_chinese_name, f.hddate FROM flowctl AS f "
            strSql = strSql & "JOIN ADMINGROUP AS GP ON f.DEPTCODE = GP.ORG_UID"
            strSql = strSql + " WHERE f.eformsn = '" & eformSN & "' AND f.gonogo = '-'"
            Dim da As New SqlDataAdapter(strSql, conn)
            da.Fill(dt)
        End Using
        Return dt
    End Function

    ''' <summary>
    ''' 取得派車調度或管制單位資料
    ''' </summary>
    ''' <param name="eformSN">表單ID</param>
    ''' <param name="strType">調度或管制單位名稱</param>
    ''' <returns>批核者資料DataTable</returns>
    ''' <remarks></remarks>
    Private Function GetControlAndDispatch(ByVal eformSN As String, ByVal strType As String) As DataTable
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim dt As DataTable = New DataTable()
        Using conn As New SqlConnection(connstr)
            'Dim strSql As String = "SELECT E.ORG_NAME,E.AD_TITLE,E.emp_chinese_name,fl.hddate,fl.gonogo FROM flowctl AS fl "
            'strSql = strSql + "JOIN flow AS f ON fl.stepsid = f.stepsid "
            'strSql = strSql + "JOIN SYSTEMOBJ AS S ON f.group_id = S.object_uid "
            'strSql = strSql + "JOIN V_EmpInfo AS E ON fl.empuid = E.employee_id "
            'strSql = strSql + "WHERE eformsn = '" + eformSN + "' AND S.object_name='" + strType + "' "
            'strSql = strSql + "AND ( gonogo='1' OR gonogo='T' ) ORDER BY hddate DESC"
            ''人員單位需顯示申請時的單位 2013/6/25 paul
            Dim strSql As String = "SELECT GP.ORG_NAME,E.AD_TITLE,E.emp_chinese_name,fl.hddate,fl.gonogo FROM flowctl AS fl "
            strSql = strSql & " LEFT JOIN ADMINGROUP GP ON FL.DEPTCODE=GP.ORG_UID"
            strSql = strSql & " JOIN flow AS f ON fl.stepsid = f.stepsid "
            strSql = strSql & " JOIN SYSTEMOBJ AS S ON f.group_id = S.object_uid "
            strSql = strSql & " JOIN V_EmpInfo AS E ON fl.empuid = E.employee_id "
            strSql = strSql & " WHERE eformsn = '" & eformSN & "' AND S.object_name='" & strType & "' "
            strSql = strSql & " AND ( gonogo='1' OR gonogo='T' ) ORDER BY hddate DESC"
            Dim da As New SqlDataAdapter(strSql, conn)
            da.Fill(dt)
        End Using
        Return dt
    End Function

    ''' <summary>
    ''' 取得派車申請單預定行程資料
    ''' </summary>
    ''' <param name="eformSN">表單ID</param>
    ''' <returns>申請單預定行程資料DataTable</returns>
    ''' <remarks></remarks>
    Private Function GetP_0305(ByVal eformSN As String) As DataTable
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim dt As DataTable = New DataTable()
        Using conn As New SqlConnection(connstr)
            Dim strSql As String = "SELECT * FROM P_0305 WHERE EFORMSN = '" + eformSN + "' ORDER BY Local_Num"
            Dim da As New SqlDataAdapter(strSql, conn)
            da.Fill(dt)
        End Using
        Return dt
    End Function

End Class
