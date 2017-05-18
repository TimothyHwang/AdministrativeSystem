Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports Microsoft.VisualBasic

Public Class C_Public
#Region "亂數(產生A~Z,0~9,的亂數)"
    ''' <summary>
    ''' 亂數(產生A~Z,0~9,的亂數)
    ''' </summary>
    ''' <param name="rand_num">位數</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function randstr(ByVal rand_num As Integer) As String
        Dim num As String = ""
        Randomize()                          '初始化亂數產生器        
        For i As Integer = 1 To rand_num
            Select Case CInt(Int((2 * Rnd()) + 1))
                Case 1
                    num = num & Chr(CInt(Int((26 * Rnd()) + 65)))
                Case 2
                    num = num & CInt(Int((9 * Rnd()) + 1))
            End Select
        Next
        randstr = num
    End Function
    ''' <summary>
    ''' 建立唯一的eformsn
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CreateNewEFormSN() As String
        Dim C_Public As New C_Public
        Dim eformsn As String = C_Public.randstr(16)
        Do While checkEFormSNExist(eformsn)
            eformsn = C_Public.randstr(16)
        Loop
        CreateNewEFormSN = eformsn
    End Function
    ''' <summary>
    ''' 判斷申請單號是否重複
    ''' </summary>
    ''' <param name="eformsn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function checkEFormSNExist(ByVal eformsn As String) As Boolean
        Try
            Dim db As New SqlConnection((New C_SQLFUN).G_conn_string)
            db.Open()
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter("select flowsn from flowctl where eformsn = '" & eformsn & "'", db)
            da.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                checkEFormSNExist = True
            Else
                checkEFormSNExist = False
            End If
            db.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            checkEFormSNExist = True
        End Try
    End Function
#End Region
#Region "找出子組織"
    ''' <summary>
    ''' 找出子組織代碼
    ''' </summary>
    ''' <param name="org_id">父組織代碼</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getchildorg(ByVal org_id As String) As String
        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
        Dim EndFlag As Boolean
        Dim sOrgID, sql1, OrgIDAll As String
        Dim Maxnum, Org_int As Integer
        EndFlag = False
        Maxnum = 0
        Org_int = 0
        '下一階層組織代號
        sOrgID = ""
        '累計下一階層組織代號
        OrgIDAll = ""

        '查最大階層
        Dim sql = "Select Max(ORG_TREE_LEVEL) AS MaxNum From ADMINGROUP"
        db.Open()
        Dim dt2 As DataTable = New DataTable("ADMINGROUP")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(sql, db)
        da2.Fill(dt2)
        db.Close()
        For x As Integer = 0 To dt2.Rows.Count - 1
            Maxnum = CType(dt2.Rows(x).Item(0), Integer)
        Next

        While EndFlag = False
            '找出子組織
            If Org_int = 0 Then
                sql1 = "SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID In ('" & org_id & "') ORDER BY ORG_TREE_LEVEL"
            Else
                sql1 = "SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID In (" & sOrgID & ") ORDER BY ORG_TREE_LEVEL"
            End If
            db.Open()
            Dim dt1 As DataTable = New DataTable("ADMINGROUP")
            Dim da1 As SqlDataAdapter = New SqlDataAdapter(sql1, db)
            da1.Fill(dt1)
            db.Close()
            sOrgID = ""

            '搜尋下一階層組織代號
            If dt1.Rows.Count <> 0 Then
                For j As Integer = 0 To dt1.Rows.Count - 1
                    If dt1.Rows(j).Item(0) <> "" Then
                        If sOrgID = "" Then
                            sOrgID = CType(("'" & dt1.Rows(j).Item(0) & "'"), String)
                        Else
                            sOrgID = CType((sOrgID & ", '" & dt1.Rows(j).Item(0) & "'"), String)
                        End If

                        '將下一階層累計
                        If OrgIDAll = "" Then
                            OrgIDAll = CType(("'" & dt1.Rows(j).Item(0) & "'"), String)
                        Else
                            OrgIDAll = CType((OrgIDAll & ", '" & dt1.Rows(j).Item(0) & "'"), String)
                        End If
                    Else
                        EndFlag = True
                    End If
                Next
            End If

            If Org_int = Maxnum Then
                EndFlag = True
            End If

            If dt1.Rows.Count = 0 Then
                EndFlag = True
            End If

            Org_int = Org_int + 1
        End While

        If OrgIDAll = "" Then
            getchildorg = "'" & org_id & "'"
        Else
            getchildorg = "'" & org_id & "'," & OrgIDAll
        End If
    End Function
#End Region
#Region "找出一級單位或二級單位"
    Public Function getUporg(ByVal org_id As String, ByVal OrgKind As Integer) As String
        Try

            Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            Dim EndFlag As Boolean = False
            Dim UpSql As String
            Dim strParentOrg As String = ""
            Dim strOrgUid As String = ""
            Dim strKind As Integer = 0

            getUporg = ""

            While EndFlag = False

                If strParentOrg = "" Then
                    UpSql = "SELECT ORG_UID,PARENT_ORG_UID,ORG_KIND FROM ADMINGROUP WHERE ORG_UID = '" & org_id & "'"
                Else
                    UpSql = "SELECT ORG_UID,PARENT_ORG_UID,ORG_KIND FROM ADMINGROUP WHERE ORG_UID = '" & strParentOrg & "'"
                End If

                '取得上一級單位
                db.Open()
                Dim strPer As New SqlCommand(UpSql, db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.Read() Then
                    strOrgUid = CType(RdPer("ORG_UID"), String)
                    strParentOrg = CType(RdPer("PARENT_ORG_UID"), String)

                    If (RdPer("ORG_KIND") Is DBNull.Value) = False Then
                        strKind = CType(RdPer("ORG_KIND"), Integer)
                    End If
                Else
                    getUporg = "無任何一級或二級單位"
                    EndFlag = True
                End If
                db.Close()

                If strKind = OrgKind Then
                    getUporg = strOrgUid
                    EndFlag = True
                End If

            End While

        Catch ex As Exception
            getUporg = ex.Message
        End Try

    End Function
#End Region
#Region "新增登入者資訊"
    ''' <summary>
    ''' 紀錄登入者log
    ''' </summary>
    ''' <param name="strIP">IP Address</param>
    ''' <param name="strUser">User ID</param>
    ''' <param name="strPage">使用頁面</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoginAction(ByVal strIP As String, ByVal strUser As String, ByVal strPage As String) As String
        Try

            Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

            '新增登入者資料
            db.Open()
            Dim insCom As New SqlCommand("INSERT INTO LoginLog(Login_IP,Login_ID,Login_Page) VALUES ('" & strIP & "','" & strUser & "','" & strPage & "')", db)
            insCom.ExecuteNonQuery()
            db.Close()

            LoginAction = ""

        Catch ex As Exception
            LoginAction = ex.Message
        End Try

    End Function
#End Region
#Region "判斷登入者權限"
    Public Function LoginCheck(ByVal strUser As String, ByVal strPage As String) As String
        Try

            Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

            Dim strsql As String = "Select Count(ROLEMAPS.ProgMap) as LoginNum FROM ROLEGROUPITEM,ROLES,ROLEMAPS WHERE ROLEGROUPITEM.Group_Uid = ROLES.Group_uid AND ROLES.LinkNum = ROLEMAPS.LinkNum AND (ROLEGROUPITEM.employee_id = '" & strUser & "')  AND (ROLEMAPS.ProgMap LIKE '%" & strPage & "%')"

            db.Open()
            Dim strPer As New SqlCommand(strsql, db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.Read() Then
                If RdPer("LoginNum") = 0 Then
                    LoginCheck = "Err"
                Else
                    LoginCheck = ""
                End If
            Else
                LoginCheck = ""
            End If
            db.Close()

        Catch ex As Exception
            LoginCheck = ex.Message
        End Try

    End Function
#End Region
#Region "判斷使用者代理人"
    ''' <summary>
    ''' 查詢被代理人(查詢上一層)
    ''' </summary>
    ''' <param name="strAgentID">代理者ID</param>
    ''' <param name="strDate">代理時間</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DoubalAgentAll(ByVal strAgentID As String, ByVal strDate As String) As String

        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        '找出代理人
        Dim strsql As String = "SELECT * FROM AGENT "
        strsql += " WHERE '" & strDate & "' BETWEEN Agent_SDate AND Agent_EDate "
        strsql += " AND Agent1 = '" & strAgentID & "'"

        '是否已代理
        db.Open()
        Dim PerCountCom As New SqlCommand(strsql, db)
        Dim PerRdv = PerCountCom.ExecuteReader()
        If PerRdv.Read() Then
            DoubalAgentAll = PerRdv("Agent2").ToString()
        Else
            DoubalAgentAll = ""
        End If
        db.Close()

    End Function

#End Region
#Region "判斷使用者代理人員"
    ''' <summary>
    ''' 查詢代理人
    ''' </summary>
    ''' <param name="sPricipalID">被代理人ID</param>
    ''' <param name="sDate">代理時間(SmallDateString)</param>
    ''' <returns>代理人 ID</returns>
    ''' <remarks></remarks>
    Public Function GetAgentID(ByVal sPricipalID As String, ByVal sDate As String) As String
        Dim sReturn As String = ""
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        strSql = "SELECT * FROM AGENT WHERE AGENT2='" + sPricipalID + "' AND '" + sDate + "' BETWEEN AGENT_SDATE AND AGENT_EDATE"
        DR = DC.CreateReader(strSql)
        If DR.Read() Then
            sReturn = DR("AGENT1").ToString()
        End If
        DC.Dispose()
        GetAgentID = sReturn
    End Function
    ''' <summary>
    ''' 查詢代理人
    ''' </summary>
    ''' <param name="sPricipalID">被代理人ID</param>
    ''' <returns>代理人ID</returns>
    ''' <remarks></remarks>
    Public Function GetAgentID(ByVal sPricipalID As String) As String
        Dim sReturn As String = ""
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        strSql = "SELECT * FROM AGENT WHERE AGENT2='" + sPricipalID + "'"
        DR = DC.CreateReader(strSql)
        If DR.Read() Then
            sReturn = DR("AGENT1").ToString()
        End If
        DC.Dispose()
        GetAgentID = sReturn
    End Function
    ''' <summary>
    ''' 查詢所有被代理人
    ''' </summary>
    ''' <param name="strAgentID">代理人ID</param>
    ''' <param name="strDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AgentAll(ByVal strAgentID As String, ByVal strDate As String) As String

        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        '找出代理人
        Dim strsql As String = "SELECT * FROM AGENT "
        strsql += " WHERE '" & strDate & "' BETWEEN Agent_SDate AND Agent_EDate "
        strsql += " AND Agent1 = '" & strAgentID & "'"

        db.Open()
        Dim dt1 As DataTable = New DataTable()
        Dim da1 As SqlDataAdapter = New SqlDataAdapter(strsql, db)
        da1.Fill(dt1)
        db.Close()

        Dim RelUserALL As String = ""

        '搜尋全部代理人
        If dt1.Rows.Count <> 0 Then
            For j As Integer = 0 To dt1.Rows.Count - 1

                Dim RelUser As String = ""

                If dt1.Rows(j).Item("Agent2") <> "" Then

                    Dim EndFlag As Boolean = False
                    Dim NextUser As String = dt1.Rows(j).Item("Agent2").ToString()

                    '遞回找尋被代理人
                    While EndFlag = False
                        If DoubalAgentAll(NextUser, strDate) = "" Then
                            RelUser = NextUser
                            EndFlag = True
                        Else
                            NextUser = DoubalAgentAll(NextUser, strDate).ToString()
                            EndFlag = False
                        End If
                    End While

                Else
                    AgentAll = ""
                End If

                If RelUserALL = "" Then
                    RelUserALL = "'" & RelUser & "'"
                Else
                    RelUserALL = RelUserALL & ", '" & RelUser & "'"
                End If

            Next
        End If

        AgentAll = RelUserALL

    End Function

#End Region
#Region "判斷使用者是否已被代理"
    ''' <summary>
    ''' 找出被代理人的代理人ID
    ''' </summary>
    ''' <param name="strUserID">被代理人ID</param>
    ''' <param name="strSDate">代理開始時間</param>
    ''' <param name="strEDate">代理結束時間</param>
    ''' <returns>代理人ID</returns>
    ''' <remarks></remarks>
    Public Function DouBleAgent(ByVal strUserID As String, ByVal strSDate As String, ByVal strEDate As String) As String

        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        Dim strsql As String = "select * from AGENT "
        strsql += " WHERE ((Agent_SDate BETWEEN '" & strSDate & "' AND '" & strEDate & "') OR (Agent_EDate BETWEEN '" & strSDate & "' AND '" & strEDate & "')) "
        strsql += " AND Agent2 = '" & strUserID & "'"

        '是否已被代理
        db.Open()
        Dim PerCountCom As New SqlCommand(strsql, db)
        Dim PerRdv = PerCountCom.ExecuteReader()
        If PerRdv.Read() Then
            DouBleAgent = PerRdv("Agent1").ToString()
        Else
            DouBleAgent = ""
        End If
        db.Close()

    End Function

    ''' <summary>
    ''' 找出被代理人的代理人ID
    ''' </summary>
    ''' <param name="strUserID">被代理人ID</param>
    ''' <returns>代理人ID</returns>
    ''' <remarks></remarks>
    Public Function DouBleAgent(ByVal strUserID As String) As String

        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        Dim strsql As String = "select * from AGENT "
        strsql += " WHERE (GetDate() BETWEEN Agent_SDate AND Agent_EDate )"
        strsql += " AND Agent2 = '" & strUserID & "'"

        '是否已被代理
        db.Open()
        Dim PerCountCom As New SqlCommand(strsql, db)
        Dim PerRdv = PerCountCom.ExecuteReader()
        If PerRdv.Read() Then
            DouBleAgent = PerRdv("Agent1").ToString()
        Else
            DouBleAgent = ""
        End If
        db.Close()

    End Function
#End Region
#Region "判斷使用者最終被代理人"
    ''' <summary>
    ''' 依據找出代理人
    ''' </summary>
    ''' <param name="strUserID"></param>
    ''' <param name="strSDate"></param>
    ''' <param name="strEDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AgentFinal(ByVal strUserID As String, ByVal strSDate As String, ByVal strEDate As String) As String

        Dim EndFlag As Boolean = False
        Dim NextUser As String
        Dim RelUser As String = ""

        NextUser = DouBleAgent(strUserID, strSDate, strEDate).ToString()

        If NextUser = "" Then
            RelUser = strUserID
        Else
            '遞回找尋代理人
            While EndFlag = False
                If DouBleAgent(NextUser, strSDate, strEDate) = "" Then
                    RelUser = NextUser
                    EndFlag = True
                Else
                    NextUser = DouBleAgent(NextUser, strSDate, strEDate).ToString()
                    EndFlag = False
                End If
            End While
        End If

        AgentFinal = RelUser

    End Function
    ''' <summary>
    ''' 依據被代理人ID找出代理人
    ''' </summary>
    ''' <param name="strUserID">被代理人ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AgentFinal(ByVal strUserID As String) As String

        Dim EndFlag As Boolean = False
        Dim NextUser As String
        Dim RelUser As String = ""

        NextUser = DouBleAgent(strUserID).ToString()

        If NextUser = "" Then
            RelUser = strUserID
        Else
            '遞回找尋代理人
            While EndFlag = False
                If DouBleAgent(NextUser) = "" Then
                    RelUser = NextUser
                    EndFlag = True
                Else
                    NextUser = DouBleAgent(NextUser).ToString()
                    EndFlag = False
                End If
            End While
        End If

        AgentFinal = RelUser

    End Function
#End Region
#Region "判斷使用者批核權限"
    ''' <summary>
    ''' 判斷批核權限
    ''' </summary>
    ''' <param name="strEformsn">表單ID</param>
    ''' <param name="strUserID">簽核者ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApproveAuth(ByVal strEformsn As String, ByVal strUserID As String) As String

        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        '找出批核表單
        Dim strsql As String = "SELECT * FROM flowctl "
        strsql += " WHERE eformsn = '" & strEformsn & "' AND hddate IS NULL "

        db.Open()
        Dim dt1 As DataTable = New DataTable()
        Dim da1 As SqlDataAdapter = New SqlDataAdapter(strsql, db)
        da1.Fill(dt1)
        db.Close()

        '權限判斷
        Dim AuthFlag As String = ""

        '搜尋全部批核人員
        If dt1.Rows.Count <> 0 Then
            For j As Integer = 0 To dt1.Rows.Count - 1

                '判斷表單批核人員是否是該使用者
                If UCase(dt1.Rows(j).Item("empuid")) = UCase(strUserID) Then
                    AuthFlag = "Y"
                Else
                    '判斷代理人條件有錯誤，另寫一function取代 20130708 paul
                    '判斷代理人是否是該使用者
                    'If UCase(AgentFinal(dt1.Rows(j).Item("empuid"), Now.Date, Now.Date)) = UCase(strUserID) Then
                    '    AuthFlag = "Y"
                    'End If
                    If UCase(AgentFinal(CType(dt1.Rows(j).Item("empuid"), String))) = UCase(strUserID) Then
                        AuthFlag = "Y"
                    End If

                End If

            Next
        End If

        ApproveAuth = AuthFlag

    End Function

#End Region
#Region "找出上一級全部單位"
    Public Function GetUpOrgAll(ByVal org_id As String) As String

        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        Dim UpSql As String = "SELECT PARENT_ORG_UID,ORG_TREE_LEVEL FROM ADMINGROUP WHERE ORG_UID = '" & org_id & "'"

        Dim strParentOrg As String = ""
        Dim strParentUpOrg As String = ""
        Dim strOrgTree As Integer
        Dim UpSqlAll As String

        '取得上一級單位以及本身單位階層
        db.Open()
        Dim strPer As New SqlCommand(UpSql, db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.Read() Then

            strParentOrg = RdPer("PARENT_ORG_UID").ToString()
            strOrgTree = CType(RdPer("ORG_TREE_LEVEL").ToString(), Integer)

        End If
        db.Close()

        Dim strParentOrgAll As String = strParentOrg

        Dim Org_Num As Integer = 1

        For Org_Num = 1 To (strOrgTree - 2)

            If strParentUpOrg = "" Then
                UpSqlAll = "SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & strParentOrg & "'"
            Else
                UpSqlAll = "SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & strParentUpOrg & "'"
            End If

            '取得上一級單位以及本身單位階層
            db.Open()
            Dim strPerAll As New SqlCommand(UpSqlAll, db)
            Dim RdPerAll = strPerAll.ExecuteReader()
            If RdPerAll.Read() Then

                strParentUpOrg = RdPerAll("PARENT_ORG_UID").ToString()

            End If
            db.Close()

            strParentOrgAll = strParentOrgAll & "," & strParentUpOrg

        Next

        GetUpOrgAll = strParentOrgAll

    End Function

#End Region
#Region "寫入登入者使用影印記錄的動作"
    Public Function ActionReWrite(ByVal PrintLog_ID As Integer, ByVal employee_id As String, ByVal movement As Integer, ByVal ActionScript As String) As Boolean
        Dim bl_ActionResult As Boolean = False
        Try
            Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            '新增登入者資料
            db.Open()
            Dim insActionCom As New SqlCommand("INSERT INTO P_0802 (PrintLog_ID,employee_id,movement,ActionScript) VALUES (@PrintLog_ID,@employee_id,@movement,@ActionScript)", db)
            insActionCom.Parameters.Add(New SqlParameter("PrintLog_ID", SqlDbType.Int)).Value = PrintLog_ID
            insActionCom.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = employee_id
            insActionCom.Parameters.Add(New SqlParameter("movement", SqlDbType.Int)).Value = movement
            insActionCom.Parameters.Add(New SqlParameter("ActionScript", SqlDbType.NVarChar)).Value = ActionScript

            insActionCom.ExecuteNonQuery()
            db.Close()
            db.Dispose()
            insActionCom.Dispose()
            db.Close()
            db.Dispose()
            bl_ActionResult = True
        Catch ex As Exception
            bl_ActionResult = False
        End Try
        ActionReWrite = bl_ActionResult
    End Function
#End Region
#Region "搜尋組織內Parent_Org_Uid Value by ORG_UID from ADMINGROUP"
    Public Function GetParent_ORG_Uid(ByVal org_uid As String, ByRef Parent_Org_Uid As String, ByRef ORG_Name As String) As String
        GetParent_ORG_Uid = ""
        Try
            Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            Dim SqlTxt As String = "SELECT Parent_Org_Uid,ORG_Name FROM ADMINGROUP where ORG_UID='" + org_uid + "'"
            Dim cmdORG_Uid As New SqlCommand(SqlTxt, db)
            db.Open()
            Dim Dr As SqlDataReader = cmdORG_Uid.ExecuteReader
            If Dr.Read() Then
                Parent_Org_Uid = Dr("PARENT_ORG_UID").ToString()
                ORG_Name = Dr("ORG_Name").ToString()
            End If
            db.Close()
            cmdORG_Uid.Dispose()
            Dr.Close()
        Catch ex As Exception
            GetParent_ORG_Uid = ex.Message
        End Try
    End Function
#End Region
#Region "Load出登入者的一級單位全名(中文)"
    Public Function GetNo1ORGName(ByVal org_uid As String) As String
        GetNo1ORGName = "國防部"
        Try
            Dim myList As New ArrayList()
            Dim Parent_Org_Uid As String = String.Empty
            Dim ORG_Name As String = String.Empty
            Dim Query_org_uid As String = org_uid
            Do Until Query_org_uid = "0"
                'Dim sSesult As String = GetParent_ORG_Uid(Query_org_uid, Parent_Org_Uid, ORG_Name)
                If Parent_Org_Uid <> "" And ORG_Name <> "" Then
                    If Parent_Org_Uid <> "0" Then
                        myList.Add(ORG_Name)
                    End If
                End If
                Query_org_uid = Parent_Org_Uid
            Loop
            Dim iTakeNum As Integer = 2
            If myList.Count < 2 Then
                iTakeNum = 1
            End If
            For i As Integer = myList.Count - 1 To myList.Count - iTakeNum Step -1
                GetNo1ORGName += myList.Item(i).ToString()
            Next
        Catch ex As Exception
            GetNo1ORGName = "0"
        End Try
    End Function
#End Region
#Region "身份證號遮罩"
    ''' <summary>
    ''' 身份證號遮罩
    ''' </summary>
    ''' <param name="iFrontDigits">前明碼位數</param>
    ''' <param name="iBackDigits">後明碼位數</param>
    ''' <returns>string MaskedIDNo</returns>
    ''' <remarks></remarks>
    Public Function IDMask(ByVal sIDNo As String, ByVal iFrontDigits As Integer, ByVal iBackDigits As Integer) As String
        Dim sReturn As String = ""
        Dim iTotalDigits As Integer = sIDNo.Length
        Dim sFrontDigits As String = ""
        Dim sMeddleDigits As String = ""
        Dim sBackDigits As String = ""

        If iTotalDigits = 0 Then
            MessageBox.Show("身份證號碼不可為空")
            IDMask = sReturn
        End If
        If (iTotalDigits < (iFrontDigits + iBackDigits)) Then
            MessageBox.Show("明碼總數不可超過身份證號字數")
            IDMask = sReturn
        End If

        If iFrontDigits > 0 Then
            sFrontDigits = Mid(sIDNo, 1, iFrontDigits)
        End If
        If iBackDigits > 0 Then
            sBackDigits = Right(sIDNo, iBackDigits)
        End If
        ''sMeddleDigits = Mid(sIDNo, iFrontDigits + 1, (iTotalDigits - iFrontDigits - iBackDigits))
        For i As Integer = 0 To (iTotalDigits - iFrontDigits - iBackDigits - 1) Step 1
            sMeddleDigits += "*"
        Next
        sReturn = sFrontDigits & sMeddleDigits & sBackDigits

        IDMask = sReturn
    End Function
#End Region

#Region "將檔案轉換成二進制資料"
    ''' <summary>
    ''' 將檔案轉換成二進制資料
    ''' </summary>
    ''' <param name="filepath"></param>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertFileToByteArray(ByVal filepath As String, ByVal fileName As String) As Byte()

        Dim filenamePath As String = HttpContext.Current.Server.MapPath(filepath) & fileName
        'Dim f As String = Path.GetFileName(fileName)

        Dim fs As FileStream = New FileStream(filenamePath, FileMode.Open, FileAccess.Read)
        Dim br As BinaryReader = New BinaryReader(fs)
        Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
        br.Close()
        fs.Close()
        Return bytes
    End Function

#End Region

#Region "將二進制資料轉換成檔案"
    ''' <summary>
    ''' 將二進制資料轉換成檔案
    ''' </summary>
    ''' <param name="filepath"></param>
    ''' <param name="fileName"></param>
    ''' <param name="data"></param>
    ''' <remarks></remarks>
    Public Sub ConvertByteArrayToFile(ByVal filepath As String, ByVal fileName As String, ByVal data As Byte())

        Dim filenamePath As String = HttpContext.Current.Server.MapPath(filepath) & fileName

        Dim f As FileStream
        f = New FileStream(filenamePath, FileMode.Create)
        f.Write(data, 0, data.Length)
        f.Close()

    End Sub
#End Region

#Region "檢測含有中文字符串的實際長度"
    ''' <summary> 
    ''' 檢測含有中文字符串的實際長度 
    ''' </summary> 
    ''' <param name="str">字符串</param> 
    Public Function ChineseStringLenth(ByVal str As String) As Integer
        Dim n As New ASCIIEncoding()
        Dim b As Byte() = n.GetBytes(str)
        Dim l As Integer = 0 ' l 為字符串之實際長度 
        For i As Integer = 0 To b.Length - 1
            If (b(i) = 63) Then '判斷是否為中文字或全型符號 
                l = l + 1
            End If
            l = l + 1
        Next
        ChineseStringLenth = l
    End Function
#End Region

#Region "由ID找人員名稱"
    Public Function GetUserNameByID(ByVal user_id As String) As String
        Dim DC As New SQLDBControl()
        Dim DR As SqlDataReader
        Dim strSql As String
        Dim sReturn As String = ""
        strSql = "SELECT EMP_CHINESE_NAME FROM EMPLOYEE WHERE EMPLOYEE_ID='" & user_id & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("EMP_CHINESE_NAME").ToString()
            End If
        Else
            sReturn = ""
        End If
        DC.Dispose()
        GetUserNameByID = sReturn
    End Function
#End Region

#Region "由ID找組織名稱"
    Public Function GetOrgNameByID(ByVal org_id As String) As String
        Dim DC As New SQLDBControl()
        Dim DR As SqlDataReader
        Dim strSql As String
        Dim sReturn As String = ""
        strSql = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID = '" & org_id & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("ORG_NAME").ToString()
            End If
        Else
            sReturn = ""
        End If
        DC.Dispose()
        GetOrgNameByID = sReturn
    End Function
#End Region

#Region "找出同屬使用者組織樹狀下特定階層的組織代號"
    Public Function GetTopORGIDFromUserIDByLevel(ByVal sUserID As String, ByVal iLevel As String, ByVal boolIsInInclude As Boolean) As String
        Dim sReturn As String = ""
        Dim iRecentLevel As Int32 = 9999
        Dim objUser As New BasicEmployee(sUserID)
        Dim RecentORG As New BasicOrganization(sUserID, objUser.ORG_UID)

        If boolIsInInclude Then iLevel += 1
        While iRecentLevel >= iLevel
            Dim DC As New SQLDBControl
            Dim DR As SqlDataReader
            Dim strSql As String = ""

            strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + RecentORG.PARENT_ORG_UID + "'"
            DR = DC.CreateReader(strSql)
            If DR.HasRows Then
                If DR.Read Then
                    sReturn = DR("ORG_UID")
                    iRecentLevel = DR("ORG_TREE_LEVEL")
                    RecentORG = New BasicOrganization(sUserID, RecentORG.PARENT_ORG_UID)
                End If
            End If
            DC.Dispose()
        End While

        GetTopORGIDFromUserIDByLevel = sReturn
    End Function
#End Region

#Region "由組織ID找組織名稱及下層組織名稱"
    Public Sub GetOrgChildNamesByID(ByRef ORGNames As String, ByVal org_id As String, ByVal sSeparateString As String)
        Dim tool As New C_Public
        Dim DC As New SQLDBControl()
        Dim DR As SqlDataReader
        Dim strSql As String

        strSql = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID = '" & org_id & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                ORGNames += sSeparateString + DR("ORG_NAME").ToString() + sSeparateString

                Dim DC2 As New SQLDBControl
                Dim DR2 As SqlDataReader
                strSql = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" + DR("ORG_UID").ToString() + "'"
                DR2 = DC2.CreateReader(strSql)
                While DR2.Read
                    If ORGNames.Length > 0 Then ORGNames += ","

                    If Not tool.IsORGLeaf(org_id) Then
                        GetOrgChildNamesByID(ORGNames, DR2("ORG_UID").ToString(), "'")
                    End If
                End While
                DC2.Dispose()
            End If
        End If
        DC.Dispose()
    End Sub
#End Region
#Region "由人員ID找組織名稱"
    Public Function GetOrgNameByIDNo(ByVal user_id As String) As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        Dim sReturn As String = ""
        strSql = "SELECT A.EMP_CHINESE_NAME,B.ORG_NAME FROM EMPLOYEE A,ADMINGROUP B WHERE A.EMPLOYEE_ID='" & user_id & "' AND A.ORG_UID=B.ORG_UID"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("ORG_NAME").ToString()
            End If
        End If
        DC.Dispose()
        GetOrgNameByIDNo = sReturn
    End Function
#End Region

#Region "由人員ID找組織"
    Public Function GetOrgIDByIDNo(ByVal user_id As String) As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        Dim sReturn As String = ""
        strSql = "SELECT A.EMP_CHINESE_NAME,B.ORG_UID FROM EMPLOYEE A,ADMINGROUP B WHERE A.EMPLOYEE_ID='" & user_id & "' AND A.ORG_UID=B.ORG_UID"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("ORG_UID").ToString()
            End If
        End If
        DC.Dispose()
        GetOrgIDByIDNo = sReturn
    End Function
#End Region

#Region "由人員ID找職稱"
    Public Function GetADTitleByIDNo(ByVal user_id As String) As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        Dim sReturn As String = ""
        strSql = "SELECT A.AD_TITLE FROM EMPLOYEE A WHERE A.EMPLOYEE_ID='" & user_id & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("AD_TITLE").ToString()
            End If
        End If
        DC.Dispose()
        GetADTitleByIDNo = sReturn
    End Function
#End Region

#Region "取出樹狀下所有組織"
    Public Function SelectWholeTreeORG_UID(ByVal sUser_id As String, ByVal iIncludeLevel As Int32) As String
        Dim sReturn As String = ""
        Dim tool As New C_Public

        sReturn = GetWholeOrgIDs(sUser_id, ",", "'", iIncludeLevel)
        SelectWholeTreeORG_UID = sReturn
    End Function
#End Region

#Region "判斷人員是否在群組內"
    Public Function CheckUserInRoleGroup(ByVal sUser_ID As String, ByVal sRoleGroupCName As String) As Boolean
        Dim boolResult As Boolean = False
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT * FROM ROLEGROUPITEM WHERE GROUP_UID='" + GetGroupIDFromName(sRoleGroupCName) + "' AND EMPLOYEE_ID = '" + sUser_ID + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            boolResult = True
        End If
        DC.Dispose()
        Return boolResult
    End Function
#End Region

#Region "由群組名稱取得群組ID"
    Public Function GetGroupIDFromName(ByVal sName As String) As String
        Dim sResult As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT * FROM ROLEGROUP WHERE GROUP_NAME='" + sName + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sResult = DR("GROUP_UID")
            End If
        End If
        DC.Dispose()
        Return sResult
    End Function
#End Region

#Region "判斷是否為關卡內人員"
    ''' <summary>
    ''' 判斷是否為關卡內人員
    ''' </summary>
    ''' <param name="sGroupName">關卡名稱</param>
    ''' <param name="user_id">登入者ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckStepGroupEMPByName(ByVal sGroupName As String, ByVal user_id As String) As Boolean
        Dim boolReturn As Boolean = False
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT A.OBJECT_UID FROM SYSTEMOBJUSE A,SYSTEMOBJ B WHERE A.EMPLOYEE_ID = '" & user_id & "' AND B.OBJECT_NAME='" & sGroupName & "' AND A.OBJECT_UID=B.OBJECT_UID"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            boolReturn = True
        End If
        DC.Dispose()

        CheckStepGroupEMPByName = boolReturn
    End Function
    ''' <summary>
    ''' 判斷是否為關卡內人員
    ''' </summary>
    ''' <param name="sGroupName">關卡名稱</param>
    ''' <param name="user_id">登入者ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckStepGroupEMPByName(ByVal sGroupName As String, ByVal user_id As String, ByVal object_type As String) As Boolean
        Dim boolReturn As Boolean = False
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        Select Case object_type
            Case "1" ''全關卡
            Case "2" ''上一級主管
            Case "3" ''同級單位
            Case "4" ''指定群組
        End Select
        strSql = "SELECT A.OBJECT_UID FROM SYSTEMOBJUSE A,SYSTEMOBJ B WHERE A.EMPLOYEE_ID = '" & user_id & "' AND B.OBJECT_NAME='" & sGroupName & "' AND A.OBJECT_UID=B.OBJECT_UID"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            boolReturn = True
        End If
        DC.Dispose()

        CheckStepGroupEMPByName = boolReturn
    End Function
#End Region

#Region "找出所有層級關卡的user_id"
    ''' <summary>
    ''' 找出所有單位資訊官的group_uid
    ''' </summary>
    ''' <param name="sGroupName">關卡名稱</param>    
    ''' <param name="sSeparateLetter">分隔符號</param>    
    ''' <param name="sIncludeLetter">前後符號</param>    
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetWholeInformationIDsByName(ByVal sGroupName As String, ByVal sSeparateLetter As String, ByVal sIncludeLetter As String) As String
        Dim sReturn As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT A.OBJECT_UID FROM SYSTEMOBJ A WHERE A.OBJECT_NAME = '" & sGroupName & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            While DR.Read
                If sReturn.Length > 0 Then sReturn += sSeparateLetter
                sReturn += sIncludeLetter + DR("OBJECT_UID").ToString() + sIncludeLetter
            End While
        End If
        DC.Dispose()
        GetWholeInformationIDsByName = sReturn
    End Function
#End Region

#Region "找出關卡的代號ID"
    ''' <summary>
    ''' 找出關卡的ID
    ''' </summary>
    ''' <param name="sGroupName">關卡名稱</param>    
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetInformationIDByName(ByVal sGroupName As String) As String
        Dim sReturn As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT A.OBJECT_UID FROM SYSTEMOBJ A WHERE A.OBJECT_NAME = '" & sGroupName & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("OBJECT_UID").ToString()
            End If
        End If
        DC.Dispose()
        GetInformationIDByName = sReturn
    End Function
#End Region

#Region "找出特定層級下關卡的user_id"
    ''' <summary>
    ''' 找出特定單位層級下所有單位資訊官的user_id
    ''' </summary>
    ''' <param name="sGroupName">關卡名稱</param>    
    ''' <param name="sSeparateLetter">分院符號</param>    
    ''' <param name="sIncludeLetter">前後符號</param>    
    ''' <param name="iLevel">組織最高層級</param>  
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetUnitIDsByStepName(ByVal sUserID As String, ByVal sGroupName As String, ByVal sSeparateLetter As String, ByVal sIncludeLetter As String, ByVal iLevel As Integer) As String
        Dim sReturn As String = ""
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        Dim bTopORG As Boolean
        Dim sUserORGID As String = GetOrgIDByIDNo(sUserID)

        bTopORG = False
        While Not bTopORG
            bTopORG = IsTopORG(sUserORGID, 3)
        End While
        Dim OrgMembers As String = ""
        GetSpecificMembers(sGroupName, ",", "", OrgMembers)
        Dim arrOrgMembers As String() = Split(OrgMembers, ",")

        GetMatchOrgUser(sUserORGID, arrOrgMembers, ",", sReturn)


        'strSql = "SELECT A.OBJECT_UID FROM SYSTEMOBJ A,SYSTEMOBJUSE B WHERE A.OBJECT_UID=B.OBJECT_UID AND A.OBJECT_NAME = '" & sGroupName & "' AND B.EMPLOYEE_ID='" + sUserID + "'"


        'DR = DC.CreateReader(strSql)
        'If DR.HasRows Then
        '    While DR.Read
        '        If sReturn.Length > 0 Then sReturn += sSeparateLetter
        '        sReturn += sIncludeLetter + DR("OBJECT_UID").ToString() + sIncludeLetter
        '    End While
        'End If

        GetUnitIDsByStepName = sReturn
    End Function

    ''' <summary>
    ''' 比對關卡人員與使用者發單人員是否同屬一個二級單位
    ''' </summary>
    ''' <param name="sUserOrgID">使用者所屬二級單位</param>
    ''' <param name="arrCompare">比對人員</param>
    ''' <param name="sReturn">同屬二級單位關卡人員</param>
    ''' <remarks></remarks>
    Public Sub GetMatchOrgUser(ByVal sUserOrgID As String, ByVal arrCompare As String(), ByVal sSeparateCharacter As String, ByRef sReturn As String)
        'If Not IsORGLeaf(sUserOrgID) Then
        '    Dim DC As New SQLDBControl
        '    Dim DR As SqlDataReader
        '    Dim strSql As String = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" + sUserOrgID + "'"
        '    DR = DC.CreateReader(strSql)
        '    If DR.HasRows Then
        '        While DR.Read()
        '            Dim DC2 As New SQLDBControl
        '            Dim DR2 As SqlDataReader
        '            strSql = "SELECT B.EMPLOYEE_ID,B.EMP_CHINESE_NAME FROM ADMINGROUP A,EMPLOYEE B WHERE A.ORG_UID='" & DR("ORG_UID").ToString() & "' AND A.ORG_UID=B.ORG_UID ORDER BY B.EMP_CHINESE_NAME"
        '            DR2 = DC2.CreateReader(strSql)
        '            If DR2.HasRows Then
        '                While DR2.Read()
        '                    If arrCompare.Contains(DR2("EMPLOYEE_ID").ToString()) Then
        '                        If sReturn.Length > 0 Then sReturn += sSeparateCharacter
        '                        sReturn += DR2("EMPLOYEE_ID").ToString()
        '                    End If
        '                End While
        '            End If
        '            DC2.Dispose()

        '            If Not IsORGLeaf(DR("ORG_UID").ToString()) Then
        '                GetMatchOrgUser(DR("ORG_UID").ToString(), arrCompare, sSeparateCharacter, sReturn)
        '            End If
        '            'If Not IsORGLeaf(DR("ORG_UID").ToString()) Then
        '            '    GetMatchOrgUser(DR("ORG_UID").ToString(), arrCompare, sReturn)
        '            'Else
        '            '    Dim OrgMembers As String = ""
        '            '    GetMembers(DR("ORG_UID").ToString(), ",", "", OrgMembers)
        '            '    Dim arrOrgMembers As String() = Split(OrgMembers, ",")
        '            '    For Each s As String In arrCompare
        '            '        If arrOrgMembers.Contains(s) Then
        '            '            If sReturn.Length > 0 Then sReturn += ","
        '            '            sReturn += s
        '            '        End If
        '            '    Next
        '            'End If
        '        End While
        '    End If
        '    DC.Dispose()
        'End If
    End Sub
#End Region

#Region "取得簽核者上一級簽核主管資料"
    ''' <summary>
    ''' 取得簽核者上一級簽核主管資料
    ''' 將簽核者上一層組織人員視為上一級主管，由此模組取出所有可簽核的人員ID及中文姓名
    ''' </summary>
    ''' <param name="sUserID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSuperiors(ByVal sUserID As String) As ArrayList
        Dim arrReturn As ArrayList = New ArrayList
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "select employee_id,emp_chinese_name from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id = '" + sUserID.Trim() + "'))"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            While DR.Read
                arrReturn.Add(New DictionaryEntry(DR("employee_id"), DR("emp_chinese_name")))
            End While
        End If
        DC.Dispose()
        Return arrReturn
    End Function
#End Region

#Region "傳回Now的中華民國日期時間"
    ''' <summary>
    ''' 回傳Now的中華民國字串
    ''' </summary>
    ''' <param name="sType">類別 1:只有日期 2:日期+時間 3:含中國民國日期 4:中華民國日期+時間</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetChineseDate(ByVal sType As String) As String
        Dim sReturn As String = ""
        Select Case sType
            Case "1"
                sReturn = Now.Year - 1911 & "年" & Now.Month & "月" & Now.Day & "日"
            Case "2"
                sReturn = Now.Year - 1911 & "年" & Now.Month & "月" & Now.Day & "日 " & Now.Hour & "時" & Now.Minute & "分"
            Case "3"
                sReturn = "中華民國" & Now.Year - 1911 & "年" & Now.Month & "月" & Now.Day & "日"
            Case "4"
                sReturn = "中華民國" & Now.Year - 1911 & "年" & Now.Month & "月" & Now.Day & "日 " & Now.Hour & "時" & Now.Minute & "分"
        End Select
        GetChineseDate = sReturn
    End Function
#End Region

#Region "判斷組織是否有下層組織"
    Public Function IsORGLeaf(ByVal sORG_UID As String) As Boolean
        Dim boolReturn = True
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        strSql = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" + sORG_UID + "'"
        DR = DC.CreateReader(strSql)
        If DR.Read Then
            boolReturn = False
        End If
        DC.Dispose()
        IsORGLeaf = boolReturn
    End Function
#End Region

#Region "取出使用者同屬組織所屬的所有組織代號"
    ''' <summary>
    ''' 取出使用者同屬組織所屬的所有組織代號
    ''' </summary>
    ''' <param name="sUser_ID"></param>
    ''' <param name="sSeparateCharacter"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSingleTreeOrgIDs(ByVal sUser_ID As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String) As String
        Dim sReturn As String = ""
        Dim sORG_UID As String = ""
        Dim sParentOrg_UID As String = ""
        Dim strSql As String = ""
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader

        If sUser_ID.Length > 0 Then
            DC = New SQLDBControl
            strSql = "SELECT * FROM EMPLOYEE A LEFT JOIN ADMINGROUP B ON A.ORG_UID=B.ORG_UID WHERE EMPLOYEE_ID='" & sUser_ID & "'"
            DR = DC.CreateReader(strSql)
            If DR.Read Then
                sParentOrg_UID = sIncludeCharacter + DR("ORG_UID").ToString() + sIncludeCharacter
                GetStepParentOrgIDs(DR("PARENT_ORG_UID").ToString(), sSeparateCharacter, sIncludeCharacter, sParentOrg_UID)
            End If
            DC.Dispose()
        End If

        sReturn = sORG_UID + "," + sParentOrg_UID
        GetSingleTreeOrgIDs = sReturn
    End Function

    Public Function GetWholeOrgIDs(ByVal sUser_ID As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String, ByVal iIncludeLevel As Int32) As String
        Dim sReturn As String = ""
        Dim sORG_UID As String = ""
        Dim sParentOrg_UID As String = ""
        Dim strSql As String = ""
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader

        If sUser_ID.Length > 0 Then
            DC = New SQLDBControl
            strSql = "SELECT * FROM EMPLOYEE A LEFT JOIN ADMINGROUP B ON A.ORG_UID=B.ORG_UID WHERE EMPLOYEE_ID='" & sUser_ID & "'"
            DR = DC.CreateReader(strSql)
            If DR.Read Then
                sORG_UID = DR("ORG_UID").ToString()
                While Not IsTopORG(sORG_UID, iIncludeLevel)

                End While

                'sORG_UID = sIncludeCharacter + sORG_UID + sIncludeCharacter
                'GetStepParentOrgIDs(sORG_UID, sSeparateCharacter, sIncludeCharacter, sParentOrg_UID)
                GetStepChildOrgIDs(sORG_UID.ToString(), sSeparateCharacter, sIncludeCharacter, sParentOrg_UID)
            End If
            DC.Dispose()
        End If

        sReturn = sIncludeCharacter + sORG_UID + sIncludeCharacter + IIf(sParentOrg_UID.Length > 0, "," + sParentOrg_UID, "")
        GetWholeOrgIDs = sReturn
    End Function

    Public Function GetWholeOrgIDs(ByVal sUser_ID As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String) As String
        Dim sReturn As String = ""
        Dim sORG_UID As String = ""
        Dim sParentOrg_UID As String = ""
        Dim strSql As String = ""
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader

        If sUser_ID.Length > 0 Then
            DC = New SQLDBControl
            strSql = "SELECT * FROM EMPLOYEE A LEFT JOIN ADMINGROUP B ON A.ORG_UID=B.ORG_UID WHERE EMPLOYEE_ID='" & sUser_ID & "'"
            DR = DC.CreateReader(strSql)
            If DR.Read Then
                sORG_UID = DR("ORG_UID").ToString()
                While Not IsTopORG(sORG_UID, 3)

                End While

                'sORG_UID = sIncludeCharacter + sORG_UID + sIncludeCharacter
                'GetStepParentOrgIDs(sORG_UID, sSeparateCharacter, sIncludeCharacter, sParentOrg_UID)
                GetStepChildOrgIDs(sORG_UID.ToString(), sSeparateCharacter, sIncludeCharacter, sParentOrg_UID)
            End If
            DC.Dispose()
        End If

        sReturn = sIncludeCharacter + sORG_UID + sIncludeCharacter + "," + sParentOrg_UID
        GetWholeOrgIDs = sReturn
    End Function

    Public Function GetWholeOrgIncludeMembersIDs(ByVal sUser_ID As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String, ByVal iRestrictLevel As Integer) As String
        Dim sReturn As String = ""
        Dim sORG_UID As String = ""
        Dim sParentOrg_UID As String = ""
        Dim strSql As String = ""
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader

        'If sUser_ID.Length > 0 Then
        'DC = New SQLDBControl
        'strSql = "SELECT * FROM EMPLOYEE A LEFT JOIN ADMINGROUP B ON A.ORG_UID=B.ORG_UID WHERE EMPLOYEE_ID='" & sUser_ID & "'"
        'DR = DC.CreateReader(strSql)
        'If DR.Read Then
        '    sORG_UID = sIncludeCharacter + DR("ORG_UID").ToString() + sIncludeCharacter
        '    GetStepParentOrgIncludeMembersIDs(DR("PARENT_ORG_UID").ToString(), sSeparateCharacter, sIncludeCharacter, sParentOrg_UID, iRestrictLevel)
        '    GetStepChildOrgIDs(DR("ORG_UID").ToString(), sSeparateCharacter, "'", sParentOrg_UID)
        'End If
        'DC.Dispose()
        'End If

        'sReturn = sORG_UID + "," + sParentOrg_UID
        sReturn = EMPRootORGID(sUser_ID, 3)
        GetWholeOrgIncludeMembersIDs = sReturn
    End Function

    Private Sub GetStepParentOrgIDs(ByVal sParentOrg As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String, ByRef sOrgIDs As String)
        Dim strSql As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID=" & sParentOrg
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                If sOrgIDs.Length > 0 Then sOrgIDs += sSeparateCharacter
                sOrgIDs += sIncludeCharacter + DR("ORG_UID").ToString() + sIncludeCharacter
                GetStepParentOrgIDs(DR("PARENT_ORG_UID").ToString(), sSeparateCharacter, sIncludeCharacter, sOrgIDs)
            End If
        End If
        DC.Dispose()
    End Sub

    Private Sub GetStepChildOrgIDs(ByVal sParentOrg As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String, ByRef sOrgIDs As String)
        Dim strSql As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        strSql = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" & sParentOrg & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            While DR.Read()
                If sOrgIDs.Length > 0 Then sOrgIDs += sSeparateCharacter
                sOrgIDs += sIncludeCharacter + DR("ORG_UID").ToString() + sIncludeCharacter
                GetStepChildOrgIDs(DR("ORG_UID").ToString(), sSeparateCharacter, sIncludeCharacter, sOrgIDs)
            End While
        End If
        DC.Dispose()
    End Sub
    ''' <summary>
    ''' 依關卡名稱取得特定關卡人員所有的ID
    ''' </summary>
    ''' <param name="sStepName">關卡名稱</param>
    ''' <param name="sSeparateCharacter">分隔符號</param>
    ''' <param name="sIncludeCharacter">附加符號</param>
    ''' <param name="sUserIDs">關卡人員ID</param>
    ''' <remarks></remarks>
    Public Sub GetSpecificMembers(ByVal sStepName As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String, ByRef sUserIDs As String)
        Dim strSql As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        strSql = "SELECT B.EMPLOYEE_ID FROM SYSTEMOBJ A,SYSTEMOBJUSE B WHERE A.OBJECT_UID=B.OBJECT_UID AND A.OBJECT_NAME='" + sStepName + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            While DR.Read()
                If sUserIDs.Length > 0 Then sUserIDs += sSeparateCharacter
                sUserIDs += sIncludeCharacter + DR("EMPLOYEE_ID").ToString() + sIncludeCharacter
            End While
        End If
        DC.Dispose()
    End Sub

    Private Sub GetStepParentOrgIncludeMembersIDs(ByVal sParentOrg As String, ByVal sSeparateCharacter As String, ByVal sIncludeCharacter As String, ByRef sOrgIDs As String, ByVal iLevel As Integer)
        Dim strSql As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" & sParentOrg & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                If sOrgIDs.Length > 0 Then sOrgIDs += sSeparateCharacter
                sOrgIDs += sIncludeCharacter + DR("ORG_UID").ToString() + sIncludeCharacter
                'GetMembers(DR("ORG_UID").ToString(), sSeparateCharacter, sIncludeCharacter, sOrgIDs)
                If DR("ORG_TREE_LEVEL") > iLevel Then GetStepParentOrgIncludeMembersIDs(DR("PARENT_ORG_UID").ToString(), sSeparateCharacter, sIncludeCharacter, sOrgIDs, iLevel)
            End If
        End If
        DC.Dispose()
    End Sub

    ''' <summary>   
    ''' 將部門資料做成TreeView，節點連結不動作(沒有postback)。  
    ''' 無回傳值，但需傳址。
    ''' </summary>   
    ''' <param name="ParentNode">要放入資料的TreeView。</param>   
    ''' <param name="D_ID">部門ID。</param>   
    Public Sub GetMembersFromORGToTreeNoPB(ByRef ParentNode As TreeNode, ByVal D_ID As String)
        Dim DC As New SQLDBControl()
        Dim DR As SqlDataReader
        Dim xmlTreeNode As TreeNode
        Dim strSql As String = "SELECT B.EMPLOYEE_ID,B.EMP_CHINESE_NAME FROM ADMINGROUP A,EMPLOYEE B WHERE A.ORG_UID='" & D_ID & "' AND A.ORG_UID=B.ORG_UID ORDER BY B.EMP_CHINESE_NAME"

        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            'xmlTreeNode = New TreeNode
            'xmlTreeNode.Text = "全部"
            'xmlTreeNode.Value = "ORG_" + D_ID
            'xmlTreeNode.SelectAction = TreeNodeSelectAction.Select
            'xmlTreeNode.NavigateUrl = "javascript:void(0)"

            'ParentNode.ChildNodes.Add(xmlTreeNode)
            While (DR.Read)
                xmlTreeNode = New TreeNode
                xmlTreeNode.Text = DR("EMP_CHINESE_NAME").ToString()
                xmlTreeNode.Value = "EMP_" + DR("EMPLOYEE_ID").ToString()
                xmlTreeNode.SelectAction = TreeNodeSelectAction.Select
                xmlTreeNode.NavigateUrl = "javascript:void(0)"

                ParentNode.ChildNodes.Add(xmlTreeNode)
            End While
        End If
        DC.Dispose()
    End Sub

    Public Function EMPRootORGID(ByVal EMPLOYEE_ID As String, ByVal MaxLevel As Integer) As String
        Dim sReturn As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        strSql = "SELECT * FROM EMPLOYEE WHERE EMPLOYEE_ID='" + EMPLOYEE_ID + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("ORG_UID").ToString()
            End If
        End If
        While (Not IsTopORG(sReturn, MaxLevel))

        End While
        DC.Dispose()

        EMPRootORGID = sReturn
    End Function

    Public Function IsTopORG(ByRef ORG_UID As String, ByVal MaxLevel As Integer) As Boolean
        Dim boolReturn As Boolean = False
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        'Dim sParent_Org_UID As String = ""
        Dim iLevel As Integer
        strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + ORG_UID + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                'sParent_Org_UID = DR("PARENT_ORG_UID")
                iLevel = CType(DR("ORG_TREE_LEVEL"), Integer)
                If MaxLevel >= iLevel Then
                    boolReturn = True
                Else
                    ORG_UID = DR("PARENT_ORG_UID").ToString()
                End If

            End If
        End If
        DC.Dispose()
        IsTopORG = boolReturn
    End Function
#End Region

#Region "找出特定層級組織的關卡簽核人員數量"
    Public Function GetOrgStepSigner(ByVal sSignerName As String, ByVal iLevel As Integer) As Integer
        'Dim iReturn As Integer = 0

        'Dim arrInformationSigner As String() = Split(GetWholeInformationIDsByName("單位資訊官", ",", "").Replace("'", ""), ",")
        'Dim strSql As String = "SELECT ORG_UID FROM ADMINGROUP WHERE ORG_UID IN (" + GetWholeOrgIDs(HttpContext.Current.Session("user_id").ToString(), ",", "'") + ")"
        'Dim DC As New SQLDBControl
        'Dim DR As SqlDataReader

        'DR = DC.CreateReader(strSql)
        'If DR.HasRows Then
        '    While DR.Read()
        '        Dim DC2 As New SQLDBControl
        '        Dim DR2 As SqlDataReader
        '        strSql = "SELECT * FROM EMPLOYEE WHERE ORG_UID='" + DR("ORG_UID").ToString() + "'"
        '        DR2 = DC2.CreateReader(strSql)
        '        While DR2.Read()
        '            If arrInformationSigner.Contains(DR2("employee_id").ToString()) Then
        '                iReturn += 1
        '            End If
        '        End While
        '        DC2.Dispose()
        '    End While
        'End If

        'DC.Dispose()
        'GetOrgStepSigner = iReturn

    End Function
#End Region

#Region "找出特定案件最終簽核紀錄的批核意見"
    Public Function GetFinalCaseComment(ByVal eformsn As String) As String
        Dim sReturn As String = ""
        Dim strSql As String = "SELECT TOP 1 * FROM FLOWCTL WHERE EFORMSN='" + eformsn + "' ORDER BY FLOWSN DESC"
        Dim DC = New SQLDBControl
        Dim DR As SqlDataReader
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("COMMENT").ToString()
            End If
        End If
        DC.Dispose()
        GetFinalCaseComment = sReturn
    End Function
#End Region

#Region "找出特定案件的最終狀態"
    Public Function GetFinalStatus(ByVal eformsn As String) As String
        Dim sReturn As String = ""
        Dim strSql As String = "SELECT TOP 1 * FROM FLOWCTL WHERE EFORMSN='" + eformsn + "' ORDER BY FLOWSN DESC"
        Dim DC = New SQLDBControl
        Dim DR As SqlDataReader
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("GONOGO").ToString()
            End If
        End If
        DC.Dispose()
        GetFinalStatus = sReturn
    End Function
#End Region

#Region "判斷特定群組是否符合簽核的群組"
    Public Function CheckGroupSigner(ByVal sGroupName As String, ByVal eformsn As String) As Boolean
        Dim boolReturn As Boolean = False
        Dim Dc As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String

        strSql = "SELECT TOP 1 * FROM FLOWCTL WHERE EFORMSN='" + eformsn + "' AND HDDATE IS NULL AND GONOGO='?'"
        DR = Dc.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                boolReturn = sGroupName = DR("group_name").ToString()
            End If
        End If
        Dc.Dispose()
        CheckGroupSigner = boolReturn
    End Function
#End Region
#Region "找出關卡的種類"
    Public Function GetObjectTypeFromStep(ByVal object_name As String) As String
        Dim sReturn As String = ""
        Dim strSql As String = "SELECT TOP 1 * FROM SYSTEMOBJ WHERE OBJECT_NAME='" + object_name + "'"
        Dim DC = New SQLDBControl
        Dim DR As SqlDataReader
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("OBJECT_TYPE").ToString()
            End If
        End If
        DC.Dispose()
        GetObjectTypeFromStep = sReturn
    End Function
#End Region

#Region "判斷資訊設備維修案件是否正在修繕中"
    Public Function IsIPRepairing(ByVal eformsn As String) As Boolean
        Dim sReturn As Boolean = False
        Dim strSql As String = "SELECT TOP 2 * FROM FLOWCTL WHERE EFORMSN='" + eformsn + "' ORDER BY FLOWSN"
        Dim DC = New SQLDBControl
        Dim DR As SqlDataReader
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = True
            End If
        End If
        DC.Dispose()
        IsIPRepairing = sReturn
    End Function
#End Region

#Region "確認使用者是否為資訊報修派工單位主管"
    Public Function IsFixmanSupervisor(ByVal sUserID As String) As Boolean
        Dim boolReturn As Boolean = False
        Dim strSql As String = ""
        Dim i As Int32 = 0
        Dim arrFixman() As String = {""}
        Dim arrFixmanSupervisor() As String = {""}
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        ''取出資訊設備維修所有維修人員名單
        strSql = "SELECT FIXIDNO FROM P_11 GROUP BY FIXIDNO"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            i = 0
            While DR.Read()
                ReDim Preserve arrFixman(i)
                arrFixman(i) = (DR("FIXIDNO").ToString())
                i = i + 1
            End While
        End If
        DC.Dispose()

        ''維修人員主管名單
        If Not IsNothing(arrFixman) AndAlso arrFixman.Length > 0 Then
            i = 0
            For Each s In arrFixman
                Dim PreOrg As New PreBasicOrganization(s)
                strSql = "SELECT * FROM EMPLOYEE WHERE ORG_UID='" + PreOrg.ORG_UID + "'"
                DC = New SQLDBControl()
                DR = DC.CreateReader(strSql)
                If DR.HasRows Then
                    While DR.Read()
                        ReDim Preserve arrFixmanSupervisor(i)
                        arrFixmanSupervisor(i) = DR("EMPLOYEE_ID").ToString()
                        i = i + 1
                    End While
                End If
            Next
            If arrFixmanSupervisor.Contains(sUserID) Then
                boolReturn = True
            End If
        End If

        IsFixmanSupervisor = boolReturn
    End Function
#End Region

#Region "找出表單相對應的主資料表名稱"
    Public Function GetPrimaryTableName(ByVal eformid As String) As String
        Dim sReturn = ""
        Dim DC As New SQLDBControl
        Dim strSql As String = ""
        Dim DR As SqlDataReader

        strSql = "SELECT PRIMARY_TABLE FROM EFORMS WHERE EFORMID = '" & eformid & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("PRIMARY_TABLE")
            End If
        End If
        DC.Dispose()
        GetPrimaryTableName = sReturn
    End Function
#End Region

#Region "由群組名稱找出群組的代號ID"
    ''' <summary>
    ''' 找出群組的ID
    ''' </summary>
    ''' <param name="sRoleGroupName">群組名稱</param>    
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRoleGroupIDByName(ByVal sRoleGroupName As String) As String
        Dim sReturn As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT A.Group_Uid FROM ROLEGROUP A WHERE A.Group_Name = '" & sRoleGroupName & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("Group_Uid").ToString()
            End If
        End If
        DC.Dispose()
        GetRoleGroupIDByName = sReturn
    End Function
#End Region

#Region "找出特定群組內的user_id"
    ''' <summary>
    ''' 找出特定群組內的user_id
    ''' </summary>
    ''' <param name="sRoleGroupName">群組名稱</param>    
    ''' <param name="sSeparateLetter">分隔符號</param>    
    ''' <param name="sIncludeLetter">前後符號</param>    
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetWholeRoleGroupIDsByName(ByVal sRoleGroupName As String, ByVal sSeparateLetter As String, ByVal sIncludeLetter As String) As String
        Dim sReturn As String = ""
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT B.employee_id FROM ROLEGROUP A,ROLEGROUPITEM B WHERE A.GROUP_UID=B.GROUP_UID AND A.GROUP_NAME = '" & sRoleGroupName & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            While DR.Read
                If sReturn.Length > 0 Then sReturn += sSeparateLetter
                sReturn += sIncludeLetter + DR("employee_id").ToString() + sIncludeLetter
            End While
        End If
        DC.Dispose()
        GetWholeRoleGroupIDsByName = sReturn
    End Function
#End Region

#Region "判斷是否為群組內人員"
    ''' <summary>
    ''' 判斷是否為群組內人員
    ''' </summary>
    ''' <param name="sRoleGroupName">群組名稱</param>
    ''' <param name="user_id">登入者ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckRoleGroupEMPByName(ByVal sRoleGroupName As String, ByVal user_id As String) As Boolean
        Dim boolReturn As Boolean = False
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""

        strSql = "SELECT B.EMPLOYEE_ID FROM ROLEGROUP A,ROLEGROUPITEM B WHERE B.EMPLOYEE_ID = '" & user_id & "' AND A.GROUP_NAME='" & sRoleGroupName & "' AND A.GROUP_UID=B.GROUP_UID"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            boolReturn = True
        End If
        DC.Dispose()

        CheckRoleGroupEMPByName = boolReturn
    End Function
#End Region

#Region "網頁字串加密"
    Public Shared Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function
#End Region

#Region "網頁字串解密"
    Public Shared Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        cipherText = cipherText.Replace(" ", "+")
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function
#End Region

    ''' <summary>
    ''' 轉換表單狀態代號
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FunStatus(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr = str

            If tmpStr = "-" Then
                tmpStr = "申請"
            ElseIf tmpStr = "F" Then
                tmpStr = "同意修繕"
            ElseIf tmpStr = "N" Then
                tmpStr = "不修繕"
            ElseIf tmpStr = "C" Then
                tmpStr = "完工"
            ElseIf tmpStr = "0" Then
                tmpStr = "駁回"
            ElseIf tmpStr = "1" Then
                tmpStr = "送件"
            ElseIf tmpStr = "?" Then
                tmpStr = "審核中"
            ElseIf tmpStr = "E" Then
                tmpStr = "完成"
            ElseIf tmpStr = "G" Then
                tmpStr = "補登"
            ElseIf tmpStr = "B" Or tmpStr = "X" Then
                tmpStr = "申請者撤銷"
            ElseIf tmpStr = "R" Then
                tmpStr = "重新分派"
            ElseIf tmpStr = "T" Then
                tmpStr = "呈轉"
            ElseIf tmpStr = "2" Then ''資訊設備報修專用狀態，用在指定給管制單位重新分派維修單位
                tmpStr = "重派"
            Else
                tmpStr = "未知"
            End If

            FunStatus = tmpStr

        Catch ex As Exception
            FunStatus = ""
        End Try
    End Function
End Class

Public Class BasicEmployee
    Inherits SQLDBControl
    Implements IDisposable

    Private _empuid As Integer = 0              '人員識別碼
    Private _employee_id As String = ""         '人員編號
    Private _member_uid As String = ""          '身份證字號
    Private _ORG_UID As String = ""             '部門代號
    Private _empemail As String = ""            'email
    Private _emp_chinese_name As String = ""    '中文姓名
    Private _emp_english_name As String = ""    '英文姓名
    Private _password As String = ""            '登入密碼
    Private _create_date As DateTime            '建立日期
    Private _modify_date As DateTime            '修改日期
    Private _leave As String = ""               '在職與否
    Private _ErrCount As Integer = 0            '登入錯誤次數
    Private _title_id As String = ""            '職務類別代碼
    Private _TU_ID As Integer = 0               '級職名稱
    Private _ArriveDate As DateTime             '到職日
    Private _LeaveDate As DateTime              '離職日
    Private _AD_DEP As String = ""              'AD部門
    Private _AD_TITLE As String = ""            'AD職稱

    ''' <summary>
    ''' 人員識別碼
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property empuid() As Integer
        Get
            Return _empuid
        End Get
        Set(value As Integer)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 人員編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property employee_id() As String
        Get
            Return _employee_id
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 身份證字號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property member_uid() As String
        Get
            Return _member_uid
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 部門代號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_UID() As String
        Get
            Return _ORG_UID
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' email
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property empemail() As String
        Get
            Return _empemail
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 中文姓名
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property emp_chinese_name() As String
        Get
            Return _emp_chinese_name
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 英文姓名
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property emp_english_name() As String
        Get
            Return _emp_english_name
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 登入密碼
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property password() As String
        Get
            Return _password
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 建立日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property create_date() As DateTime
        Get
            Return _create_date
        End Get
        Set(value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 修改日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property modify_date() As DateTime
        Get
            Return _modify_date
        End Get
        Set(value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 在職與否
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property leave() As String
        Get
            Return _leave
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 登入錯誤次數
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ErrCount() As Integer
        Get
            Return _ErrCount
        End Get
        Set(value As Integer)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 職務類別代碼
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property title_id() As String
        Get
            Return _title_id
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 級職名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property TU_ID() As Integer
        Get
            Return _TU_ID
        End Get
        Set(value As Integer)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 到職日
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ArriveDate() As DateTime
        Get
            Return _ArriveDate
        End Get
        Set(value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 離職日
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property LeaveDate() As DateTime
        Get
            Return _LeaveDate
        End Get
        Set(value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' AD部門
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property AD_DEP() As String
        Get
            Return _AD_DEP
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' AD職稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property AD_TITLE() As String
        Get
            Return _AD_TITLE
        End Get
        Set(value As String)
            Return
        End Set
    End Property

    'Public Sub New(ByVal Server As String)
    '    MyBase.New(Server)
    'End Sub
    Public Sub New(ByVal sUser_ID As String)
        MyBase.New()
        Load(sUser_ID)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub Load(ByVal sEMP_UID As String)
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        strSql = "SELECT * FROM EMPLOYEE WHERE EMPLOYEE_ID='" + sEMP_UID + "'"
        DR = DC.CreateReader(strSql)
        If DR.Read() Then
            _empuid = CType(DR("empuid"), Integer)
            _employee_id = sEMP_UID
            _member_uid = DR("member_uid").ToString()
            _ORG_UID = DR("ORG_UID").ToString()
            _empemail = DR("empemail").ToString()
            _emp_chinese_name = DR("emp_chinese_name").ToString()
            _emp_english_name = DR("emp_english_name").ToString()
            _password = DR("password").ToString()
            If DR("create_date").ToString().Length > 0 Then _create_date = CType(DR("create_date").ToString(), Date)
            If DR("modify_date").ToString().Length > 0 Then _modify_date = CType(DR("modify_date").ToString(), Date)
            _leave = DR("leave").ToString()
            _ErrCount = CType(DR("ErrCount").ToString(), Integer)
            _title_id = DR("title_id").ToString()
            _TU_ID = CType(DR("TU_ID").ToString(), Integer)
            If DR("ArriveDate").ToString().Length > 0 Then _ArriveDate = CType(DR("ArriveDate").ToString(), Date)
            If DR("LeaveDate").ToString().Length > 0 Then _LeaveDate = CType(DR("LeaveDate").ToString(), Date)
            _AD_DEP = DR("AD_DEP").ToString()
            _AD_TITLE = DR("AD_TITLE").ToString()
        End If
        DC.Dispose()
    End Sub

End Class

''' <summary>
''' 使用者上一級組織
''' </summary>
''' <remarks></remarks>
Public Class PreBasicOrganization
    Inherits SQLDBControl
    Implements IDisposable

    Private _ORG_UID As String = ""         '組織編號
    Private _PARENT_ORG_UID As String = ""  '上一層組織編號
    Private _ORG_NAME As String = ""        '組織名稱
    Private _ORG_NAME_E As String = ""      '組織名稱(英)
    Private _ORG_TREE_LEVEL As Int32 = 0    '組織層級
    Private _ORG_KIND As Int32 = 0          '組織種類
    Private _ORG_AD As String = ""          '組織頭銜
    Private _CREATE_DATE As DateTime        '建立日期
    Private _MODIFY_DATE As DateTime        '修改日期

    ''' <summary>
    ''' 組織編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_UID() As String
        Get
            Return _ORG_UID
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 上一層組織編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property PARENT_ORG_UID() As String
        Get
            Return _PARENT_ORG_UID
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_NAME() As String
        Get
            Return _ORG_NAME
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織名稱(英)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_NAME_E() As String
        Get
            Return _ORG_NAME_E
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織層級
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_TREE_LEVEL() As Int32
        Get
            Return _ORG_TREE_LEVEL
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織種類
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_KIND() As String
        Get
            Return _ORG_KIND
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織頭銜
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_AD() As String
        Get
            Return _ORG_AD
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 建立日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property create_date() As DateTime
        Get
            Return _CREATE_DATE
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 修改日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property modify_date() As DateTime
        Get
            Return _MODIFY_DATE
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property

    'Public Sub New(ByVal Server As String)
    '    MyBase.New(Server)
    'End Sub
    Public Sub New(ByVal sUser_ID As String)
        MyBase.New()
        Load(sUser_ID)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub Load(ByVal sEMP_UID As String)
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim emp As New BasicEmployee(sEMP_UID)
        Dim sParentOrgUID As String = ""
        strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + emp.ORG_UID + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                sParentOrgUID = DR("PARENT_ORG_UID").ToString()
            End If
        End If
        DC.Dispose()

        DC = New SQLDBControl()
        strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + sParentOrgUID + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                _ORG_UID = sParentOrgUID
                _PARENT_ORG_UID = CType(IIf(IsDBNull(DR("PARENT_ORG_UID")), "", DR("PARENT_ORG_UID")), String)
                _ORG_NAME = CType(IIf(IsDBNull(DR("ORG_NAME")), "", DR("ORG_NAME")), String)
                _ORG_UID = CType(IIf(IsDBNull(DR("ORG_UID")), "", DR("ORG_UID")), String)
                _ORG_NAME_E = CType(IIf(IsDBNull(DR("ORG_NAME_E")), "", DR("ORG_NAME_E")), String)
                _ORG_TREE_LEVEL = CType(IIf(IsDBNull(DR("ORG_TREE_LEVEL")), 0, DR("ORG_TREE_LEVEL")), Int32)
                _ORG_KIND = CType(IIf(IsDBNull(DR("ORG_KIND")), 0, DR("ORG_KIND")), Int32)
                _ORG_AD = CType(IIf(IsDBNull(DR("ORG_AD")), "", DR("ORG_AD")), String)
                If DR("CREATE_DATE").ToString().Length > 0 Then _CREATE_DATE = CType(DR("CREATE_DATE").ToString(), Date)
                If DR("MODIFY_DATE").ToString().Length > 0 Then _MODIFY_DATE = CType(DR("MODIFY_DATE").ToString(), Date)
            End If
        End If
        DC.Dispose()
    End Sub
End Class

''' <summary>
''' 使用者組織
''' </summary>
''' <remarks></remarks>
Public Class BasicOrganization
    Inherits SQLDBControl
    Implements IDisposable

    Private _ORG_UID As String = ""         '組織編號
    Private _PARENT_ORG_UID As String = ""  '上一層組織編號
    Private _ORG_NAME As String = ""        '組織名稱
    Private _ORG_NAME_E As String = ""      '組織名稱(英)
    Private _ORG_TREE_LEVEL As Int32 = 0    '組織層級
    Private _ORG_KIND As Int32 = 0          '組織種類
    Private _ORG_AD As String = ""          '組織頭銜
    Private _CREATE_DATE As DateTime        '建立日期
    Private _MODIFY_DATE As DateTime        '修改日期

    ''' <summary>
    ''' 組織編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_UID() As String
        Get
            Return _ORG_UID
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 上一層組織編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property PARENT_ORG_UID() As String
        Get
            Return _PARENT_ORG_UID
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_NAME() As String
        Get
            Return _ORG_NAME
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織名稱(英)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_NAME_E() As String
        Get
            Return _ORG_NAME_E
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織層級
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_TREE_LEVEL() As Int32
        Get
            Return _ORG_TREE_LEVEL
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織種類
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_KIND() As String
        Get
            Return _ORG_KIND
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 組織頭銜
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ORG_AD() As String
        Get
            Return _ORG_AD
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 建立日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property create_date() As DateTime
        Get
            Return _CREATE_DATE
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 修改日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property modify_date() As DateTime
        Get
            Return _MODIFY_DATE
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property

    'Public Sub New(ByVal Server As String)
    '    MyBase.New(Server)
    'End Sub

    Public Sub New(ByVal sUser_ID As String, ByVal sOrg_UID As String)
        MyBase.New()
        Load(sUser_ID, sOrg_UID)
    End Sub

    Public Sub New(ByVal sUser_ID As String)
        MyBase.New()
        Load(sUser_ID)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub Load(ByVal sEMP_UID As String, ByVal sOrg_UID As String)
        Try
            Dim strSql As String
            Dim DC As New SQLDBControl
            Dim DR As SqlDataReader
            strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + sOrg_UID + "'"
            DR = DC.CreateReader(strSql)
            If DR.HasRows Then
                If DR.Read() Then
                    _ORG_UID = sOrg_UID
                    _PARENT_ORG_UID = CType(IIf(IsDBNull(DR("PARENT_ORG_UID")), "", DR("PARENT_ORG_UID")), String)
                    _ORG_NAME = CType(IIf(IsDBNull(DR("ORG_NAME")), "", DR("ORG_NAME")), String)
                    _ORG_UID = CType(IIf(IsDBNull(DR("ORG_UID")), "", DR("ORG_UID")), String)
                    _ORG_NAME_E = CType(IIf(IsDBNull(DR("ORG_NAME_E")), "", DR("ORG_NAME_E")), String)
                    _ORG_TREE_LEVEL = CType(IIf(IsDBNull(DR("ORG_TREE_LEVEL")), 0, DR("ORG_TREE_LEVEL")), Int32)
                    _ORG_KIND = CType(IIf(IsDBNull(DR("ORG_KIND")), 0, DR("ORG_KIND")), Int32)
                    _ORG_AD = CType(IIf(IsDBNull(DR("ORG_AD")), "", DR("ORG_AD")), String)
                    If Not IsDBNull(DR("CREATE_DATE")) Then _CREATE_DATE = CType(DR("CREATE_DATE").ToString(), Date)
                    If Not IsDBNull(DR("MODIFY_DATE")) Then _MODIFY_DATE = CType(DR("MODIFY_DATE").ToString(), Date)
                End If
            End If
            DC.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub Load(ByVal sEMP_UID As String)
        Try
            Dim strSql As String
            Dim DC As New SQLDBControl
            Dim DR As SqlDataReader
            Dim emp As New BasicEmployee(sEMP_UID)
            Dim sParentOrgUID As String = ""
            strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + emp.ORG_UID + "'"
            DR = DC.CreateReader(strSql)
            If DR.HasRows Then
                If DR.Read() Then
                    sParentOrgUID = DR("PARENT_ORG_UID").ToString()
                End If
            End If
            DC.Dispose()

            DC = New SQLDBControl()
            strSql = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + sParentOrgUID + "'"
            DR = DC.CreateReader(strSql)
            If DR.HasRows Then
                If DR.Read() Then
                    _ORG_UID = sParentOrgUID
                    _PARENT_ORG_UID = CType(IIf(IsDBNull(DR("PARENT_ORG_UID")), "", DR("PARENT_ORG_UID")), String)
                    _ORG_NAME = CType(IIf(IsDBNull(DR("ORG_NAME")), "", DR("ORG_NAME")), String)
                    _ORG_UID = CType(IIf(IsDBNull(DR("ORG_UID")), "", DR("ORG_UID")), String)
                    _ORG_NAME_E = CType(IIf(IsDBNull(DR("ORG_NAME_E")), "", DR("ORG_NAME_E")), String)
                    _ORG_TREE_LEVEL = CType(IIf(IsDBNull(DR("ORG_TREE_LEVEL")), 0, DR("ORG_TREE_LEVEL")), Int32)
                    _ORG_KIND = CType(IIf(IsDBNull(DR("ORG_KIND")), 0, DR("ORG_KIND")), Int32)
                    _ORG_AD = CType(IIf(IsDBNull(DR("ORG_AD")), "", DR("ORG_AD")), String)
                    If Not IsDBNull(DR("CREATE_DATE")) Then _CREATE_DATE = CType(DR("CREATE_DATE").ToString(), Date)
                    If Not IsDBNull(DR("MODIFY_DATE")) Then _MODIFY_DATE = CType(DR("MODIFY_DATE").ToString(), Date)
                End If
            End If
            DC.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
