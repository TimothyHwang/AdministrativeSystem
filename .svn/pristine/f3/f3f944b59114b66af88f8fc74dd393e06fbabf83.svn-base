Imports System.Data.SqlClient

Partial Class Source_00_MOA00020
    Inherits System.Web.UI.Page
    Dim conn As New C_SQLFUN
    Dim connstr As String = conn.G_conn_string
    Dim db As New SqlConnection(connstr)

    Public PrePageValue, PageFolder, PagePathAll, streformid, eformsn, read_only As String
    Dim user_id As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

        '開啟連線
        Dim db As New SqlConnection(connstr)

        PrePageValue = Trim(Request.QueryString("x"))       'x抓上頁值
        PageFolder = Mid(PrePageValue, 4, 2)
        PagePathAll = PageFolder & "/" & PrePageValue

        streformid = Trim(Request.QueryString("y"))       'eformid

        eformsn = Trim(Request.QueryString("eformsn"))       'eformsn

        read_only = Trim(Request.QueryString("read_only"))       'read_only

        'eformsn空白則為新表單
        If eformsn = "" Then
            Dim C_Public As New C_Public
            'eformsn = C_Public.randstr(16)            '建立唯一的eformsn
            eformsn = C_Public.CreateNewEFormSN()      '加入重覆防呆功能 20130710 paul

            'Dim EndFlag As Boolean = False

            'While EndFlag = False

            '    '判斷申請單號是否重複
            '    db.Open()
            '    Dim strTopPer As New SqlCommand("select flowsn from flowctl where eformsn = '" & eformsn & "'", db)
            '    Dim RdTopPer = strTopPer.ExecuteReader()
            '    If RdTopPer.read() Then
            '        EndFlag = False
            '        'eformsn = C_Public.randstr(16)            '建立唯一的eformsn
            '        eformsn = C_Public.CreateNewEFormSN()      '加入重覆防呆功能 20130710 paul
            '    Else
            '        EndFlag = True
            '        eformsn = eformsn
            '    End If
            '    db.Close()

            'End While

        Else
            eformsn = eformsn
        End If

        '列印資訊媒體攜出入申請單
        Dim fn As String = Request.QueryString("fn")

        If fn <> "" Then

            Response.Write("<script language='javascript'>")
            Response.Write("window.onload = function() {")
            Response.Write("window.location.replace('../../drs/" & fn & "');")
            Response.Write("}")
            Response.Write("</script>")

        End If

        fn = ""

        Dim streformTable As String = ""

        '找出表單資料表
        db.Open()
        Dim strTopTable As New SqlCommand("select PRIMARY_TABLE from EFORMS where eformid = '" & streformid & "'", db)
        Dim RdTopTable = strTopTable.ExecuteReader()
        If RdTopTable.read() Then
            streformTable = RdTopTable("PRIMARY_TABLE").ToString()
        End If
        db.Close()

        Dim Org_Down As New C_Public

        '判斷使用者表單權限
        If read_only = "1" Then

            Dim strTopUser As String = ""

            '判斷是否為系統管理員
            db.Open()
            Dim strTopPer As New SqlCommand("select ROLEGROUP.Group_Name from ROLEGROUP,ROLEGROUPITEM where ROLEGROUP.Group_Uid = ROLEGROUPITEM.Group_Uid AND ROLEGROUPITEM.employee_id = '" & user_id & "'  AND (ROLEGROUP.Group_Uid = 'GETK539lC0')", db)
            Dim RdTopPer = strTopPer.ExecuteReader()
            If RdTopPer.read() Then
                strTopUser = RdTopPer("Group_Name").ToString()
            End If
            db.Close()

            If strTopUser = "" Then

                Dim strflowsn As String = ""
                Dim strPAIDNO As String = ""
                Dim strAgent1 As String = ""
                Dim strAgent2 As String = ""
                Dim strAgent3 As String = ""
                Dim strEmpName As String = ""
                Dim strAgentFlag As String = ""
                Dim strSigner As String = ""

                '找尋批核者
                db.Open()
                Dim strPer As New SqlCommand("SELECT flowsn,empuid FROM flowctl WHERE empuid = '" & user_id & "' and eformsn = '" & eformsn & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    strflowsn = RdPer("flowsn").ToString()
                    strSigner = RdPer("empuid").ToString()
                End If
                db.Close()

                Dim printRecordsReportID As String = GetEFormId("影印紀錄呈核單") '取得影印紀錄呈核單ID
                Dim doorAndMeetingControlID As String = GetEFormId("門禁會議管制申請單") '取得門禁會議管制申請單ID
                If streformid = "YAqBTxRP8P" Then '請假申請單

                    '找尋登入者姓名
                    db.Open()
                    Dim strName As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
                    Dim RdName = strName.ExecuteReader()
                    If RdName.Read() Then
                        strEmpName = RdName("emp_chinese_name").ToString()
                    End If
                    db.Close()

                    '找尋代理人
                    db.Open()
                    Dim strAgentPer As New SqlCommand("SELECT PAIDNO,nAGENT1,nAGENT2,nAGENT3 FROM P_01 WHERE EFORMSN = '" & eformsn & "'", db)
                    Dim RdAgentPer = strAgentPer.ExecuteReader()
                    If RdAgentPer.Read() Then
                        strPAIDNO = RdAgentPer("PAIDNO").ToString()
                        strAgent1 = RdAgentPer("nAGENT1").ToString()
                        strAgent2 = RdAgentPer("nAGENT2").ToString()
                        strAgent3 = RdAgentPer("nAGENT3").ToString()
                    End If
                    db.Close()

                    If strEmpName = strAgent1 Or strEmpName = strAgent2 Or strEmpName = strAgent3 Or strPAIDNO = user_id Then
                        strAgentFlag = "Y"
                    End If
                ElseIf streformid = printRecordsReportID Then   '影印紀錄呈核單
                    '影印紀錄呈核單不做任何動作,相關按鈕顯示與否,在MOA08014.aspx內處理
                ElseIf streformid = doorAndMeetingControlID Then   '門禁會議管制申請單
                    '門禁會議管制申請單不做任何動作,相關按鈕顯示與否,在MOA09001.aspx內處理
                Else

                    '找尋申請人
                    db.Open()
                    Dim strAgentPer As New SqlCommand("SELECT PAIDNO FROM " & streformTable & " WHERE EFORMSN = '" & eformsn & "'", db)
                    Dim RdAgentPer = strAgentPer.ExecuteReader()
                    If RdAgentPer.Read() Then
                        strPAIDNO = RdAgentPer("PAIDNO").ToString()
                    End If
                    db.Close()
                    ''2014/7/3 Paul 修改：當瀏覽者是申請人或是簽核人時，可以檢視詳細資料
                    If (strPAIDNO = user_id) Or (strSigner = user_id) Then
                        strAgentFlag = "Y"
                    End If

                End If
                ''資訊報修需開權限給派工單位主管查詢表單，既使未簽核也需要查看表單
                If Org_Down.IsFixmanSupervisor(user_id) And streformid = GetEFormId("資訊設備維修申請單") Then
                    strAgentFlag = "Y"
                End If
                If strflowsn = "" And strAgentFlag = "" Then
                    Response.Write("權限不足")
                    Response.End()
                End If

            End If

        ElseIf read_only = "2" Then

            '沒有批核權限不可執行動作
            If (Org_Down.ApproveAuth(eformsn, user_id)) = "" Then

                Dim strflowsn As String = ""

                '找尋批核者
                db.Open()
                Dim strPer As New SqlCommand("SELECT flowsn FROM flowctl WHERE empuid = '" & user_id & "' and eformsn = '" & eformsn & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.Read() Then
                    strflowsn = RdPer("flowsn").ToString()
                End If
                db.Close()

                If strflowsn = "" Then
                    Response.Write("權限不足")
                    Response.End()
                End If

            End If

        End If

    End Sub

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        db.Open()
        Dim sqlcomm As New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db)
        GetEFormId = sqlcomm.ExecuteScalar()
        db.Close()
    End Function
End Class
