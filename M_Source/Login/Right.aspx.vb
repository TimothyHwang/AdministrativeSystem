Imports System.Data.SqlClient
Partial Class OA_Right
    Inherits System.Web.UI.Page

    Dim connstr As String = ""
    Dim conn As New C_SQLFUN
    Dim user_id, org_uid As String
    Public strApprote As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        '
        '表單未批核總筆數
        LabAppALL.Text = SQLALL(user_id, "", "1")

        '判斷批核表單筆數是否超過一筆
        'If SQLALL(user_id, "", "1") > 0 Then
        strApprote = "Show"
        'End If

        '請假表單未批核筆數
        LabApp1.Text = SQLALL(user_id, "YAqBTxRP8P", "1") + SQLALL(user_id, "5D82872F5L", "1")
        '會議室表單未批核筆數
        LabApp2.Text = SQLALL(user_id, "4rM2YFP73N", "1")
        '派車表單未批核筆數
        LabApp3.Text = SQLALL(user_id, "j2mvKYe3l9", "1")
        '房舍表單未批核筆數
        LabApp4.Text = SQLALL(user_id, "F9MBD7O97G", "1") + SQLALL(user_id, "61TY3LELYT", "1") + SQLALL(user_id, "4ZXNVRV8B6", "1")
        '會客表單未批核筆數
        LabApp5.Text = SQLALL(user_id, "U28r13D6EA", "1")
        '門禁會議管制申請單未批核筆數
        LabApp6.Text = SQLALL(user_id, "S9QR2W8X6U", "1")
        '資訊設備維修申請單未批核筆數
        LabApp7.Text = SQLALL(user_id, "BL7U2QP3IG", "1")

        '未完成表單總筆數
        LabWaitALL.Text = SQLALL(user_id, "", "2")

        '請假表單未完成表單筆數
        LabWait1.Text = SQLALL(user_id, "YAqBTxRP8P", "2") + SQLALL(user_id, "5D82872F5L", "2")
        '會議室表單未完成表單筆數
        LabWait2.Text = SQLALL(user_id, "4rM2YFP73N", "2")
        '派車表單未完成表單筆數
        LabWait3.Text = SQLALL(user_id, "j2mvKYe3l9", "2")
        '房舍表單未完成表單筆數
        LabWait4.Text = SQLALL(user_id, "F9MBD7O97G", "2") + SQLALL(user_id, "61TY3LELYT", "2") + SQLALL(user_id, "4ZXNVRV8B6", "2")
        '會客表單未完成表單筆數
        LabWait5.Text = SQLALL(user_id, "U28r13D6EA", "2")
        '門禁會議管制申請單未完成表單筆數
        LabWait6.Text = SQLALL(user_id, "S9QR2W8X6U", "2")
        '資訊設備維修申請單未完成表單筆數
        LabWait7.Text = SQLALL(user_id, "BL7U2QP3IG", "2")

    End Sub

    Public Function SQLALL(ByVal Uid As String, ByVal eformid As String, ByVal strFlag As String) As Integer

        Try

            connstr = conn.G_conn_string

            '新增連線
            Dim db As New SqlConnection(connstr)

            If strFlag = "1" Then

                '找出代理表單
                Dim AgentFlag As New C_Public

                If AgentFlag.AgentAll(user_id, Now.Date) = "" Then

                    '未批核
                    If eformid = "" Then
                        '全部統計
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT count(*) as FlowCount FROM flowctl WHERE empuid = '" & user_id & "' AND hddate IS NULL and (flowctl.gonogo='?' OR flowctl.gonogo='R')", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.read() Then
                            SQLALL = RdPer("FlowCount")
                        End If
                        db.Close()
                    Else
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT count(*) as FlowCount FROM flowctl WHERE empuid = '" & user_id & "' AND hddate IS NULL and (flowctl.gonogo='?' OR flowctl.gonogo='R') AND eformid='" & eformid & "'", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.read() Then
                            SQLALL = RdPer("FlowCount")
                        End If
                        db.Close()
                    End If

                Else

                    '未批核
                    If eformid = "" Then
                        '全部統計
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT count(*) as FlowCount FROM flowctl WHERE (empuid = '" & user_id & "' OR empuid IN (" & AgentFlag.AgentAll(user_id, Now.Date) & ")) AND hddate IS NULL and (flowctl.gonogo='?' OR flowctl.gonogo='R')", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.read() Then
                            SQLALL = RdPer("FlowCount")
                        End If
                        db.Close()
                    Else
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT count(*) as FlowCount FROM flowctl WHERE (empuid = '" & user_id & "' OR empuid IN (" & AgentFlag.AgentAll(user_id, Now.Date) & ")) AND hddate IS NULL and (flowctl.gonogo='?' OR flowctl.gonogo='R') AND eformid='" & eformid & "'", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.read() Then
                            SQLALL = RdPer("FlowCount")
                        End If
                        db.Close()
                    End If

                End If


            Else

                '申請未完成
                If eformid = "" Then
                    '全部統計
                    db.Open()
                    Dim strNoFlow As New SqlCommand("SELECT count(*) as FlowCountNO FROM flowctl WHERE hddate IS NULL AND eformsn IN (SELECT f.eformsn FROM flowctl f WHERE f.empuid = '" & user_id & "' AND gonogo='-')", db)
                    Dim RdNoFlow = strNoFlow.ExecuteReader()
                    If RdNoFlow.read() Then
                        SQLALL = RdNoFlow("FlowCountNO")
                    End If
                    db.Close()
                Else
                    db.Open()
                    Dim strNoFlow As New SqlCommand("SELECT count(*) as FlowCountNO FROM flowctl WHERE hddate IS NULL AND eformsn IN (SELECT f.eformsn FROM flowctl f WHERE f.empuid = '" & user_id & "' AND gonogo='-') AND eformid='" & eformid & "'", db)
                    Dim RdNoFlow = strNoFlow.ExecuteReader()
                    If RdNoFlow.read() Then
                        SQLALL = RdNoFlow("FlowCountNO")
                    End If
                    db.Close()
                End If

            End If

        Catch ex As Exception
            SQLALL = 0
        End Try


    End Function

End Class
