Imports System.Data.sqlclient
Partial Class Source_05_MOA05005
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Dim connstr As String = ""
    Dim user_id, org_uid As String

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

            If IsPostBack = False Then

                '找出上一級單位以下全部單位
                Dim Org_Down As New C_Public

                ''判斷登入者權限
                'If Session("Role") = "1" Then
                '    SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                'ElseIf Session("Role") = "2" Then
                '    SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(Session("PARENT_ORG_UID")) & ") ORDER BY ORG_NAME"
                'Else
                '    SqlDataSource1.SelectCommand = ""
                'End If

                Dim nAPPTIME1 As TextBox = Me.FindControl("nAPPTIME1")
                Dim nAPPTIME2 As TextBox = Me.FindControl("nAPPTIME2")

                Dim dt As Date = Now()

                If (nAPPTIME1.Text = "") Then
                    nAPPTIME1.Text = dt.AddDays(-14).Date
                End If
                If (nAPPTIME2.Text = "") Then
                    nAPPTIME2.Text = dt.Date
                End If

                '登入馬上查詢
                ImgSearch_Click(Nothing, Nothing)

            End If

        End If

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)
        Dim tool As New C_Public

        If (sort) Then
            GridView1.Sort("PAUNIT", SortDirection.Ascending)
            Return
        End If

        '新增連線
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        'Dim VistorFlag As String = ""
        'Dim sqlVistor As String = ""

        ''判斷登入者是否為會客相關群組人員
        'db.Open()
        'Dim carCom As New SqlCommand("select object_uid from SYSTEMOBJUSE where employee_id = '" & user_id & "' AND ((object_uid = 'E539') OR (object_uid = 'Y965'))", db)
        'Dim RdvCar = carCom.ExecuteReader()
        'If RdvCar.read() Then
        '    VistorFlag = RdvCar("object_uid")
        'End If
        'db.Close()

        'If VistorFlag = "E539" Then
        '    sqlVistor = " and nRECROOM='第一會客室' "
        'ElseIf VistorFlag = "Y965" Then
        '    sqlVistor = " and nRECROOM='第二會客室' "
        'End If

        Dim arrOpenGateName() As String = {}
        Dim sqlVistor As String = ""
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String
        Dim i As Integer = 0
        ''會客管制群組內的人員可以看到會客統計資料
        If tool.CheckRoleGroupEMPByName("會客管制群組", user_id) Then
            ''取出現行開放之營門統計
            DC = New SQLDBControl
            strSql = "SELECT * FROM SYSKIND WHERE KIND_NUM='3' AND STATE_ENABLED='1'"
            DR = DC.CreateReader(strSql)
            If DR.HasRows Then
                While DR.Read
                    ReDim Preserve arrOpenGateName(i)
                    arrOpenGateName(i) = DR("State_Name")
                    i = i + 1
                End While
            End If
            DC.Dispose()
        End If

        If arrOpenGateName.Length > 0 Then
            For Each s As String In arrOpenGateName
                If sqlVistor.Length > 0 Then sqlVistor = sqlVistor + ","
                sqlVistor = sqlVistor + "'" + s + "'"
            Next
            sqlVistor = " AND nRECROOM IN (" + sqlVistor + ")"
        End If
        
        Dim sql As String

        sql = "SELECT PAUNIT,count(*) count, P_05.nRECROOM"
        sql += "  FROM P_05 CROSS JOIN P_0501"
        sql += " where P_05.EFORMSN=P_0501.EFORMSN AND PENDFLAG='E'"

        If (PAUNIT.Text <> "") Then sql += " and P_05.PAUNIT like '%" + PAUNIT.Text + "%'" '申請人單位
        If (nAPPTIME1.Text <> "") Then sql += " AND (nRECDATE between '" & nAPPTIME1.Text & "' AND '" & nAPPTIME2.Text & "')" '會客時間

        'If (nAPPTIME1.Text <> "") Then sql += " and P_05.nAPPTIME>=CONVERT(datetime,'" + nAPPTIME1.Text + " 00:00:00')" '申請時間
        'If (nAPPTIME2.Text <> "") Then sql += " and P_05.nAPPTIME<=CONVERT(datetime,'" + nAPPTIME2.Text + " 23:59:59')" '申請時間

        If (sqlVistor <> "") Then sql += sqlVistor '會客室管理者

        sql += " group by PAUNIT,nRECROOM"
        sql += " order by PAUNIT"

        SqlDataSource2.SelectCommand = sql
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        PAUNIT.Text = ""
        nAPPTIME1.Text = ""
        nAPPTIME2.Text = ""
    End Sub

End Class
