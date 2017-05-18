Imports System.Data.sqlclient
Partial Class Source_01_MOA01002
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Public OrgChange As String              '判斷組織是否變更
    Dim user_id, org_uid, org_name As String
    Dim ParentFlag As String = ""
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        org_name = Session("ORG_NAME")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            Dim connstr As String = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷是否有下一級單位
            db.Open()
            Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                ParentFlag = "Y"
            End If
            db.Close()

            If IsPostBack = False Then
				
				Dim UnitFlag As String = ""

				'判斷是否有處單位人事員
				db.Open()
				Dim strUnit As New SqlCommand("SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = 'JKGJZ4439V' AND employee_id = '" & user_id & "'", db)
				Dim RdUnit = strUnit.ExecuteReader()
				If RdUnit.read() Then
					UnitFlag = "Y"
				End If
				db.Close()
				
                '找出上一級單位以下全部單位
                Dim Org_Down As New C_Public

				'找出登入者的二級單位
				Dim strParentOrg As String = ""
				strParentOrg = Org_Down.getUporg(org_uid, 2)

			
                '判斷登入者權限
				If Session("Role") = "1" Then
					SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
				Else

                If ParentFlag = "Y" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") ORDER BY ORG_NAME"
                ElseIf UnitFlag = "Y" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") ORDER BY ORG_NAME"
                Else
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "' ORDER BY ORG_NAME"
                End If

            End If

                '設定Default日期
                Dim Sdate As TextBox = Me.FindControl("Sdate")
                Dim Edate As TextBox = Me.FindControl("Edate")

                Dim dt As Date = Now()
                If (Sdate.Text = "") Then
                    Sdate.Text = dt.AddDays(-14).Date
                End If

                If (Edate.Text = "") Then
                    Edate.Text = dt.Date
                End If

            End If

        End If


    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Dim strOrd As String

        strOrd = " ORDER BY nAPPTIME DESC"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

        OrgChange = ""

    End Sub

    Protected Sub UserSel_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles UserSel.PreRender

        If OrgChange = "1" Then
            Dim tLItm As New ListItem("請選擇", "")

            '人員加請選擇
            UserSel.Items.Insert(0, tLItm)
            If UserSel.Items.Count > 1 Then
                UserSel.Items(1).Selected = False
            End If

        End If

    End Sub

    Protected Sub OrgSel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OrgSel.SelectedIndexChanged

        '判斷組織是否變更
        OrgChange = "1"

        '清空User重新讀取
        UserSel.Items.Clear()

        If OrgSel.SelectedValue = "" Then
            SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE 1=2"
        Else
            'If Session("Role") = "1" Or ParentFlag = "Y" Then
                SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & OrgSel.SelectedValue & "' ORDER BY emp_chinese_name"
            'Else
            '    SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & OrgSel.SelectedValue & "' AND employee_id='" & user_id & "' ORDER BY emp_chinese_name"
            'End If
        End If

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete
        Try

            If Not IsPostBack Then

                Dim tLItm As New ListItem("請選擇", "")

                '系統管理員組織加請選擇

                OrgSel.Items.Insert(0, tLItm)
                If OrgSel.Items.Count > 1 Then
                    OrgSel.Items(1).Selected = False
                End If

                '人員加請選擇
                UserSel.Items.Insert(0, tLItm)
                If UserSel.Items.Count > 1 Then
                    UserSel.Items(1).Selected = False
                End If

                '登入馬上查詢
                ImgSearch_Click(Nothing, Nothing)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY nAPPTIME DESC"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Public Function SQLALL(ByVal OrgSel, ByVal UserSel, ByVal SDate, ByVal EDate)

        Try

            Dim strsel As String = ""

			Dim conn As New C_SQLFUN
            Dim connstr As String = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)
			
            '整合SQL搜尋字串
            Dim strsql, strDate As String
            strsql = "SELECT EFORMSN, PWUNIT, PWTITLE, PWNAME, PWIDNO, PAUNIT, PANAME, PATITLE, PAIDNO, nSTATUS, nAPPTIME, nTYPE, nPROVEMENT, CONVERT(char(12), nSTARTTIME, 111) AS nSTARTTIME, nSTHOUR, CONVERT(char(12), nENDTIME, 111) AS nENDTIME, nETHOUR, nDAY, nHOUR, nAGENT1, nAGENT2, nAGENT3, nPLACE, nPHONE, nREASON,P_NUM,PENDFLAG,ORG_UID FROM P_01,EMPLOYEE WHERE PAIDNO=employee_id and eformsn in (select eformsn from flowctl) "

			Dim UnitFlag As String = ""

			'判斷是否有處單位人事員
			db.Open()
			Dim strUnit As New SqlCommand("SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = 'JKGJZ4439V' AND employee_id = '" & user_id & "'", db)
			Dim RdUnit = strUnit.ExecuteReader()
			If RdUnit.read() Then
				UnitFlag = "Y"
			End If
			db.Close()
		
            '系統管理員
            If Session("Role") = "1" Then

                '組織搜尋
                If OrgSel = "" Then
                    strsel = " AND 1=1"
                Else
                    strsel = " AND ORG_UID='" & OrgSel & "'"
                End If

                '人員
                If UserSel <> "" Then
                    strsel += " AND PAIDNO = '" & UserSel & "'"
                End If

                '申請日期搜尋
                strDate = " AND (nAPPTIME between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"

            Else

                If ParentFlag = "Y" Or UnitFlag = "Y" Then

                    '主官管
                    '組織搜尋
                    If OrgSel = "" Then
                        strsel = " AND ORG_UID='" & org_uid & "'"
                    Else
                        strsel = " AND ORG_UID='" & OrgSel & "'"
                    End If

                    '人員
                    If UserSel <> "" Then
                        strsel += " AND PAIDNO = '" & UserSel & "'"
                    End If


                Else

                    '人員
                    strsel = " AND PAIDNO = '" & Session("user_id") & "'"

                End If

                '申請日期搜尋
                strDate = " AND (nAPPTIME between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"

            End If

            SQLALL = strsql & strsel & strDate

        Catch ex As Exception

            SQLALL = ""

        End Try

    End Function

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim streformsn As String = ""
        Dim strPath As String = ""

        '顯示選取的表單資料
        Dim strP_NUM As String = GridView1.Rows(GridView1.SelectedIndex.ToString()).Cells(1).Text

        '取得連線
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '取得表單EFORMSN
        db.Open()
        Dim strPer As New SqlCommand("SELECT EFORMSN FROM P_01 WHERE P_NUM = '" & strP_NUM & "'", db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.read() Then
            streformsn = RdPer("EFORMSN")
        End If
        db.Close()

        '表單資料夾
        strPath = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & streformsn

        Response.Write(" <script language='javascript'>")
        Response.Write(" sPath = '" & strPath & "';")
        Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=700px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
        Response.Write(" showModalDialog(sPath,self,strFeatures);")
        Response.Write(" </script>")


    End Sub

    Protected Function FunStatus(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr = Eval(str)

            If tmpStr = "-" Then
                tmpStr = "申請"
            ElseIf tmpStr = "0" Then
                tmpStr = "駁回"
            ElseIf tmpStr = "1" Then
                tmpStr = "送件"
            ElseIf tmpStr = "?" Or tmpStr = "" Then
                tmpStr = "審核中"
            ElseIf tmpStr = "E" Then
                tmpStr = "完成"
            ElseIf tmpStr = "G" Then
                tmpStr = "補登"
            ElseIf tmpStr = "B" Then
                tmpStr = "撤銷"
            ElseIf tmpStr = "R" Then
                tmpStr = "重新分派"
            ElseIf tmpStr = "T" Then
                tmpStr = "呈轉"
            Else
                tmpStr = "未知"
            End If

            FunStatus = tmpStr

        Catch ex As Exception
            FunStatus = ""
        End Try
    End Function

    Protected Sub ImgRevoke_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgRevoke.Click

        Dim ErrMsg As String = ""

        '表單撤銷
        Dim FC As New C_FlowSend.C_FlowSend

        Dim i As Integer = 0
        For i = 0 To GridView1.Rows.Count - 1
            If CType(GridView1.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then

                Dim streformsn As String = ""
                Dim strPENDFLAG As String = ""
                Dim streformid As String = "YAqBTxRP8P"

                '顯示選取的表單資料
                Dim strP_NUM As String = GridView1.Rows(i).Cells(1).Text

                '取得連線
                Dim connstr As String
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '取得表單EFORMSN
                db.Open()
                Dim strPer As New SqlCommand("SELECT EFORMSN,PENDFLAG FROM P_01 WHERE P_NUM = '" & strP_NUM & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    streformsn = RdPer("EFORMSN")
                    strPENDFLAG = RdPer("PENDFLAG")
                End If
                db.Close()

                '表單已批核不可撤銷
                If strPENDFLAG <> "E" Then

                    Dim Val_P As String = ""

                    Dim SendVal As String = streformid & "," & streformsn

                    Val_P = FC.F_DrawBack(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

                    If Val_P = "1" Then
                        ErrMsg = "表單已批核不可撤銷"
                    Else
                        ErrMsg = "表單撤銷完成"
                    End If

                End If

            End If
        Next

        If ErrMsg <> "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('" & ErrMsg & "!!!');")
            Response.Write(" </script>")

        End If

        '重新整理GridView
        Dim strOrd As String

        strOrd = " ORDER BY nAPPTIME DESC"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

        OrgChange = ""

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '隱藏流水號
            e.Row.Cells(1).Visible = False
            e.Row.Cells(12).Visible = False
        End If

    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                'e.Row.Cells(10).Text = FunStatus(e.Row.Cells(10).Text)
                Dim P_0101Status As String = ""
                Dim DC As New SQLDBControl
                Dim DR As SqlDataReader
                Dim strSql As String = "SELECT * FROM P_0101 WHERE NEFORMSN='" + e.Row.Cells(12).Text + "'"
                DR = DC.CreateReader(strSql)
                If DR.HasRows Then
                    If DR.Read Then
                        P_0101Status = DR("PENDFLAG")
                    End If
                End If
                DC.Dispose()
                If P_0101Status = "E" Then
                    e.Row.Cells(10).Text = "銷假"
                End If
        End Select
    End Sub

    Protected Sub UserSel_DataBound(sender As Object, e As System.EventArgs) Handles UserSel.DataBound
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        Dim UnitFlag As String = ""
        Dim ParentFlag As String = ""

        '判斷是否有處單位人事員
        DC = New SQLDBControl
        strSql = "SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = 'JKGJZ4439V' AND employee_id = '" & user_id & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                UnitFlag = "Y"
            End If
        End If
        DC.Dispose()

        '判斷是否有下一級單位
        DC = New SQLDBControl
        strSql = "SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                ParentFlag = "Y"
            End If
        End If
        DC.Dispose()

        Dim ddl As DropDownList = CType(sender, DropDownList)
        ''設定預設下拉資料                
        If ddl.Items.Contains(ddl.Items.FindByValue(LCase(Session("user_id")))) Then
            ddl.SelectedValue = LCase(Session("user_id"))
        ElseIf ddl.Items.Contains(ddl.Items.FindByValue(UCase(Session("user_id")))) Then
            ddl.SelectedValue = UCase(Session("user_id"))
        End If

        '系統管理員
        If (Not Session("Role") = "1") And Not (ParentFlag = "Y") And (Not UnitFlag = "Y") Then
            UserSel.Enabled = False
        End If
    End Sub
End Class
