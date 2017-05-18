Imports System.Data.SqlClient
Imports System.Data

Partial Class Source_00_MOA00011
    Inherits Page
    Dim conn As New C_SQLFUN
    Dim connstr As String = conn.G_conn_string
    Dim db As New SqlConnection(connstr)
    Public OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        user_id = Session("user_id").ToString()
        org_uid = Session("ORG_UID").ToString()

        Try

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00011") <> "" Then

                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00011.aspx")
                    Response.End()

                End If

                If IsPostBack = False Then

                    '先設定起始日期
                    Dim dt As Date = Now()
                    If (Sdate.Text = "") Then
                        Sdate.Text = dt.AddDays(-14).Date.ToShortDateString()
                    End If

                    If (Edate.Text = "") Then
                        Edate.Text = dt.Date.ToShortDateString()
                    End If

                    '找出上一級單位以下全部單位
                    'Dim Org_Down As New C_Public

                    '判斷登入者權限1:系統管理員2:單位管理員3:其他
                    If Session("Role") = "1" Then
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = ""
                        'ElseIf Session("Role") = "2" Then
                        '    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(Session("PARENT_ORG_UID")) & ") ORDER BY ORG_NAME"
                    Else
                        SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID ='" & org_uid & "' ORDER BY ORG_NAME"
                        SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE employee_id ='" & user_id & "'"
                    End If

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As EventArgs) Handles Me.PreRenderComplete

        Try
            If Not IsPostBack Then
                Dim tLItm As New ListItem("請選擇", "")

                '表單增加請選擇
                FormSel.Items.Insert(0, tLItm)
                'FormSel.Items(1).Selected = False

                '系統管理員組織加請選擇
                If Session("Role") = "1" Then

                    OrgSel.Items.Insert(0, tLItm)
                    If OrgSel.Items.Count > 1 Then
                        OrgSel.Items(1).Selected = False
                    End If

                End If

                '人員加請選擇
                UserSel.Items.Insert(0, tLItm)
                If UserSel.Items.Count > 1 Then
                    UserSel.Items(1).Selected = False
                End If

                '重新整理GridView
                Dim strOrd As String

                strOrd = " ORDER BY appdate DESC"

                '是否由未批核表單頁面傳送過來
                Dim strHyper As String = Request.QueryString("strHyper")

                '是否由入口網站頁面傳送過來
                Dim strPortal As String = Request.QueryString("strPortal")

                If strHyper = "" And strPortal = "" Then
                    SqlDataSource2.SelectCommand = SQLALL(FormSel.SelectedValue, OrgSel.SelectedItem.Text, UserSel.SelectedValue, Sdate.Text, Edate.Text, StateSel.SelectedValue, "") & strOrd
                ElseIf strHyper <> "" And strPortal = "" Then

                    If strHyper = "YAqBTxRP8P" Or strHyper = "4rM2YFP73N" Or strHyper = "j2mvKYe3l9" _
                        Or strHyper = "61TY3LELYT" Or strHyper = "U28r13D6EA" Or strHyper = "5D82872F5L" _
                        Or strHyper = "4ZXNVRV8B6" Or strHyper = "S9QR2W8X6U" Or strHyper = "BL7U2QP3IG" Then '加新表單BL7U2QP3IG

                        SqlDataSource2.SelectCommand = SQLALL(strHyper, OrgSel.SelectedItem.Text, UserSel.SelectedValue, "", "", StateSel.SelectedValue, "1") & strOrd

                    Else

                        '判斷登入者權限
                        Dim LoginCheck As New C_Public

                        '禁止無帳號者竄入
                        LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00011.aspx")
                        Response.End()

                    End If
                ElseIf strHyper = "" And strPortal <> "" Then
                    If strPortal = "portal" Then

                        SqlDataSource2.SelectCommand = SQLALL(FormSel.SelectedValue, OrgSel.SelectedItem.Text, user_id, "", "", StateSel.SelectedValue, "1") & strOrd

                    Else

                        '判斷登入者權限
                        Dim LoginCheck As New C_Public

                        '禁止無帳號者竄入
                        LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00011.aspx")
                        Response.End()

                    End If
                End If

                OrgChange = ""

            End If

        Catch ex As Exception

        End Try



    End Sub

    Protected Sub UserSel_PreRender(ByVal sender As Object, ByVal e As EventArgs) Handles UserSel.PreRender

        Try

            If OrgChange = "1" Then
                Dim tLItm As New ListItem("請選擇", "")

                '人員加請選擇
                UserSel.Items.Insert(0, tLItm)
                If UserSel.Items.Count > 1 Then
                    UserSel.Items(1).Selected = False
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub OrgSel_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles OrgSel.SelectedIndexChanged
        Try

            '判斷組織是否變更
            OrgChange = "1"

            '清空User重新讀取
            UserSel.Items.Clear()

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As ImageClickEventArgs) Handles ImgSearch.Click

        Try

            Const strOrd As String = " ORDER BY appdate DESC"
            Sdate.Text = Page.Request.Form("Sdate")
            Edate.Text = Page.Request.Form("Edate")
            SqlDataSource2.SelectCommand = SQLALL(FormSel.SelectedValue.ToString(), OrgSel.SelectedItem.Text, UserSel.SelectedValue.ToString(), Sdate.Text, Edate.Text, StateSel.SelectedValue.ToString(), "") & strOrd
            ''SqlDataSource2.SelectCommand = SQLALL(FormSel.SelectedValue, OrgSel.SelectedItem.Text, UserSel.SelectedValue, CType(Page.Request.Form("Sdate"), String), CType(Page.Request.Form("Edate"), String), StateSel.SelectedValue, "") & strOrd

            OrgChange = ""
            GridView1.DataSourceID = SqlDataSource2.ID

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles GridView1.PageIndexChanged

        Try

            '分頁
            Dim strOrd As String

            strOrd = " ORDER BY appdate DESC"

            SqlDataSource2.SelectCommand = SQLALL(FormSel.SelectedValue, OrgSel.SelectedItem.Text, UserSel.SelectedValue, Sdate.Text, Edate.Text, StateSel.SelectedValue, "") & strOrd

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles GridView1.RowCreated

        Try

            If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

                '隱藏eformsn,eformid
                e.Row.Cells(7).Visible = False
                e.Row.Cells(8).Visible = False

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As EventArgs) Handles GridView1.Sorted

        Try

            Dim strOrd As String

            '排序條件
            strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

            SqlDataSource2.SelectCommand = SQLALL(FormSel.SelectedValue, OrgSel.SelectedItem.Text, UserSel.SelectedValue, Sdate.Text, Edate.Text, StateSel.SelectedValue, "") & strOrd

        Catch ex As Exception

        End Try

    End Sub

    Public Function SQLALL(ByVal sFormSel As String, ByVal sOrgSel As String, ByVal sUserSel As String, ByVal sSDate As String, ByVal sEDate As String, ByVal sStateSel As String, ByVal sQueryStr As String) As String

        '整合SQL搜尋字串
        Dim strsel As String
        Dim strDate As String = ""
        Const strsql As String = "Select flowsn,eformsn,eformid,empuid,emp_chinese_name,appdate,deptcode,frm_chinese_name,ORG_NAME,status From V_EformFlow WHERE 1=1 "

        '系統管理員
        If Session("Role") = "1" Then

            '組織搜尋
            If sOrgSel = "請選擇" Then
                strsel = " AND 1=1"
            Else
                strsel = " AND ORG_NAME = '" & sOrgSel & "'"
            End If

            '人員
            If sUserSel <> "" Then
                strsel += " AND empuid = '" & sUserSel & "'"
            End If

        Else

            '人員
            strsel = " AND empuid = '" & user_id & "'"

        End If

        '表單搜尋
        If sFormSel <> "" Then

            strsel += " AND eformid = '" & sFormSel & "'"

        End If

        '是否由上一頁傳送過來查詢
        If sQueryStr = "" Then

            '狀態搜尋
            If sStateSel <> "請選擇" Then
                If sStateSel = "未批核" Then
                    strsel += " AND status = '?'"
                ElseIf sStateSel = "駁回" Then
                    strsel += " AND status = '0'"
                Else
                    strsel += " AND status <> '?'"
                End If
            End If

        Else
            strsel += " AND status = '?'"
        End If

        '申請日期搜尋
        If sSDate <> "" And sEDate <> "" Then
            strDate = " AND (appdate between '" & sSDate & " 00:00:00 ' AND '" & sEDate & " 23:59:59')"
        End If

        SQLALL = strsql & strsel & strDate

    End Function

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles GridView1.SelectedIndexChanged

        Try

            Dim streformid, streformsn As String
            Dim strPath As String = ""

            '顯示選取的表單資料
            streformsn = GridView1.Rows(GridView1.SelectedIndex).Cells(7).Text
            streformid = GridView1.Rows(GridView1.SelectedIndex).Cells(8).Text

            '表單資料夾
            Dim printRecordsReportID As String = GetEFormId("影印紀錄呈核單") '取得影印紀錄呈核單ID
            Dim doorAndMeetingControlID As String = GetEFormId("門禁會議管制申請單") '取得門禁會議管制申請單ID
            If streformid = "YAqBTxRP8P" Then       '請假申請單
                strPath = "MOA00020.aspx?x=MOA01001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "4rM2YFP73N" Then   '會議室申請單
                strPath = "MOA00020.aspx?x=MOA02001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "j2mvKYe3l9" Then   '派車申請單
                strPath = "MOA00020.aspx?x=MOA03001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "61TY3LELYT" Then   '房舍水電申請單
                strPath = "MOA00020.aspx?x=MOA04001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "F9MBD7O97G" Then   '房舍水電申請單(新) peter
                strPath = "MOA00020.aspx?x=MOA04100&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "4ZXNVRV8B6" Then   '完工報告單
                strPath = "MOA00020.aspx?x=MOA04003&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "U28r13D6EA" Then   '會客洽公申請單
                strPath = "MOA00020.aspx?x=MOA05001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
                strPath = "MOA00020.aspx?x=MOA06001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "9JKSDRR5V3" Then   '報修申請單
                strPath = "MOA00020.aspx?x=MOA07001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "5D82872F5L" Then   '銷假申請單
                strPath = "MOA00020.aspx?x=MOA01003&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "74BN58683M" Then   '影印使用申請單
                strPath = "MOA00020.aspx?x=MOA08001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = printRecordsReportID Then   '影印紀錄呈核單
                strPath = "MOA00020.aspx?x=MOA08014&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = doorAndMeetingControlID Then   '門禁會議管制申請單
                strPath = "MOA00020.aspx?x=MOA09001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            ElseIf streformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
                strPath = "MOA00020.aspx?x=MOA11001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
            End If

            Response.Write(" <script language='javascript'>")
            Response.Write(" sPath = '" & strPath & "';")
            Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=700px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
            Response.Write(" showModalDialog(sPath,self,strFeatures);")
            Response.Write(" </script>")

        Catch ex As Exception

        End Try

    End Sub

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        'db.Open()
        'Dim sqlcomm As New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db)
        'GetEFormId = sqlcomm.ExecuteScalar().ToString()
        'db.Close()
        Dim sReturn As String = ""
        Dim strSql As String = "SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'"
        Dim DC = New SQLDBControl
        Dim DR As SqlDataReader
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("eformid")
            End If
        End If
        DC.Dispose()
        GetEFormId = sReturn
    End Function

    Protected Function FunStatus(ByVal str As String) As String

        Try
            '轉換表單狀態代號
            Dim tmpStr As String = Eval(str).ToString()

            If tmpStr = "-" Then
                tmpStr = "申請"
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

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles GridView1.RowDataBound

        Select Case (e.Row.RowType)
            Case DataControlRowType.DataRow
                ''匯出外部資料庫時預設全選
                If CType(ViewState("HasOutputData"), Boolean) = True Then
                    ''Dim chk As CheckBox = CType(e.Row.Cells(0).Controls(1), CheckBox)
                    Dim chk As CheckBox = CType(e.Row.Cells(0).FindControl("selchk"), CheckBox)
                    If chk IsNot Nothing Then
                        chk.Checked = True
                    End If
                End If
        End Select

        '等中科院驗測再換版
        'If e.Row.RowType = DataControlRowType.DataRow Then

        '    If Not e.Row.RowIndex = -1 Then

        '        Dim streformid, streformsn As String
        '        Dim strPath As String = ""

        '        '顯示選取的表單資料
        '        streformsn = e.Row.Cells(7).Text
        '        streformid = e.Row.Cells(8).Text

        '        '表單資料夾

        '        If streformid = "YAqBTxRP8P" Then       '請假申請單
        '            strPath = "MOA00020.aspx?x=MOA01001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "4rM2YFP73N" Then   '會議室申請單
        '            strPath = "MOA00020.aspx?x=MOA02001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "j2mvKYe3l9" Then   '派車申請單
        '            strPath = "MOA00020.aspx?x=MOA03001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "61TY3LELYT" Then   '房舍水電申請單
        '            strPath = "MOA00020.aspx?x=MOA04001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "4ZXNVRV8B6" Then   '完工報告單
        '            strPath = "MOA00020.aspx?x=MOA04002&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "U28r13D6EA" Then   '會客洽公申請單
        '            strPath = "MOA00020.aspx?x=MOA05001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
        '            strPath = "MOA00020.aspx?x=MOA06001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "9JKSDRR5V3" Then   '報修申請單
        '            strPath = "MOA00020.aspx?x=MOA07001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        ElseIf streformid = "5D82872F5L" Then   '銷假申請單
        '            strPath = "MOA00020.aspx?x=MOA01003&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        '        End If


        '        '移入顏色 
        '        e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='#E6F5FA'")
        '        '移出顏色 
        '        e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='#FFFFFF'")

        '        e.Row.Attributes.Add("onclick", "location.href='" & strPath & "'")

        '    End If

        'End If

    End Sub

    Protected Sub ImgRevoke_Click(ByVal sender As Object, ByVal e As ImageClickEventArgs) Handles ImgRevoke.Click
        Dim doorAndMeetingControlID As String = GetEFormId("門禁會議管制申請單") '取得門禁會議管制申請單ID
        Try
            Dim ErrMsg As String = ""
            '表單撤銷
            Dim FC As New C_FlowSend.C_FlowSend
            Dim i As Integer
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then
                    Dim eformsn, eformid As String
                    Dim strPENDFLAG As String = ""
                    eformsn = GridView1.Rows(i).Cells(7).Text
                    eformid = GridView1.Rows(i).Cells(8).Text
                    '取得連線
                    Dim conn As New C_SQLFUN
                    Dim connstr As String = conn.G_conn_string.ToString()
                    '開啟連線
                    Dim db As New SqlConnection(connstr)
                    If eformid = doorAndMeetingControlID Then '撤銷門禁會議管制申請單
                        '判斷表單是否已批核完成
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT EFORMSN,PENDFLAG FROM P_09 WHERE EFORMSN = '" & eformsn & "'", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.Read() Then
                            'PENDFLAG: "-":申請, "0":駁回, "1":送件, "?":審核中, "E":完成, "G":補登, "B" or "X":撤銷, "R":重新分派, "T":呈轉
                            strPENDFLAG = RdPer("PENDFLAG").ToString()
                        End If
                        db.Close()
                        '表單已批核不可撤銷
                        If strPENDFLAG <> "E" Then                            
                            Dim SendVal As String = eformid & "," & eformsn
                            Dim Val_P As String = FC.F_DrawBack(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString).ToString()
                            If Val_P = "1" Then
                                ErrMsg = "表單已批核不可撤銷"
                            Else
                                db.Open()
                                '將門禁會議管制申請單(P_09)狀態變更為3, 表單狀態 [ 0：未登管 1：已登管 2：退件 3：撤銷 ]
                                Dim insCom As New SqlCommand("UPDATE P_09 SET Status = 3,ModifyBy = @ModifyBy,ModifyDate = GETDATE() WHERE EFORMSN = @EFORMSN", db)
                                insCom.Parameters.Add("@ModifyBy", SqlDbType.VarChar, 10).Value = user_id
                                insCom.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = user_id
                                insCom.ExecuteNonQuery()
                                db.Close()
                                ErrMsg = "表單撤銷完成"
                            End If
                        End If
                    Else '撤銷非門禁會議管制申請單
                        '判斷表單是否已批核完成
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT EFORMSN,PENDFLAG FROM P_01 WHERE EFORMSN = '" & eformsn & "'", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.Read() Then
                            'PENDFLAG: "-":申請, "0":駁回, "1":送件, "?":審核中, "E":完成, "G":補登, "B" or "X":撤銷, "R":重新分派, "T":呈轉
                            strPENDFLAG = RdPer("PENDFLAG").ToString()
                        End If
                        db.Close()

                        '表單已批核不可撤銷
                        If strPENDFLAG <> "E" Then                            
                            Dim SendVal As String = eformid & "," & eformsn
                            Dim Val_P As String = FC.F_DrawBack(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString).ToString()
                            If Val_P = "1" Then
                                ErrMsg = "表單已批核不可撤銷"
                            Else
                                ErrMsg = "表單撤銷完成"
                            End If
                        End If

                        '撤銷所申請的會議室
                        If eformid = "4rM2YFP73N" Then
                            '將會議室申請時段放到移除資料表
                            db.Open()
                            Dim insCom As New SqlCommand("INSERT INTO P_0205 (MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, DelUser) SELECT MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, '" & user_id & "' FROM P_0204 WHERE EFORMSN='" & eformsn & "'", db)
                            insCom.ExecuteNonQuery()
                            db.Close()
                            db.Open()
                            Dim delCom As New SqlCommand("DELETE FROM P_0204 WHERE (EFORMSN = '" & eformsn & "')", db)
                            delCom.ExecuteNonQuery()
                            db.Close()
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

            strOrd = " ORDER BY appdate DESC"

            SqlDataSource2.SelectCommand = SQLALL(FormSel.SelectedValue, OrgSel.SelectedItem.Text, UserSel.SelectedValue, Sdate.Text, Edate.Text, StateSel.SelectedValue, "") & strOrd

            OrgChange = ""

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Calendar1.SelectionChanged

        Try

            Sdate.Text = Calendar1.SelectedDate.Date.ToString()
            Div_grid.Visible = False

        Catch ex As Exception

        End Try
    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Calendar2.SelectionChanged

        Try

            Edate.Text = Calendar2.SelectedDate.Date.ToString()
            Div_grid2.Visible = False

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

    Protected Sub FormSel_DataBound(sender As Object, e As System.EventArgs) Handles FormSel.DataBound
        'If Not IsNothing(FormSel.Items.FindByValue(Request("strHyper"))) Then
        '    FormSel.SelectedValue = Request("strHyper")
        'End If
    End Sub
End Class
