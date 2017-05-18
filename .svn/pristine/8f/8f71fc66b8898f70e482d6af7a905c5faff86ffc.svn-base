Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_00_MOA02002
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer
    Dim connstr As String
    Dim user_id, org_uid As String
    Dim MyConnection As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '日期
                Dim Sdate As TextBox = Me.FindControl("Sdate")
                Dim Edate As TextBox = Me.FindControl("Edate")

                Dim dt As Date = Now()
                If (Sdate.Text = "") Then
                    Sdate.Text = dt.Date
                End If

                If (Edate.Text = "") Then
                    Edate.Text = dt.AddDays(7).Date
                End If

                '找出登入者的一級單位
                Dim strParentOrg As String = ""
                Dim Org_UP As New C_Public
                strParentOrg = Org_UP.getUporg(org_uid, 1)

                '會議室
                SqlDataSource1.SelectCommand = "SELECT [MeetSn], MeetName+'('+(select emp_chinese_name from EMPLOYEE where Owner = employee_id)+')' as MeetName FROM [P_0201] WHERE (share = 1 OR Org_Uid IN (" & Org_UP.getchildorg(strParentOrg) & ")) AND Enabled=1 ORDER BY Share DESC,[MeetName]"

            End If

        Catch ex As Exception

        End Try


    End Sub

    Public Function ShowMeeting()

        Dim MeetSn As DropDownList = Me.FindControl("MeetSn")
        Dim Sdate As TextBox = Me.FindControl("Sdate")
        Dim Edate As TextBox = Me.FindControl("Edate")

        Dim sql As String
        Dim sdt As Date = Sdate.Text
        Dim edt As Date = Edate.Text
        Dim tmpStr As String = ""
        Dim day As Int32 = chk.BetweenDay(sdt, edt)
        Dim row(day, 4) As String
        Dim i As Int32 = 0
        Dim dd As Date
        If sdt.Date > edt.Date Then
            dd = sdt
            sdt = edt
            edt = dd
        End If

        For i = 0 To day - 1 Step 1
            tmpStr = sdt.AddDays(i)
            row(i, 0) = tmpStr
        Next

        sql = "SELECT convert(nvarchar,MeetTime,111) as MeetTime,MeetHour,P_0204.EFORMSN"
        sql += " FROM P_0204,P_02"
        sql += " WHERE MeetSn='" + MeetSn.Text + "'"
        sql += " and P_0204.EFORMSN=P_02.EFORMSN"
        sql += " and PENDFLAG<>'0'"
        If Sdate.Text <> "" Then
            sql += " AND MeetTime>=CONVERT(datetime,'" + sdt.Date + "')"
        End If
        If Edate.Text <> "" Then
            sql += " AND MeetTime<=CONVERT(datetime,'" + edt.Date + " 23:59:59')"
        End If
        sql += " ORDER BY MeetTime"

        MyConnection.Open()
        Dim dt1 As DataTable = New DataTable("P_0204")
        Dim da1 As SqlDataAdapter = New SqlDataAdapter(sql, MyConnection)
        da1.Fill(dt1)
        MyConnection.Close()

        For x As Integer = 0 To dt1.Rows.Count - 1

            For i = 0 To day - 1 Step 1

                dd = dt1.Rows(x).Item(0)

                tmpStr = dd.Date
                If row(i, 0) = tmpStr Then
                    If dt1.Rows(x).Item(1) = "全天" Then
                        row(i, 1) = dt1.Rows(x).Item(2)
                    ElseIf dt1.Rows(x).Item(1) = "上午" Then
                        row(i, 2) = dt1.Rows(x).Item(2)
                    ElseIf dt1.Rows(x).Item(1) = "下午" Then
                        row(i, 3) = dt1.Rows(x).Item(2)
                    End If
                    Exit For
                End If
            Next
        Next

        'Conn連線字串
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        Dim HyperLink1 As HyperLink
        Dim ShowMeet As Button

        '會議室申請基本資料
        Dim MeetShow1 As String = ""
        Dim MeetShow2 As String = ""
        Dim MeetShow3 As String = ""

        '漸層CSS
        Dim strclass As String = ""

        PlaceHolder1.Controls.Add(New LiteralControl("<table width='100%' border='3' bordercolor='#ccddee'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr><td align='center' class='RowClass' style='width:30%;'>日期</td><td align='center' class='RowClass' style='width:35%;'>上午</td><td align='center' class='RowClass' style='width:35%;'>下午</td></tr>"))
        For i = 0 To day - 1 Step 1
            tmpStr = sdt.Date

            If i Mod 2 = 0 Then
                strclass = "row_1"
            Else
                strclass = "row_2"
            End If

            PlaceHolder1.Controls.Add(New LiteralControl("<tr class='" & strclass & "'><td align='center' style='width:30%;'>" & row(i, 0) & DateShow(row(i, 0)) & "</td>"))

            '開啟連線
            Dim db As New SqlConnection(connstr)

            If row(i, 1) <> "" Then '全天

                '會議室申請基本資料
                MeetShow1 = ""

                '填表人資料
                db.Open()
                Dim strPer As New SqlCommand("SELECT PAUNIT,PANAME,nMEETNAME FROM P_02 WHERE EFORMSN = '" & row(i, 1) & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    MeetShow1 = RdPer("PAUNIT") & "-" & RdPer("PANAME") & "-" & RdPer("nMEETNAME")
                End If
                db.Close()

                HyperLink1 = New HyperLink
                With HyperLink1
                    .Text = MeetShow1
                    '.NavigateUrl = "../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N&MeetTime ='" & row(i, 0) & "'&MeetHour ='全天'"
                End With

                ShowMeet = New Button
                Dim strMeetList As String = ""
                With ShowMeet
                    strMeetList = "../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N&Read_Only=1&EFORMSN=" & row(i, 1)
                    .OnClientClick = "OpenMeetList('" & strMeetList & "')"
                    .Text = "詳細資料"
                End With

                PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' colspan=2 style='width:70%;'>"))
                PlaceHolder1.Controls.Add(HyperLink1)
                If MeetShow1 <> "" Then
                    PlaceHolder1.Controls.Add(New LiteralControl("<br>"))
                    PlaceHolder1.Controls.Add(ShowMeet)
                End If
                PlaceHolder1.Controls.Add(New LiteralControl("</td></tr>"))

            Else
                '會議室申請基本資料
                MeetShow2 = ""
                MeetShow3 = ""

                If row(i, 2) <> "" Then '上午

                    '填表人資料
                    db.Open()
                    Dim strPer As New SqlCommand("SELECT PAUNIT,PANAME,nMEETNAME FROM P_02 WHERE EFORMSN = '" & row(i, 2) & "'", db)
                    Dim RdPer = strPer.ExecuteReader()
                    If RdPer.read() Then
                        MeetShow2 = RdPer("PAUNIT") & "-" & RdPer("PANAME") & "-" & RdPer("nMEETNAME")
                    End If
                    db.Close()

                    tmpStr = MeetShow2
                Else
                    tmpStr = "　"
                End If

                HyperLink1 = New HyperLink
                With HyperLink1
                    .Text = tmpStr
                    '.NavigateUrl = "../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N&MeetTime ='" & row(i, 0) & "'&MeetHour ='上午'"
                End With

                ShowMeet = New Button
                Dim strMeetList1 As String = ""
                With ShowMeet
                    strMeetList1 = "../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N&Read_Only=1&EFORMSN=" & row(i, 2)
                    .OnClientClick = "OpenMeetList('" & strMeetList1 & "')"
                    .Text = "詳細資料"
                End With

                PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' style='width:35%;'>"))
                PlaceHolder1.Controls.Add(HyperLink1)
                If MeetShow2 <> "" Then
                    PlaceHolder1.Controls.Add(New LiteralControl("<br>"))
                    PlaceHolder1.Controls.Add(ShowMeet)
                End If
                PlaceHolder1.Controls.Add(New LiteralControl("</td>"))

                If row(i, 3) <> "" Then '下午

                    '填表人資料
                    db.Open()
                    Dim strPer As New SqlCommand("SELECT PAUNIT,PANAME,nMEETNAME FROM P_02 WHERE EFORMSN = '" & row(i, 3) & "'", db)
                    Dim RdPer = strPer.ExecuteReader()
                    If RdPer.read() Then
                        MeetShow3 = RdPer("PAUNIT") & "-" & RdPer("PANAME") & "-" & RdPer("nMEETNAME")
                    End If
                    db.Close()

                    tmpStr = MeetShow3
                Else
                    tmpStr = "　"
                End If

                HyperLink1 = New HyperLink
                With HyperLink1
                    .Text = tmpStr
                    '.NavigateUrl = "../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N&MeetTime ='" & row(i, 0) & "'&MeetHour ='下午'"
                End With

                ShowMeet = New Button
                Dim strMeetList2 As String = ""
                With ShowMeet
                    strMeetList2 = "../00/MOA00020.aspx?x=MOA02001&y=4rM2YFP73N&Read_Only=1&EFORMSN=" & row(i, 3)
                    .OnClientClick = "OpenMeetList('" & strMeetList2 & "')"
                    .Text = "詳細資料"
                End With

                PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' style='width:35%;'>"))
                PlaceHolder1.Controls.Add(HyperLink1)
                If MeetShow3 <> "" Then
                    PlaceHolder1.Controls.Add(New LiteralControl("<br>"))
                    PlaceHolder1.Controls.Add(ShowMeet)
                End If
                PlaceHolder1.Controls.Add(New LiteralControl("</td></tr>"))

            End If

        Next

        PlaceHolder1.Controls.Add(New LiteralControl("</table>"))

        '開啟會議室申請詳細資料
        PlaceHolder1.Controls.Add(New LiteralControl(" <script language='javascript'>"))
        PlaceHolder1.Controls.Add(New LiteralControl(" function OpenMeetList(x){"))
        PlaceHolder1.Controls.Add(New LiteralControl(" strFeatures = 'dialogWidth=900px;dialogHeight=700px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';"))
        PlaceHolder1.Controls.Add(New LiteralControl(" showModalDialog(x,self,strFeatures);"))
        PlaceHolder1.Controls.Add(New LiteralControl(" }"))
        PlaceHolder1.Controls.Add(New LiteralControl(" </script>"))

        ShowMeeting = ""

    End Function

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If MeetSn.Text <> "" Then
            '登入此頁面自動顯示出可點選的會議室日期
            ShowMeeting()
        End If

    End Sub

    Public Function DateShow(ByVal strdate As String) As String

        If DatePart(DateInterval.Weekday, CDate(strdate)) = "1" Then
            DateShow = "(日)"
        ElseIf DatePart(DateInterval.Weekday, CDate(strdate)) = "2" Then
            DateShow = "(一)"
        ElseIf DatePart(DateInterval.Weekday, CDate(strdate)) = "3" Then
            DateShow = "(二)"
        ElseIf DatePart(DateInterval.Weekday, CDate(strdate)) = "4" Then
            DateShow = "(三)"
        ElseIf DatePart(DateInterval.Weekday, CDate(strdate)) = "5" Then
            DateShow = "(四)"
        ElseIf DatePart(DateInterval.Weekday, CDate(strdate)) = "6" Then
            DateShow = "(五)"
        Else
            DateShow = "(六)"
        End If

    End Function

    'Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

    '    Div_grid.Visible = True
    '    Div_grid.Style("Top") = "70px"
    '    Div_grid.Style("left") = "400px"

    '    Calendar1.SelectedDate = Sdate.Text

    'End Sub

    'Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

    '    Div_grid2.Visible = True
    '    Div_grid2.Style("Top") = "70px"
    '    Div_grid2.Style("left") = "550px"

    '    Calendar2.SelectedDate = Edate.Text

    'End Sub

    'Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

    '    Sdate.Text = Calendar1.SelectedDate.Date
    '    Div_grid.Visible = False

    'End Sub

    'Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

    '    Edate.Text = Calendar2.SelectedDate.Date
    '    Div_grid2.Visible = False

    'End Sub

    'Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

    '    Div_grid.Visible = False

    'End Sub

    'Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

    '    Div_grid2.Visible = False

    'End Sub

End Class
