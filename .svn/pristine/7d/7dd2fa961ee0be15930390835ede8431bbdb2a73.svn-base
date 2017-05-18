Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_01_MOA01006
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer
    Dim connstr As String
    Dim user_id, org_uid As String

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

            End If

        Catch ex As Exception

        End Try


    End Sub

    Public Function ShowMeeting()

        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        Dim Sdate As TextBox = Me.FindControl("Sdate")
        Dim Edate As TextBox = Me.FindControl("Edate")

        Dim sql As String
        Dim sdt As Date = Sdate.Text
        Dim edt As Date = Edate.Text
        Dim tmpStr As String = ""
        Dim tmpStr2 As String = ""
        Dim day As Int32 = chk.BetweenDay(sdt, edt)
        Dim row(day, 4) As String
        Dim i As Int32 = 0
        Dim dd As Date
        Dim dd2 As Date

        If sdt.Date > edt.Date Then
            dd = sdt
            sdt = edt
            edt = dd
        End If

        For i = 0 To day - 1 Step 1
            tmpStr = sdt.AddDays(i)
            row(i, 0) = tmpStr
        Next

        Dim ParentFlag As String = ""

        '判斷是否有下一級單位(單位主管)
        db.Open()
        Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.read() Then
            ParentFlag = "Y"
        End If
        db.Close()

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

        '判斷是否為室單位人事員
        db.Open()
        Dim ds As New DataSet()
        Dim adapter As New SqlDataAdapter()
        adapter.SelectCommand = New SqlCommand("SELECT * FROM ROLEGROUPITEM AS RGI JOIN ROLEGROUP AS RG ON RGI.Group_Uid = RG.Group_Uid WHERE Group_Name = '室單位人事員' AND employee_id = '" & user_id & "'", db)
        adapter.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then '若登入者為室單位人事員,則找出其所屬一級單位
            UnitFlag = "Y"
            strParentOrg = Org_Down.getUporg(org_uid, 1)
        End If
        db.Close()

        '系統管理員
        If Session("Role") = "1" Then

            sql = "SELECT convert(nvarchar,nSTARTTIME,111) as nSTARTTIME,convert(nvarchar,nENDTIME,111) as nENDTIME,EFORMSN,nSTHOUR,nETHOUR,ORG_UID,PANAME,nTYPE"
            sql += " FROM P_01,EMPLOYEE WHERE employee_id=PAIDNO AND PENDFLAG='E' and EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') "
            sql += " AND (nSTARTTIME >= '" + sdt.Date + " 00:00:00' AND nSTARTTIME <= '" + edt.Date + " 23:59:59' OR nENDTIME >= '" + sdt.Date + " 00:00:00' AND nENDTIME <= '" + edt.Date + " 23:59:59') "
            sql += " ORDER BY nSTARTTIME"
        Else

            If UnitFlag = "Y" Then '處/室單位人事員
                sql = "SELECT convert(nvarchar,nSTARTTIME,111) as nSTARTTIME,convert(nvarchar,nENDTIME,111) as nENDTIME,EFORMSN,nSTHOUR,nETHOUR,ORG_UID,PANAME,nTYPE"
                sql += " FROM P_01,EMPLOYEE WHERE employee_id=PAIDNO AND PENDFLAG='E' and EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') "
                sql += " AND (nSTARTTIME >= '" + sdt.Date + " 00:00:00' AND nSTARTTIME <= '" + edt.Date + " 23:59:59' OR nENDTIME >= '" + sdt.Date + " 00:00:00' AND nENDTIME <= '" + edt.Date + " 23:59:59') "
                sql += " AND ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ")"
                sql += " ORDER BY nSTARTTIME"
            ElseIf ParentFlag = "Y" Then '單位主管
                sql = "SELECT convert(nvarchar,nSTARTTIME,111) as nSTARTTIME,convert(nvarchar,nENDTIME,111) as nENDTIME,EFORMSN,nSTHOUR,nETHOUR,ORG_UID,PANAME,nTYPE"
                sql += " FROM P_01,EMPLOYEE WHERE employee_id=PAIDNO AND PENDFLAG='E' and EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') "
                sql += " AND (nSTARTTIME >= '" + sdt.Date + " 00:00:00' AND nSTARTTIME <= '" + edt.Date + " 23:59:59' OR nENDTIME >= '" + sdt.Date + " 00:00:00' AND nENDTIME <= '" + edt.Date + " 23:59:59') "
                sql += " AND ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ")"
                sql += " ORDER BY nSTARTTIME"
            Else
                sql = "SELECT convert(nvarchar,nSTARTTIME,111) as nSTARTTIME,convert(nvarchar,nENDTIME,111) as nENDTIME,EFORMSN,nSTHOUR,nETHOUR,ORG_UID,PANAME,nTYPE"
                sql += " FROM P_01,EMPLOYEE WHERE employee_id=PAIDNO AND PENDFLAG='E' and EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') "
                sql += " AND (nSTARTTIME >= '" + sdt.Date + " 00:00:00' AND nSTARTTIME <= '" + edt.Date + " 23:59:59' OR nENDTIME >= '" + sdt.Date + " 00:00:00' AND nENDTIME <= '" + edt.Date + " 23:59:59') "
                sql += " AND PAIDNO = '" & Session("user_id") & "'"
                sql += " ORDER BY nSTARTTIME"
            End If
        End If

        db.Open()
        Dim dt1 As DataTable = New DataTable("")
        Dim da1 As SqlDataAdapter = New SqlDataAdapter(sql, db)
        da1.Fill(dt1)
        db.Close()

        Dim HyperLink1 As HyperLink

        '漸層CSS
        Dim strclass As String = ""

        PlaceHolder1.Controls.Add(New LiteralControl("<table width='100%' border='3' bordercolor='#ccddee'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr><td align='center' class='RowClass' colspan='3'>" + sdt.Date + "~" + edt.Date + "</td></tr>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr><td align='center' class='RowClass' style='width:30%;'>日期</td><td align='center' class='RowClass' style='width:35%;'>上午</td><td align='center' class='RowClass' style='width:35%;'>下午</td></tr>"))

        For i = 0 To day - 1 Step 1

            '列變色
            If i Mod 2 = 0 Then
                strclass = "row_1"
            Else
                strclass = "row_2"
            End If

            PlaceHolder1.Controls.Add(New LiteralControl("<tr class='" & strclass & "'>"))

            PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' style='width:30%;'>" & row(i, 0) & DateShow(row(i, 0)) & "</td>"))

            PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' style='width:35%;'>"))

            Dim strFind1 As String = ""

            For x As Integer = 0 To dt1.Rows.Count - 1


                '起始日
                dd = dt1.Rows(x).Item(0)

                '終止日
                dd2 = dt1.Rows(x).Item(1)

                tmpStr = dd.Date
                tmpStr2 = dd2.Date

                '判斷是否為中間天數
                If CDate(tmpStr) < CDate(row(i, 0)) And CDate(tmpStr2) > CDate(row(i, 0)) Then

                    HyperLink1 = New HyperLink
                    With HyperLink1
                        .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                        .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                    End With

                    PlaceHolder1.Controls.Add(HyperLink1)

                    strFind1 = "1"

                Else

                    '起始日
                    If row(i, 0) = tmpStr Then

                        If dt1.Rows(x).Item(0) = dt1.Rows(x).Item(1) Then
                            '同一天

                            If dt1.Rows(x).Item(3) >= "08" And dt1.Rows(x).Item(4) <= "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind1 = "1"

                            ElseIf dt1.Rows(x).Item(3) >= "12" Then
                            ElseIf dt1.Rows(x).Item(3) >= "08" And dt1.Rows(x).Item(4) > "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind1 = "1"

                            End If

                        Else
                            '兩天以上

                            If dt1.Rows(x).Item(3) < "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind1 = "1"

                            ElseIf dt1.Rows(x).Item(3) >= "12" Then
                            End If
                        End If

                        '終止日
                    ElseIf row(i, 0) = tmpStr2 Then

                        If dt1.Rows(x).Item(0) <> dt1.Rows(x).Item(1) Then
                            If dt1.Rows(x).Item(4) <= "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind1 = "1"

                            ElseIf dt1.Rows(x).Item(4) > "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind1 = "1"

                            End If

                        End If

                    End If

                End If

            Next

            If strFind1 = "" Then

                PlaceHolder1.Controls.Add(New LiteralControl("　"))

            End If

            PlaceHolder1.Controls.Add(New LiteralControl("</td>"))

            '下午
            PlaceHolder1.Controls.Add(New LiteralControl("<td align='center' style='width:35%;'>"))

            Dim strFind2 As String = ""

            For x As Integer = 0 To dt1.Rows.Count - 1


                '起始日
                dd = dt1.Rows(x).Item(0)

                '終止日
                dd2 = dt1.Rows(x).Item(1)

                tmpStr = dd.Date
                tmpStr2 = dd2.Date

                '判斷是否為中間天數
                If CDate(tmpStr) < CDate(row(i, 0)) And CDate(tmpStr2) > CDate(row(i, 0)) Then

                    HyperLink1 = New HyperLink
                    With HyperLink1
                        .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                        .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                    End With

                    PlaceHolder1.Controls.Add(HyperLink1)

                    strFind2 = "1"

                Else

                    '起始日
                    If row(i, 0) = tmpStr Then

                        If dt1.Rows(x).Item(0) = dt1.Rows(x).Item(1) Then
                            '同一天

                            If dt1.Rows(x).Item(3) >= "08" And dt1.Rows(x).Item(4) <= "12" Then
                            ElseIf dt1.Rows(x).Item(3) >= "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind2 = "1"

                            ElseIf dt1.Rows(x).Item(3) >= "08" And dt1.Rows(x).Item(4) > "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind2 = "1"

                            End If

                        Else
                            '兩天以上

                            If dt1.Rows(x).Item(3) < "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind2 = "1"

                            ElseIf dt1.Rows(x).Item(3) >= "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind2 = "1"

                            End If
                        End If

                        '終止日
                    ElseIf row(i, 0) = tmpStr2 Then

                        If dt1.Rows(x).Item(0) <> dt1.Rows(x).Item(1) Then
                            If dt1.Rows(x).Item(4) <= "12" Then
                            ElseIf dt1.Rows(x).Item(4) > "12" Then

                                HyperLink1 = New HyperLink
                                With HyperLink1
                                    .Text = dt1.Rows(x).Item("PANAME") & "(" & dt1.Rows(x).Item("nTYPE") & ")<br>"
                                    .NavigateUrl = "../00/MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & dt1.Rows(x).Item(2)
                                End With

                                PlaceHolder1.Controls.Add(HyperLink1)

                                strFind2 = "1"

                            End If

                        End If

                    End If

                End If

            Next

            If strFind2 = "" Then

                PlaceHolder1.Controls.Add(New LiteralControl("　"))

            End If

            PlaceHolder1.Controls.Add(New LiteralControl("</td>"))

            PlaceHolder1.Controls.Add(New LiteralControl("</tr>"))

        Next

        PlaceHolder1.Controls.Add(New LiteralControl("</table>"))

        ShowMeeting = ""

    End Function

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        '登入此頁面自動顯示出請假人員
        'ShowMeeting()

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


    Protected Sub ImgSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        ShowMeeting()
    End Sub
End Class
