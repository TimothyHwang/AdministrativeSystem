Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Hosting

Partial Class Source_06_MOA06003
    Inherits System.Web.UI.Page

    Dim EFORMSN As String
    Dim sql As String
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim D2 As New System.Data.DataTable
    Dim D1 As New System.Data.DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        EFORMSN = Request.QueryString("EFORMSN")

        sql = "SELECT PWUNIT, PWTITLE, PWNAME, PWIDNO, PAUNIT, PANAME, PATITLE, PAIDNO, nAPPTIME, nREASON, convert(nvarchar,nDATE,111) nDATE, nPLACE"
        sql += " FROM P_06"
        sql += " where EFORMSN='" + EFORMSN + "'"

        SqlDataSource1.SelectCommand = sql

        sql = "SELECT Info_Num, nKind, nMName, nDocnum, nAmount, nContent, nClass"
        sql += " FROM P_0601"
        sql += " where EFORMSN='" + EFORMSN + "'"

        SqlDataSource2.SelectCommand = sql
    End Sub

    Protected Sub Prt_06003(ByVal prt As C_Xprint)
        sql = "SELECT PWUNIT, PWTITLE, PWNAME, PWIDNO, PAUNIT, PANAME, PATITLE, PAIDNO, nAPPTIME, nREASON, convert(nvarchar,nDATE,111) nDATE, nPLACE"
        sql += " FROM P_06"
        sql += " where EFORMSN='" + EFORMSN + "'"
        If do_sql.db_sql(sql, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            D2 = do_sql.G_table
        Else
            Exit Sub
        End If
        sql = "SELECT Info_Num, nKind, nMName, nDocnum, nAmount, nContent, nClass"
        sql += " FROM P_0601"
        sql += " where EFORMSN='" + EFORMSN + "' order by 2"
        If do_sql.db_sql(sql, do_sql.G_conn_string) = False Then
            Exit Sub
        End If

        Dim h2 As Int16 = 93
        Dim i As Int16
        Dim append_file As Boolean = False
        Dim hi As Integer
        For i = 0 To 2
            D1 = do_sql.G_table

            prt.Add("申請出入日期", CDate(D2.Rows(0).Item("nDATE").ToString()).ToString("yyyy/MM/dd"), 0, i * h2)
            prt.Add("申請人", D2.Rows(0).Item("PANAME").ToString(), 0, i * h2)
            If D2.Rows(0).Item("nREASON").ToString() = "其他" Then
                prt.Add("事由_其他", "V", 0, i * h2)
            Else
                prt.Add(D2.Rows(0).Item("nREASON").ToString(), "V", 0, i * h2)
            End If
            prt.Add("使用地點", D2.Rows(0).Item("nPLACE").ToString(), 0, i * h2)

            Dim block_name As String = ""
            Dim si As String = ""
            Dim tabled(7) As Integer
            For hi = 1 To 6
                tabled(hi) = 0
            Next
            For hi = 0 To D1.Rows.Count - 1
                Select Case D1.Rows(hi).Item("nKind").ToString()
                    Case "電腦主機"
                        si = "1"
                    Case "磁片光碟"
                        si = "2"
                    Case "行動硬碟"
                        si = "3"
                    Case "文書圖表"
                        si = "4"
                    Case "保密裝備"
                        si = "5"
                    Case Else
                        si = "6"
                End Select
                tabled(CInt(si)) = tabled(CInt(si)) + 1
                If tabled(CInt(si)) < 2 Then
                    If D1.Rows(hi).Item("nKind").ToString() = "其他" Then
                        prt.Add("區分_其他", "V", 0, i * h2)
                    Else
                        prt.Add(D1.Rows(hi).Item("nKind").ToString(), "V", 0, i * h2)
                    End If
                    prt.Add("名稱機型_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nMName").ToString(), 0, i * h2)
                    prt.Add("編（文）號_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nDocnum").ToString(), 0, i * h2)
                    prt.Add("數量_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nAmount").ToString(), 0, i * h2)
                    prt.Add("內容概要_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nContent").ToString(), 0, (i * h2) - 1)
                    prt.Add("機密等級_" & D1.Rows(hi).Item("nClass").ToString() & "_" & si, "V", 0, i * h2)
                Else
                    append_file = True
                End If
            Next
        Next
        '附件
        If append_file = True Then
            prt.NewPage("rpt060031.txt")

            Dim h As Integer = 8
            Dim row As Integer = 0
            For hi = 0 To D1.Rows.Count - 1
                prt.Add("區分", D1.Rows(hi).Item("nKind").ToString(), 0, row * h)
                prt.Add("名稱機型", D1.Rows(hi).Item("nMName").ToString(), 0, row * h)
                prt.Add("編（文）號", D1.Rows(hi).Item("nDocnum").ToString(), 0, row * h)
                prt.Add("數量", Space(9 - D1.Rows(hi).Item("nAmount").ToString().Length) + D1.Rows(hi).Item("nAmount").ToString(), 0, row * h)
                prt.Add("內容概要", D1.Rows(hi).Item("nContent").ToString(), 0, row * h)
                prt.Add("機密等級", D1.Rows(hi).Item("nClass").ToString(), 0, row * h)
                row = row + 1
            Next
        End If
    End Sub

    Protected Sub ImgPrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgPrint.Click
        Dim filename As String = "prt_06003" & Rnd() & ".drs"
        Dim print As New C_Xprint
        print.C_Xprint("rpt060030.txt", filename)
        print.NewPage()
        Prt_06003(print)
        print.EndFile()
        If (print.ErrMsg <> "") Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('" & print.ErrMsg & "');")
            Response.Write("</script>")
        Else
            Response.Write("<script language='javascript'>")
            Response.Write("window.onload = function() {")
            Response.Write("window.location.replace('../../drs/" & filename & "');")
            Response.Write("}")
            Response.Write("</script>")
        End If
    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA06002.aspx")

    End Sub
End Class
