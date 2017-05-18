Imports System.Data
Imports System.Data.SqlClient
Partial Class OA_Menu
    Inherits System.Web.UI.Page

    Dim connstr, SqlTM, SqlTC As String
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'Conn連線字串
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        Dim ParentFlag As String = ""

        '判斷是否有下一級單位
        db.Open()
        Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.read() Then
            ParentFlag = "Y"
        End If
        db.Close()

        PlaceHolder1.Controls.Add(New LiteralControl("<script language='javascript'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("var defaultDiv = 1;"))
        PlaceHolder1.Controls.Add(New LiteralControl("var lastDisplay = 'div' + defaultDiv;"))

        '建立Menu標題
        If ParentFlag = "Y" Then
            SqlTM = "SELECT DISTINCT c.GroupName, c.GroupID, c.GroupImg FROM ROLES AS b INNER JOIN ROLEGROUPITEM AS a ON b.Group_uid = a.Group_Uid INNER JOIN ROLEMAPS AS c ON b.LinkNum = c.LinkNum WHERE (a.employee_id = '" & Session("user_id") & "') OR (c.GroupID = '99') ORDER BY c.GroupID"
        Else
            SqlTM = "SELECT DISTINCT c.GroupName, c.GroupID, c.GroupImg FROM ROLES AS b INNER JOIN ROLEGROUPITEM AS a ON b.Group_uid = a.Group_Uid INNER JOIN ROLEMAPS AS c ON b.LinkNum = c.LinkNum WHERE (a.employee_id = '" & Session("user_id") & "') ORDER BY c.GroupID"
        End If

        db.Open()
        Dim dt As DataTable = New DataTable("rolemaps")
        Dim da As SqlDataAdapter = New SqlDataAdapter(SqlTM, db)
        da.Fill(dt)
        db.Close()

        For i As Integer = 0 To dt.Rows.Count - 1

            '建立Menu內容
            If ParentFlag = "Y" And dt.Rows(i).Item(1) = 99 Then
                SqlTC = "SELECT c.ProgMap, c.ProgName, c.DisplayOrder FROM ROLES AS b INNER JOIN ROLEGROUPITEM AS a ON b.Group_uid = a.Group_Uid INNER JOIN ROLEMAPS AS c ON b.LinkNum = c.LinkNum WHERE (a.employee_id = '" & Session("user_id") & "') AND (c.GroupID = '" & dt.Rows(i).Item(1) & "') OR (c.LinkNum = '9') GROUP BY c.ProgMap, c.ProgName, c.DisplayOrder ORDER BY c.DisplayOrder"
            Else
                SqlTC = "SELECT c.ProgMap, c.ProgName, c.DisplayOrder FROM ROLES AS b INNER JOIN ROLEGROUPITEM AS a ON b.Group_uid = a.Group_Uid INNER JOIN ROLEMAPS AS c ON b.LinkNum = c.LinkNum WHERE (a.employee_id = '" & Session("user_id") & "') AND (c.GroupID = '" & dt.Rows(i).Item(1) & "') GROUP BY c.ProgMap, c.ProgName, c.DisplayOrder ORDER BY c.DisplayOrder"
            End If

            db.Open()
            Dim dt1 As DataTable = New DataTable("rolemaps")
            Dim da1 As SqlDataAdapter = New SqlDataAdapter(SqlTC, db)
            da1.Fill(dt1)
            db.Close()

            '將過長字串轉折
            Dim strNewProgTitle As String = dt.Rows(i).Item(0)
            If Len(strNewProgTitle) <= 8 Then
                strNewProgTitle = strNewProgTitle
            Else
                strNewProgTitle = Left(strNewProgTitle, 8) & "<BR>　　" & Mid(strNewProgTitle, 9)
            End If

            '建立Menu選項
            'PlaceHolder1.Controls.Add(New LiteralControl("myFolder" & i + 1 & "Title = '" & dt.Rows(i).Item(0) & "';"))
            PlaceHolder1.Controls.Add(New LiteralControl("myFolder" & i + 1 & "Title = '" & strNewProgTitle & "';"))
            PlaceHolder1.Controls.Add(New LiteralControl("myFolder" & i + 1 & "Image = '" & dt.Rows(i).Item(2) & "';"))
            PlaceHolder1.Controls.Add(New LiteralControl("myFolder" & i + 1 & " = new Array("))

            Dim MenuCTAll As String = ""
            Dim MunuCT As String = ""

            '取得Menu內容
            For x As Integer = 0 To dt1.Rows.Count - 1

                '將過長字串轉折
                Dim strNewProgName As String = dt1.Rows(x).Item(1)
                If Len(strNewProgName) <= 9 Then
                    strNewProgName = strNewProgName
                Else
                    strNewProgName = Left(strNewProgName, 9) & "<BR>" & Mid(strNewProgName, 10)
                End If

                MenuCTAll += """" & strNewProgName & """,""OnGoTo(this,'../" & dt1.Rows(x).Item(0) & "')"","

            Next

            MunuCT = Left(MenuCTAll, Len(MenuCTAll) - 1)

            '組合Menu內容
            PlaceHolder1.Controls.Add(New LiteralControl(MunuCT))

            PlaceHolder1.Controls.Add(New LiteralControl(");"))

        Next

        PlaceHolder1.Controls.Add(New LiteralControl("</script>"))

    End Sub
End Class
