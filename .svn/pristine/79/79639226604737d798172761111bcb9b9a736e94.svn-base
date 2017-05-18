Imports System.Data.SqlClient
Partial Class OA_System_RoleRight
    Inherits System.Web.UI.Page

    Dim MyConnection As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
    Dim strGroup As String
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            strGroup = Request.QueryString("Group_Uid")  '群組代號

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00051") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00053.aspx")
                    Response.End()
                End If

                If Not IsPostBack Then
                    If strGroup <> "" Then

                        RightIn()
                        LeftIn()

                    End If
                End If

            End If


        Catch ex As Exception

        End Try

    End Sub

    Protected Sub RightIn()
        '已設定功能項目
        Dim objCmd As New SqlCommand("SELECT rolemaps.LinkNum As LinkNum, GroupName, ProgName,'('+GroupName+') '+ProgName a FROM rolemaps where rolemaps.LinkNum in (select LinkNum from roles where roles.Group_uid='" & strGroup & "') order by GroupID,DisPlayOrder", MyConnection)
        MyConnection.Open()
        Me.right.DataValueField = "LinkNum"
        Me.right.DataTextField = "a"
        Me.right.DataSource = objCmd.ExecuteReader
        right.DataBind()
        objCmd.Dispose()
        MyConnection.Close()
    End Sub

    Protected Sub LeftIn()
        '可設定功能項目
        Dim objCmd As New SqlCommand("SELECT rolemaps.LinkNum As LinkNum, GroupName, ProgName,'('+GroupName+') '+ProgName a FROM rolemaps where rolemaps.LinkNum not in (select LinkNum from roles where roles.Group_uid='" & strGroup & "') order by  GroupID,DisPlayOrder", MyConnection)
        MyConnection.Open()
        Me.left.DataValueField = "LinkNum"
        Me.left.DataTextField = "a"
        Me.left.DataSource = objCmd.ExecuteReader
        left.DataBind()
        objCmd.Dispose()
        MyConnection.Close()
    End Sub

    Protected Sub AllToRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AllToRight.Click
        '可設定功能項目 → 已設定功能項目
        Dim iIndexb, i As Integer
        Dim oItema As ListItem
        iIndexb = left.SelectedIndex
        If iIndexb <> -1 Then
            For i = 0 To left.Items.Count - 1
                If left.Items(i).Selected = True Then
                    oItema = New ListItem
                    oItema.Text = left.Items(i).Text
                    oItema.Value = left.Items(i).Value
                    right.Items.Add(oItema)

                    '寫入DB
                    MyConnection.Open()
                    Dim MyCommand1 As New SqlCommand("INSERT INTO ROLES(Group_uid,LinkNum) VALUES('" & strGroup & "','" & oItema.Value & "')", MyConnection)
                    MyCommand1.ExecuteNonQuery()
                    MyConnection.Close()
                End If
            Next

            '移除可設定功能項目
            For i = 0 To left.Items.Count - 1
                left.Items.Remove(left.SelectedItem)
            Next
        End If

    End Sub

    Protected Sub AllToLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AllToLeft.Click
        '可設定功能項目 ← 已設定功能項目
        Dim iIndexb, i As Integer
        Dim oItemb As ListItem
        iIndexb = right.SelectedIndex
        If iIndexb <> -1 Then
            For i = 0 To right.Items.Count - 1
                If right.Items(i).Selected = True Then
                    oItemb = New ListItem
                    oItemb.Text = right.Items(i).Text
                    oItemb.Value = right.Items(i).Value
                    left.Items.Add(oItemb)

                    '刪除未寫入DB者
                    MyConnection.Open()
                    Dim MyCommand3 As New SqlCommand("DELETE FROM ROLES WHERE Group_uid = '" & strGroup & "' and LinkNum = '" & oItemb.Value & "'", MyConnection)
                    MyCommand3.ExecuteNonQuery()
                    MyConnection.Close()
                End If
            Next

            '移除已設定功能項目
            For i = 0 To right.Items.Count - 1
                right.Items.Remove(right.SelectedItem)
            Next
        End If

    End Sub
End Class
