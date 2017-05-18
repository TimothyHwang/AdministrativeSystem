Imports System.Data.sqlclient
Partial Class Source_00_MOA00033
    Inherits System.Web.UI.Page

    Public Neworg_uid As String = ""
    Dim user_id, org_uid As String
    Dim connstr, UseFlag, PARENT_ORG_UID, ORG_NAME, ORG_NAME_E As String
    Dim ORG_TREE_LEVEL, ORG_KIND As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            user_id = Session("user_id")
            org_uid = Request.QueryString("org_uid")

            '新增=1,修改=2
            UseFlag = Request.QueryString("UseFlag")

            If UseFlag = "1" Then
                LabelTitle.Text = "單位新增"
            Else
                LabelTitle.Text = "單位修改"
            End If

            'session被清空回首頁
            If user_id = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00030") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00033.aspx")
                    Response.End()
                End If

                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '單位資料
                db.Open()
                Dim strPer As New SqlCommand("SELECT * FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    PARENT_ORG_UID = RdPer("PARENT_ORG_UID")
                    ORG_NAME = RdPer("ORG_NAME")

                    If RdPer("ORG_NAME_E") Is DBNull.Value = False Then
                        ORG_NAME_E = RdPer("ORG_NAME_E")
                    End If

                    ORG_TREE_LEVEL = RdPer("ORG_TREE_LEVEL")
                    ORG_KIND = RdPer("ORG_KIND")
                End If
                db.Close()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub btnImgOK_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgOK.Click

        Try
            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string
            Dim ShowMsg As String = ""

            If UseFlag = "1" Then

                If txtorg_id.Text = "" Or txtorg_name.Text = "" Or org_uid = "" Then

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('請輸入單位代號或是單位名稱');")
                    Response.Write(" </script>")

                Else

                    '開啟連線
                    Dim db As New SqlConnection(connstr)

                    '新增組織資料
                    db.Open()
                    Dim insCom As New SqlCommand("INSERT INTO ADMINGROUP(ORG_UID,PARENT_ORG_UID,ORG_NAME,ORG_NAME_E,ORG_TREE_LEVEL) VALUES ('" & txtorg_id.Text & "','" & org_uid & "','" & txtorg_name.Text & "','" & txtorg_enname.Text & "','" & ORG_TREE_LEVEL + 1 & "')", db)
                    insCom.ExecuteNonQuery()
                    db.Close()

                    ShowMsg = "單位新增完成"

                    'Server.Transfer("MOA00032.aspx")

                    '清空選擇的組織級別
                    DropOrgKind.Items(0).Selected = False
                    DropOrgKind.Items(1).Selected = False
                    DropOrgKind.Items(2).Selected = False

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('" & ShowMsg & "');")
                    Response.Write(" window.parent.location='MOA00030.aspx?Sel_Org=" & org_uid & "';")
                    Response.Write(" </script>")

                End If

            Else

                If txtorg_id.Text = "" Or txtorg_name.Text = "" Then

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('請輸入單位代號或是單位名稱');")
                    Response.Write(" </script>")

                Else

                    '開啟連線
                    Dim db As New SqlConnection(connstr)

                    '修改組織資料
                    db.Open()
                    Dim updCom As New SqlCommand("UPDATE ADMINGROUP SET ORG_NAME='" & txtorg_name.Text & "',ORG_NAME_E='" & txtorg_enname.Text & "',PARENT_ORG_UID='" & DropOrgUp.SelectedValue & "',ORG_KIND='" & DropOrgKind.SelectedValue & "',MODIFY_DATE=getdate() WHERE ORG_UID = '" & org_uid & "'", db)
                    updCom.ExecuteNonQuery()
                    db.Close()

                    ShowMsg = "單位修改完成"

                    'Server.Transfer("MOA00032.aspx")

                    '清空選擇的組織級別
                    DropOrgKind.Items(0).Selected = False
                    DropOrgKind.Items(1).Selected = False
                    DropOrgKind.Items(2).Selected = False

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('" & ShowMsg & "');")
                    Response.Write(" window.parent.location='MOA00030.aspx?Sel_Org=" & org_uid & "';")
                    Response.Write(" </script>")

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA00032.aspx?org_uid=" & org_uid)

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        Try
            If IsPostBack = False Then

                If UseFlag = "1" Then

                    '取得亂數的org_id
                    Dim randstr As New C_Public
                    Neworg_uid = randstr.randstr(10)

                    txtorg_id.Text = Neworg_uid

                    txtorg_id.Enabled = False

                Else

                    txtorg_id.Text = org_uid
                    txtorg_name.Text = ORG_NAME
                    txtorg_enname.Text = ORG_NAME_E

                    txtorg_id.Enabled = False

                    '判斷組織級別
                    If ORG_KIND = "1" Then
                        DropOrgKind.Items(1).Selected = True
                    ElseIf ORG_KIND = "2" Then
                        DropOrgKind.Items(2).Selected = True
                    End If

                    Dim i As Integer
                    For i = 0 To DropOrgUp.Items.Count - 1

                        If CStr(DropOrgUp.Items(i).Value) = CStr(PARENT_ORG_UID) Then
                            DropOrgUp.Items(i).Selected = True
                        End If

                    Next

                End If


            End If


        Catch ex As Exception

        End Try

    End Sub
End Class
