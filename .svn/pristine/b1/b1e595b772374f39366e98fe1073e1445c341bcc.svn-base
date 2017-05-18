Imports System.Data.sqlclient
Partial Class Source_00_MOA00108
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer
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

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00107") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00108.aspx")
                    Response.End()
                End If

                NoteDate.Text = Request.QueryString("Date")
                SqlDataSource1.SelectCommand = "SELECT Note_Num,NoteContent FROM P_1201 WHERE ORG_UID='" & Session("ORG_UID") & "' and NoteDate='" & NoteDate.Text & "' ORDER BY Note_Num"


            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub ImgInsert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgInsert.Click
        ErrMsg.Text = ""
        Try
            chk.CheckDataLen(NoteContent.Text, 255, "新增時：<事項>", True)
            If (Session("user_id") = "") Then
                Throw New Exception("請重新登錄系統")
            End If

            Dim sql As String
            sql = "INSERT INTO [P_1201] ([NoteDate], [NoteContent],[ORG_UID]) VALUES (@NoteDate, @NoteContent,@ORG_UID)"
            SqlDataSource1.InsertCommand = sql
            SqlDataSource1.InsertParameters.Clear()
            SqlDataSource1.InsertParameters.Add("NoteDate", NoteDate.Text)
            SqlDataSource1.InsertParameters.Add("NoteContent", NoteContent.Text)
            SqlDataSource1.InsertParameters.Add("ORG_UID", Session("ORG_UID"))
            SqlDataSource1.Insert()
        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try
    End Sub
End Class
