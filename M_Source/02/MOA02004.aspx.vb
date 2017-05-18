
Partial Class Source_02_MOA02004
    Inherits System.Web.UI.Page
    Dim chk As New C_CheckFun
    Dim user_id, org_uid As String

    Protected Sub ImgBtn2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn2.Click

        Dim strOrd As String

        strOrd = " ORDER BY MeetName"

        SqlDataSource1.SelectCommand = SQLALL(meetname.Text, Owner.Text) & strOrd

    End Sub

    Protected Function FunShare(ByVal str As String) As String
        Try
            '轉換共用代號
            Dim tmpStr = Eval(str)

            If tmpStr = "1" Then
                tmpStr = "共用"
            ElseIf tmpStr = "2" Then
                tmpStr = "不共用"
            End If

            FunShare = tmpStr

        Catch ex As Exception
            FunShare = ""
        End Try
    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY MeetName"

        SqlDataSource1.SelectCommand = SQLALL(meetname.Text, Owner.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            '隱藏MeetSn
            e.Row.Cells(0).Visible = False
        End If

    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating

        '修改時欄位不可空白
        Dim lbErrMsg As Label = GridView1.Rows(e.RowIndex).FindControl("ErrMsg")
        Try

            chk.CheckDataLen(e.NewValues("MeetName"), 50, "修改時：<會議室名稱>", True)
            chk.CheckDataLen(e.NewValues("Tel"), 50, "修改時：<電話>", True)
            

        Catch ex As Exception
            e.Cancel = True
            lbErrMsg.Text = ex.Message
        End Try


    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        '顯示該會議室有哪些設備
        Dim strMeetsn As Integer = GridView1.Rows(GridView1.SelectedIndex).Cells(0).Text
        Server.Transfer("MOA02007.aspx?strMeetsn=" & strMeetsn)

    End Sub

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

            '系統管理員可以看全部的會議室
            If Session("Role") = "1" Then
                SqlDataSource1.SelectCommand = "SELECT MeetSn, P_0201.Org_Uid, MeetName, Owner, Tel, ContainNum, Share,emp_chinese_name,Enabled FROM P_0201,EMPLOYEE WHERE employee_id = P_0201.Owner ORDER BY MeetName"
                SqlDataSource2.SelectCommand = "SELECT employee_id, emp_chinese_name FROM V_EmpInfo ORDER BY emp_chinese_name"
            Else
                SqlDataSource1.SelectCommand = "SELECT MeetSn, P_0201.Org_Uid, MeetName, Owner, Tel, ContainNum, Share,emp_chinese_name,Enabled FROM P_0201,EMPLOYEE WHERE employee_id = P_0201.Owner AND (P_0201.Owner = '" & user_id & "') ORDER BY MeetName"
            End If

        End If

    End Sub

    Public Function SQLALL(ByVal MeetSel, ByVal OwnerSel)

        '整合SQL搜尋字串
        Dim strsql As String = ""
        Dim strsel As String = ""

        '系統管理員可以看全部的會議室
        If Session("Role") = "1" Then
            strsql = "SELECT MeetSn, P_0201.Org_Uid, MeetName, Owner, Tel, ContainNum, Share,emp_chinese_name,Enabled FROM P_0201,EMPLOYEE WHERE employee_id = P_0201.Owner "
        Else
            strsql = "SELECT * FROM P_0201,EMPLOYEE WHERE employee_id = P_0201.Owner AND P_0201.ORG_UID = '" & Session("ORG_UID") & "'"
        End If

        '會議室名稱
        If meetname.Text <> "" Then
            strsql += " and MeetName like '%" & Trim(MeetSel) & "%'"
        End If

        '管理者
        If Owner.Text <> "" Then
            strsql += " and emp_chinese_name like '%" & Trim(OwnerSel) & "%'"
        End If

        SQLALL = strsql & strsel

    End Function

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource1.SelectCommand = SQLALL(meetname.Text, Owner.Text) & strOrd

    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Server.Transfer("MOA02005.aspx")

    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim gv As GridView = CType(sender, GridView)
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim Enabled As Label = e.Row.FindControl("Label6")
                If Not IsNothing(Enabled) Then
                    Enabled.Text = IIf(Enabled.Text = "1", "是", IIf(Enabled.Text = "0", "否", "不明"))
                End If
        End Select
    End Sub
End Class
