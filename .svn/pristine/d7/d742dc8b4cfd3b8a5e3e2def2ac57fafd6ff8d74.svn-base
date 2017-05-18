Imports System.Data.SqlClient

Partial Class M_Source_00_MOA00210
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Shared strGindex As Integer


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY DATE_d DESC"

        SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit

        '重新查詢
        Dim strOrd As String

        strOrd = " ORDER BY DATE_d DESC"

        SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then
            Dim ImageButton1 As ImageButton
            Dim ImageButton2 As ImageButton
            Dim HiddenField1 As HiddenField

            ImageButton1 = e.Row.Cells(7).FindControl("ImageButton1")
            ImageButton2 = e.Row.Cells(7).FindControl("ImageButton2")
            HiddenField1 = e.Row.Cells(4).FindControl("HiddenField1")
            If IsNothing(HiddenField1) = False Then
                If HiddenField1.Value.ToLower <> user_id.ToLower Then
                    ImageButton1.Visible = False
                    ImageButton2.Visible = False
                End If
                If getAdminFlag(user_id) = True Then
                    ImageButton1.Visible = True
                    ImageButton2.Visible = True
                End If
            End If
        End If
    End Sub

    Public Function getAdminFlag(ByVal _userid As String) As Boolean
        Dim adminflag As Boolean
        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()
        Dim strSqlComd As New SqlCommand("select Group_Uid from ROLEGROUPITEM where employee_id ='" + _userid + "'", db)
        Dim dr = strSqlComd.ExecuteReader()
        adminflag = False
        Do While dr.Read()
            If dr.item("Group_Uid").ToLower = "GETK539lC0".ToLower Then
                adminflag = True
                Exit Do
            End If
        Loop
        db.Close()
        Return adminflag
    End Function



    Protected Sub GridView1_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView1.RowDeleted

        '重新查詢
        Dim strOrd As String

        strOrd = " ORDER BY DATE_d DESC"

        SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing

        '取得選擇的GridView列
        strGindex = e.NewEditIndex

        '重新查詢
        Dim strOrd As String

        strOrd = " ORDER BY DATE_d DESC"

        SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd


    End Sub

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

                If IsPostBack = False Then

                    '判斷登入者權限
                    Dim LoginCheck As New C_Public

                    If LoginCheck.LoginCheck(user_id, "MOA00210") <> "" Then
                        LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00210.aspx")
                        Response.End()
                    End If

                    '先設定起始日期
                    Dim dt As Date = Now()
                    If (SDate.Text = "") Then
                        SDate.Text = dt.AddDays(-14).Date
                    End If

                    If (EDate.Text = "") Then
                        EDate.Text = dt.Date
                    End If

                    '重新找出查詢結果
                    Dim strOrd As String = ""

                    strOrd = " ORDER BY DATE_d DESC"

                    SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Function FunShare(ByVal str As String) As String
        Try
            '轉換公告單位代號
            Dim tmpStr = Eval(str)

            If tmpStr = "0" Then
                tmpStr = "單位內部"
            ElseIf tmpStr = "1" Then
                tmpStr = "國防部全體"
            ElseIf tmpStr = "2" Then
                tmpStr = "系統公告"
            End If

            FunShare = tmpStr

        Catch ex As Exception
            FunShare = ""
        End Try
    End Function

    Public Function SQLALL(ByVal TitleVal, ByVal SDate, ByVal EDate)

        Try

            Dim strSql As String = ""
            Dim strDate As String = ""
            Dim strContent As String = ""

            '找出登入者的一級單位
            Dim strParentOrg As String = ""
            Dim Org_UP As New C_Public
            strParentOrg = Org_UP.getUporg(org_uid, 1)

            '判斷登入者權限
            If Session("Role") = "1" Then
                strSql = "SELECT ANNID_i, CONVERT (char(12), DATE_d, 111) AS DATE_d , DEPT_i, TITLE_nvc, CONTENT_nvc, uicode,  CONVERT (char(12), DATE_e, 111) AS DATE_e , ANN_Name, ANN_Phone,e.employee_id as employeeid  FROM [T_ANNOUNTCEMENT] ta join EMPLOYEE e on e.emp_chinese_name=ta.ANN_Name WHERE 1=1  "
            ElseIf Session("Role") = "2" Then
                strSql = "SELECT ANNID_i, CONVERT (char(12), DATE_d, 111) AS DATE_d , DEPT_i, TITLE_nvc, CONTENT_nvc, uicode,  CONVERT (char(12), DATE_e, 111) AS DATE_e , ANN_Name, ANN_Phone,e.employee_id as employeeid  FROM [T_ANNOUNTCEMENT] ta join EMPLOYEE e on e.emp_chinese_name=ta.ANN_Name WHERE 1=1 AND DEPT_i IN (" & Org_UP.getchildorg(strParentOrg) & ") "
            Else
                strSql = "SELECT ANNID_i, CONVERT (char(12), DATE_d, 111) AS DATE_d , DEPT_i, TITLE_nvc, CONTENT_nvc, uicode,  CONVERT (char(12), DATE_e, 111) AS DATE_e , ANN_Name, ANN_Phone,e.employee_id as employeeid  FROM [T_ANNOUNTCEMENT] ta join EMPLOYEE e on e.emp_chinese_name=ta.ANN_Name WHERE 1=1 AND DEPT_i = '" & org_uid & "'"
            End If

            strDate = " AND (DATE_d BETWEEN '" & SDate & "' AND '" & EDate & "')"

            '標題
            If TitleVal <> "" Then
                strContent = " AND TITLE_nvc LIKE '%" & TitleVal & "%' "
            End If

            If strSql <> "" Then
                SQLALL = strSql & strDate & strContent
            Else
                SQLALL = ""
            End If

        Catch ex As Exception

            SQLALL = ""

        End Try

    End Function

    Protected Sub ImgBtn2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn2.Click

        '查詢
        Dim strOrd As String

        strOrd = " ORDER BY DATE_d DESC"

        SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles GridView1.RowUpdated

        '查詢
        Dim strOrd As String

        strOrd = " ORDER BY DATE_d DESC"

        SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd


    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource1.SelectCommand = SQLALL(T_Content.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Server.Transfer("MOA00211.aspx")

    End Sub

End Class
