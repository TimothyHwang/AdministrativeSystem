Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_03_MOA03008
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If IsPostBack = False Then

            Dim PCN_Date1 As TextBox = Me.FindControl("PCN_Date1")
            Dim PCN_Date2 As TextBox = Me.FindControl("PCN_Date2")
            Dim dt As Date = Now()

            If (PCN_Date1.Text = "") Then
                PCN_Date1.Text = dt.Date
            End If
            If (PCN_Date2.Text = "") Then
                PCN_Date2.Text = dt.AddDays(10).Date()
            End If

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        Search(False)
    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Search(False)
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)
        If (sort) Then
            GridView1.Sort("PCN_Date,PCK_Name", SortDirection.Ascending)
            Return
        End If
        Dim sql As String
        sql = "SELECT PCN_NUM,convert(char(10),PCN_Date,111) PCN_Date,PCK_Name,PCN_Use"
        sql += "  FROM P_0306"
        sql += " where P_0306.PCN_NUM=P_0306.PCN_NUM"
        If (PCN_Date1.Text <> "") Then sql += " and P_0306.PCN_Date>=CONVERT(datetime,'" + PCN_Date1.Text + " 00:00:00')" '日期
        If (PCN_Date2.Text <> "") Then sql += " and P_0306.PCN_Date<=CONVERT(datetime,'" + PCN_Date2.Text + " 23:59:59')" '日期

        SqlDataSource2.SelectCommand = sql
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Dim sDate As Date = PCN_Date1.Text
        Dim eDate As Date = PCN_Date2.Text
        Insert0306(sDate, eDate)

        Search(True)
    End Sub

    Sub Insert0306(ByVal sDate As Date, ByVal eDate As Date)
        Dim ErrMsg As Label = Me.FindControl("ErrMsg")
        Dim dDate As Date = sDate
        Dim sql As String = ""
        Dim dv As DataView

        Do While dDate.Date <= eDate.Date
            sql = "select count(*) from P_0306 where P_0306.PCN_Date=@DATE" '日期
            SqlDataSource1.SelectCommand = sql
            SqlDataSource1.SelectParameters.Clear()
            SqlDataSource1.SelectParameters.Add("DATE", dDate.Date)
            dv = CType(SqlDataSource1.Select(DataSourceSelectArguments.Empty), DataView)
            If (dv.Table.Rows(0)(0).ToString() = "0") Then
                sql = "insert into P_0306(PCN_Date,PCK_Name) select @DATE,PCK_Name from P_0302 group by PCK_Name"
                SqlDataSource1.InsertCommand = sql
                SqlDataSource1.InsertParameters.Clear()
                SqlDataSource1.InsertParameters.Add("DATE", dDate.Date)
                Try
                    SqlDataSource1.Insert()
                Catch ex As Exception
                    ErrMsg.Text = ex.Message + dDate.Date
                    Exit Sub
                End Try
            End If
            dDate = dDate.AddDays(1)
        Loop
    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "120px"

        Calendar1.SelectedDate = PCN_Date1.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "220px"

        Calendar2.SelectedDate = PCN_Date2.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        PCN_Date1.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        PCN_Date2.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

End Class
