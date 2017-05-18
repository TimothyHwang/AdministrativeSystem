
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_03_MOA03004
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim PCK_Name As String = e.NewValues("PCK_Name")
            Dim PCK_OilLose As String = e.NewValues("PCK_OilLose")

            chk.CheckDataLen(PCK_Name, 50, "修改時：<車種>", True)
            chk.CheckDataLen(PCK_OilLose, 30, "修改時：<耗油量>", False)

        Catch ex As Exception
            e.Cancel = True
            ErrMsg.Text = ex.Message
        End Try
    End Sub

    Protected Function Chg_OilItem(ByVal str As String) As String
        Try
            Dim tmpStr = Eval(str)
            If tmpStr = "1" Then
                tmpStr = "汽油"
            ElseIf tmpStr = "2" Then
                tmpStr = "柴油"
            End If
            Chg_OilItem = tmpStr
        Catch ex As Exception
            Chg_OilItem = ""
        End Try
    End Function

    Protected Sub ImgInsert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim PCK_Name As TextBox
            Dim PCK_OilLose As TextBox

            PCK_Name = Me.DetailsView1.FindControl("PCK_Name")
            PCK_OilLose = Me.DetailsView1.FindControl("PCK_OilLose")
            chk.CheckDataLen(PCK_Name.Text, 50, "新增時：<車種>", True)
            chk.CheckDataLen(PCK_OilLose.Text, 30, "新增時：<耗油量>", False)

            Dim btnInsert As ImageButton
            btnInsert = Me.DetailsView1.FindControl("ImgInsert")
            btnInsert.CommandName = "insert"
        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try

    End Sub
End Class
