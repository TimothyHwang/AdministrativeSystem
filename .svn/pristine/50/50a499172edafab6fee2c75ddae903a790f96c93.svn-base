
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_03_MOA03007
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim Drive_Name As String = e.NewValues("Drive_Name")
            Dim Drive_Tel As String = e.NewValues("Drive_Tel")
            Dim Drive_Status As String = e.NewValues("Drive_Status")

            chk.CheckDataLen(Drive_Name, 20, "修改時：<駕駛名稱>", True)
            chk.CheckDataLen(Drive_Tel, 10, "修改時：<駕駛電話>", True)
            chk.CheckDataLen(Drive_Status, 10, "新增時：<駕駛狀況>", True)

        Catch ex As Exception
            e.Cancel = True
            ErrMsg.Text = ex.Message
        End Try
    End Sub


    Protected Function Chg_Drive_Status(ByVal str As String) As String
        Try
            Dim tmpStr = Eval(str)
            If tmpStr = "1" Then
                tmpStr = "到勤"
            ElseIf tmpStr = "2" Then
                tmpStr = "休假"
            ElseIf tmpStr = "3" Then
                tmpStr = "公差"
            End If
            Chg_Drive_Status = tmpStr
        Catch ex As Exception
            Chg_Drive_Status = ""
        End Try
    End Function

    Protected Sub ImgInsert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim Drive_Name As TextBox
            Dim Drive_Tel As TextBox
            Dim Drive_Status As DropDownList

            Drive_Name = Me.DetailsView1.FindControl("Drive_Name")
            Drive_Tel = Me.DetailsView1.FindControl("Drive_Tel")
            Drive_Status = Me.DetailsView1.FindControl("Drive_Status")
            chk.CheckDataLen(Drive_Name.Text, 20, "新增時：<駕駛名稱>", True)
            chk.CheckDataLen(Drive_Tel.Text, 10, "新增時：<駕駛電話>", True)
            chk.CheckDataLen(Drive_Status.Text, 10, "新增時：<駕駛狀況>", True)

            Dim btnInsert As ImageButton
            btnInsert = Me.DetailsView1.FindControl("ImgInsert")
            btnInsert.CommandName = "insert"
        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try


    End Sub
End Class
