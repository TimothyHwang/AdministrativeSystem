
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_04_MOA04023
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim House_Name As String = e.NewValues("House_Name")
            Dim House_Tel As String = e.NewValues("House_Tel")
            Dim House_Status As String = e.NewValues("House_Status")

            chk.CheckDataLen(House_Name, 20, "修改時：<現勘人員名稱>", True)
            chk.CheckDataLen(House_Tel, 10, "修改時：<現勘人員電話>", True)
            chk.CheckDataLen(House_Status, 10, "新增時：<現勘人員狀況>", True)

        Catch ex As Exception
            e.Cancel = True
            ErrMsg.Text = ex.Message
        End Try
    End Sub


    Protected Function Chg_House_Status(ByVal str As String) As String
        Try
            Dim tmpStr = Eval(str)
            If tmpStr = "1" Then
                tmpStr = "內派"
            ElseIf tmpStr = "2" Then
                tmpStr = "外包"
            End If
            Chg_House_Status = tmpStr
        Catch ex As Exception
            Chg_House_Status = ""
        End Try
    End Function

    Protected Sub ImgInsert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim House_Name As TextBox
            Dim House_Tel As TextBox
            Dim House_Status As DropDownList

            House_Name = Me.DetailsView1.FindControl("House_Name")
            House_Tel = Me.DetailsView1.FindControl("House_Tel")
            House_Status = Me.DetailsView1.FindControl("House_Status")
            chk.CheckDataLen(House_Name.Text, 20, "新增時：<現勘人員名稱>", True)
            chk.CheckDataLen(House_Tel.Text, 10, "新增時：<現勘人員電話>", True)
            chk.CheckDataLen(House_Status.Text, 10, "新增時：<現勘人員狀況>", True)

            Dim btnInsert As ImageButton
            btnInsert = Me.DetailsView1.FindControl("ImgInsert")
            btnInsert.CommandName = "insert"
        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try


    End Sub
End Class
