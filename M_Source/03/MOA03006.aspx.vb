Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_03_MOA03006
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer

    Protected Sub btnInsert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim PCK_Num As DropDownList = Me.DetailsView1.FindControl("PCK_Num")
            Dim PCI_CarNumber As TextBox = Me.DetailsView1.FindControl("PCI_CarNumber")
            Dim PCI_Status As DropDownList = Me.DetailsView1.FindControl("PCI_Status")
            Dim PCI_OnlySir As TextBox = Me.DetailsView1.FindControl("PCI_OnlySir")
            Dim PCI_Use As DropDownList = Me.DetailsView1.FindControl("PCI_Use")

            If PCK_Num.Text = "-1" Then
                Throw New Exception("新增時：<車種>不可空白")

            End If
            chk.CheckDataLen(PCI_CarNumber.Text, 10, "新增時：<車號>", True)
            chk.CheckDataLen(PCI_Status.Text, 10, "新增時：<車輛類型>", True)
            chk.CheckDataLen(PCI_OnlySir.Text, 10, "新增時：<專用車長官>", False)
            chk.CheckDataLen(PCI_Use.Text, 10, "新增時：<使用狀況>", True)

            Dim btnInsert As Button
            btnInsert = Me.DetailsView1.FindControl("btnInsert")
            btnInsert.CommandName = "insert"
        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try
    End Sub
End Class
