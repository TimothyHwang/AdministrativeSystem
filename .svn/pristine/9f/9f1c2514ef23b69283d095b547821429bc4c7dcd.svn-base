
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_03_MOA03005
    Inherits System.Web.UI.Page

    Dim chk As New C_CheckFun
    Dim len As New Integer

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        Search(False)
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        'Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Try
            Dim PCK_Num As String = e.NewValues("PCK_Num")
            Dim PCI_CarNumber As String = e.NewValues("PCI_CarNumber")
            Dim PCI_Status As String = e.NewValues("PCI_Status")
            Dim PCI_OnlySir As String = e.NewValues("PCI_OnlySir")
            Dim PCI_Use As String = e.NewValues("PCI_Use")

            If PCK_Num = "-1" Then
                Throw New Exception("修改時：<車種>不可空白")
            End If
            chk.CheckDataLen(PCI_CarNumber, 10, "修改時：<車號>", True)
            chk.CheckDataLen(PCI_Status, 10, "修改時：<車輛類型>", True)
            chk.CheckDataLen(PCI_OnlySir, 10, "修改時：<專用車長官>", False)
            chk.CheckDataLen(PCI_Use, 10, "修改時：<使用狀況>", True)

        Catch ex As Exception
            e.Cancel = True
            'ErrMsg.Text = ex.Message
        End Try
    End Sub

    Protected Function Chg_Status(ByVal str As String) As String
        Try
            Dim tmpStr = Eval(str)
            If tmpStr = "1" Then
                tmpStr = "一般車"
            ElseIf tmpStr = "2" Then
                tmpStr = "經常性支援"
            ElseIf tmpStr = "3" Then
                tmpStr = "長官車"
            End If
            Chg_Status = tmpStr
        Catch ex As Exception
            Chg_Status = ""
        End Try
    End Function

    Protected Function Chg_Use(ByVal str As String) As String
        Try
            Dim tmpStr = Eval(str)
            If tmpStr = "1" Then
                tmpStr = "待命"
            ElseIf tmpStr = "2" Then
                tmpStr = "派遣"
            ElseIf tmpStr = "3" Then
                tmpStr = "維修"
            End If
            Chg_Use = tmpStr
        Catch ex As Exception
            Chg_Use = ""
        End Try
    End Function

    Protected Sub btnSelect_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnSelect.Click

        'Dim ErrMsg As Label = Me.DetailsView1.FindControl("ErrMsg")
        Search(True)

    End Sub

    Sub Search(ByVal sort As Boolean)
        If (sort) Then
            GridView1.Sort("PCI_Num", SortDirection.Ascending)
            Return
        End If

        Try
            Dim sql As String

            sql = "SELECT [PCI_Num], [P_0303].[PCK_Num], [PCI_CarNumber], [PCI_Status], [PCI_OnlySir],[PCK_Name],[PCI_Use]"
            sql += " FROM [P_0303],[P_0302]"
            sql += " WHERE [P_0303].[PCK_Num]=[P_0302].[PCK_Num]"
            If PCK_Num.Text <> "-1" Then
                sql += " AND [P_0303].[PCK_Num]=" + Trim(PCK_Num.Text)
            End If
            If PCI_CarNumber.Text <> "" Then
                sql += " AND [PCI_CarNumber] like '%" + Trim(PCI_CarNumber.Text) + "%'"
            End If
            If PCI_Status.Text <> "" Then
                sql += " AND [PCI_Status]=" + Trim(PCI_Status.Text)
            End If
            If PCI_OnlySir.Text <> "" Then
                sql += " AND [PCI_OnlySir] like '%" + Trim(PCI_OnlySir.Text) + "%'"
            End If
            If PCI_Use.Text <> "" Then
                sql += " AND [PCI_Use] like '%" + Trim(PCI_Use.Text) + "%'"
            End If
            sql += " ORDER BY [PCI_Num]"

            SqlDataSource1.SelectCommand = sql

        Catch ex As Exception
            'ErrMsg.Text = ex.Message
        End Try

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Server.Transfer("MOA03006.aspx")

    End Sub
End Class
