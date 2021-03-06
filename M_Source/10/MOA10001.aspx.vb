﻿Imports System.Drawing
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration


Partial Class M_Source_10_MOA10001
    Inherits Page
    Protected Conn_E As String = WebConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString
#Region "Custom Function"
    ''**********************  以下為Custom Function  **************************
#End Region

#Region "Form Function"
    ''**********************  以下為  Form Function  **************************
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then

        End If
        ''繫結Header
        Page.Header.DataBind()
    End Sub

    Protected Sub gvManager_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvManager.RowDataBound
        ''變更顯示資料
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim lblStatus As Label = CType(e.Row.FindControl("lblStatus"), Label)
                If lblStatus Is Nothing Then Return
                Select Case lblStatus.Text
                    Case "0"
                        lblStatus.Text = "未在營"
                        lblStatus.ForeColor = Color.Red
                    Case "1"
                        lblStatus.Text = "在營"
                End Select
        End Select
    End Sub

    Protected Sub gvManager_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles gvManager.RowCommand
        ''繫結按鈕事件
        Select Case e.CommandName
            Case "Add"
            Case "Select"
                'Dim pk_index As Integer = CInt(e.CommandArgument)
                'Dim strManagerName = CType(sender, GridView).Rows(pk_index).Cells(2).Text
                'Response.Redirect("MOA10003.aspx?MID=" & gvManager.DataKeys(pk_index).Value.ToString() & "&MName=" & strManagerName)
            Case "ChangeSort"
                Dim sort As String = e.CommandArgument.ToString() ''點選該列的sort數字
                Dim DownUp As String = CType(e.CommandSource, ImageButton).ToolTip ''下移或上移的Command

                ''region 取得要交換sort的兩筆資料
                Dim sql As String = CType((" SELECT Top 2 [P_Num],[RankId] " _
                                           & " FROM [P_10]" _
                                           & " Where RankId" + IIf(DownUp = "Down", ">=", "<=") + sort _
                                           & " Order by RankId " + IIf(DownUp = "Down", "ASC", "DESC")), String)

                ''以上是呈現端Order by sort ASC的語法
                ''若呈現端Order by sort DESC的話，條件要顛倒


                Dim dt As DataTable = queryDataTable(sql)

                If (dt.Rows.Count >= 2) Then ''防呆第一列上移動作和最後一列下移動作

                    Dim firstID As String = dt.Rows(0)("P_Num").ToString()
                    Dim firstSort As String = dt.Rows(0)("RankId").ToString()

                    Dim secondID As String = dt.Rows(1)("P_Num").ToString()
                    Dim secondSort As String = dt.Rows(1)("RankId").ToString()

                    '*將第一筆sort換成第二筆*'
                    '*將第二筆sort換成第一筆*'
                    Dim s As String = <s>SET XACT_ABORT ON 
                        Begin Transaction
                        Update P_10 Set RankId='@secondSort@' Where P_Num = '@firstID@' 
                        Update P_10 Set RankId='@firstSort@' Where P_Num = '@secondID@' 
                        Commit Transaction 
                                    </s>.Value.Replace(vbLf, vbCrLf)

                    s = s.Replace("@firstID@", firstID)
                    s = s.Replace("@firstSort@", firstSort)
                    s = s.Replace("@secondID@", secondID)
                    s = s.Replace("@secondSort@", secondSort)
                    ExecuteNonQuery(s)
                End If
                gvManager.DataBind()
            Case "Delete"
                Dim pid As String = e.CommandArgument.ToString()

                ''region 取得要交換sort的兩筆資料
                Dim sql As String = <s>SET XACT_ABORT ON
                                        Begin Transaction
                                        DELETE FROM P_10 Where P_Num='@pid@'
                                        DELETE FROM P_1001 WHERE MANAGER_ID='@pid@'
                                        Commit Transaction 
                                    </s>.Value.Replace(vbLf, vbCrLf)
                sql = sql.Replace("@pid@", pid)
                ExecuteNonQuery(sql)
                'Case "Update"
                '    Dim pid As String = e.CommandArgument.ToString()
                '    Dim Name As TextBox = gvManager.Rows(pid).FindControl("")
                '    Dim sql As String = <s>SET XACT_ABORT ON
                '                            Begin Transaction
                '                            UPDATE P_10 SET NAME='@name@' WHERE P_Num='@pid@'                                        
                '                            Commit Transaction 
                '                        </s>.Value.Replace(vbLf, vbCrLf)
                '    sql = sql.Replace("@pid@", pid)
                '    ExecuteNonQuery(sql)
        End Select
    End Sub
    Protected Function queryDataTable(sql As String) As DataTable
        Using conn As New SqlConnection(Conn_E)
            Dim da As SqlDataAdapter = New SqlDataAdapter(sql, conn)
            Dim ds As DataSet = New DataSet()
            da.Fill(ds)
            Return CType(IIf(ds.Tables.Count > 0, ds.Tables(0), New DataTable), DataTable)
        End Using
    End Function

    Protected Function ExecuteNonQuery(sql As String) As Integer
        Using conn As SqlConnection = New SqlConnection(Conn_E)
            Dim cmd As SqlCommand = New SqlCommand(sql, conn)
            Dim rows As Integer
            conn.Open()
            rows = cmd.ExecuteNonQuery()
            Return rows
        End Using
    End Function

    Protected Sub ibnManagerAdd_Click(sender As Object, e As ImageClickEventArgs) Handles ibnManagerAdd.Click
        Response.Redirect("MOA10002.aspx")
    End Sub
#End Region

End Class
