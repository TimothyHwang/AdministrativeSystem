Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_02_MOA02007
    Inherits System.Web.UI.Page

    Dim Meetsn As Integer

    Protected Sub ChkBox1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkBox1.PreRender

        '新增資料
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '找出會議室擁有的設備
        Dim MeetDT As DataTable
        db.Open()
        Dim StepSQL As String = "select DeviceSn from P_0203 where MeetSn = " & Meetsn & " ORDER BY DeviceSn "
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
        da.Fill(ds)
        MeetDT = ds.Tables(0)
        db.Close()

        '判斷會議室設備是否有加入
        Dim i, j As Integer
        For i = 0 To MeetDT.Rows.Count - 1

            '全部會議室設備
            For j = 0 To ChkBox1.Items.Count - 1

                If MeetDT.Rows(i).Item("DeviceSn") = ChkBox1.Items(j).Value Then
                    ChkBox1.Items(j).Selected = True
                End If

            Next

        Next

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Meetsn = Request.QueryString("strMeetsn")

    End Sub

    Protected Sub ImgOK_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgOK.Click

        Try

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '先刪除該會議室相關設備
            db.Open()

            Dim delCom As New SqlCommand("DELETE P_0203 WHERE Meetsn = " & Meetsn, db)
            delCom.ExecuteNonQuery()

            db.Close()

            Dim j As Integer
            '全部會議室設備
            For j = 0 To ChkBox1.Items.Count - 1

                If ChkBox1.Items(j).Selected = True Then

                    '新增會議室設備關聯資料
                    db.Open()

                    Dim insCom As New SqlCommand("INSERT INTO P_0203(MeetSn,DeviceSn) VALUES ('" & Meetsn & "','" & ChkBox1.Items(j).Value & "')", db)
                    insCom.ExecuteNonQuery()

                    db.Close()

                End If

            Next

            Server.Transfer("MOA02004.aspx")

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA02004.aspx")

    End Sub
End Class
