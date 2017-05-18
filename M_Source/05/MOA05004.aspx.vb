Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_05_MOA05004
    Inherits System.Web.UI.Page

    Dim EFORMSN As String
    Public Shared HistoryGo As Int16
    Dim TheSameDay As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        EFORMSN = Request.QueryString("EFORMSN")

        Dim SelDate As Date = Request.QueryString("SelDate")

        '查詢日期不是當天不可修改換證資料
        If (DateDiff(DateInterval.Day, Now(), SelDate)) <> 0 Then
            TheSameDay = 1
        End If

        Dim sql As String

        sql = "SELECT EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME,PATITLE,PAIDNO,nAPPTIME,nRECROOM,nRECEXIT,nPLACE,nPHONE,convert(nvarchar,nRECDATE,111) nRECDATE,nSTARTTIME,nENDTIME,nREASON,nSTATUS"
        sql += " FROM P_05"
        sql += " where EFORMSN='" + EFORMSN + "'"

        SqlDataSource1.SelectCommand = sql

        sql = "SELECT EFORMSN,Receive_Num,nName,nSex,nService,nCarNo,nID,nPassID,nInDate,nLeaveDate,nCreateDate"
        sql += " FROM P_0501"
        sql += " where EFORMSN='" + EFORMSN + "'"

        SqlDataSource2.SelectCommand = sql
        If IsPostBack Then
            HistoryGo -= 1
            Exit Sub
        End If
        HistoryGo = -1
    End Sub

    Protected Sub btn_nInDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        CType(GridView1.Rows(GridView1.EditIndex).FindControl("nInDate"), Label).Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub

    Protected Sub btn_nLeaveDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        CType(GridView1.Rows(GridView1.EditIndex).FindControl("nLeaveDate"), Label).Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        'If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
        If e.Row.RowType = DataControlRowType.Header Then

            '隱藏按鈕
            If TheSameDay = 1 Then
                'e.Row.Cells(8).Visible = False
            End If

        End If

    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim sql As String

        sql = "UPDATE [P_05]"
        sql += " SET [nSTATUS] = @nSTATUS"
        sql += " WHERE [EFORMSN] = @EFORMSN"
        If (e.NewValues("nLeaveDate") <> "") Then
            SqlDataSource1.UpdateCommand = sql
            SqlDataSource1.UpdateParameters.Add("EFORMSN", e.Keys("EFORMSN"))
            SqlDataSource1.UpdateParameters.Add("nSTATUS", "2")
            SqlDataSource1.Update()
        ElseIf (e.NewValues("nInDate") <> "") Then
            sql += " and [nSTATUS] <> '2'"
            SqlDataSource1.UpdateCommand = sql
            SqlDataSource1.UpdateParameters.Add("EFORMSN", e.Keys("EFORMSN"))
            SqlDataSource1.UpdateParameters.Add("nSTATUS", "1")
            SqlDataSource1.Update()
        End If

        sql = "UPDATE [P_0501]"
        sql += " SET [nPassID] = @nPassID"
        sql += ", [nInDate] = @nInDate"
        sql += ",[nLeaveDate]=@nLeaveDate"
        sql += " WHERE [EFORMSN] = @EFORMSN"
        sql += " and [Receive_Num] = @Receive_Num"

        SqlDataSource2.UpdateCommand = sql

    End Sub
End Class
