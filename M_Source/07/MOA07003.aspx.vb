Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Hosting

Partial Class Source_07_MOA07003
    Inherits System.Web.UI.Page

    Dim EFORMSN As String
    Dim sql As String
    Public Shared HistoryGo As Int16

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        EFORMSN = Request.QueryString("EFORMSN")

        sql = "SELECT PWUNIT, PWTITLE, PWNAME, PWIDNO, PAUNIT, PANAME, PATITLE, PAIDNO, nAPPTIME, nTel, nSeat, PENDFLAG"
        sql += " FROM P_07"
        sql += " where EFORMSN='" + EFORMSN + "'"

        SqlDataSource1.SelectCommand = sql

        sql = "SELECT nAssetNum, nAssetName, nBios, nLabel, nLabelNum, nAmount, nREASON,nContent"
        sql += " FROM P_0701"
        sql += " where EFORMSN='" + EFORMSN + "'"

        SqlDataSource2.SelectCommand = sql

        If IsPostBack Then
            HistoryGo -= 1
            Exit Sub
        End If
        HistoryGo = -1
    End Sub
End Class
