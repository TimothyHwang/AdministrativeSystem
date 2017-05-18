Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

<WebService(Namespace:="http://portalkm.oa.mil.tw/MOA/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class WaitApprove
    Inherits System.Web.Services.WebService

    Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

    <WebMethod()> _
    Public Function UserApproveALL(ByVal user_id As String) As DataTable

        '送出等待批核表單
        Dim sql = "Select * From V_EformFlow WHERE Status='?' and empuid='" & user_id & "'"
        db.Open()
        Dim dt2 As DataTable = New DataTable("UserApprove")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(sql, db)
        da2.Fill(dt2)

        Return dt2

        db.Close()
    End Function

    <WebMethod()> _
    Public Function WaitApproveALL(ByVal user_id As String) As DataTable

        '未批核表單
        Dim sql = "SELECT flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent, V_EformShow.PWNAME AS emp_chinese_name FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') and flowctl.empuid ='" & user_id & "'"
        db.Open()
        Dim dt2 As DataTable = New DataTable("WaitApprove")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(sql, db)
        da2.Fill(dt2)

        Return dt2

        db.Close()
    End Function

    <WebMethod()> _
    Public Function UserApproveCount(ByVal user_id As String) As DataTable

        '送出等待批核表單總數
        Dim sql = "Select Count(*) From V_EformFlow WHERE Status='?' and empuid='" & user_id & "'"
        db.Open()
        Dim dt2 As DataTable = New DataTable("UserApprove")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(sql, db)
        da2.Fill(dt2)

        Return dt2

        db.Close()
    End Function

    <WebMethod()> _
    Public Function WaitApproveCount(ByVal user_id As String) As DataTable

        '未批核表單總數
        Dim sql = "SELECT Count(*) FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') and flowctl.empuid ='" & user_id & "'"
        db.Open()
        Dim dt2 As DataTable = New DataTable("WaitApprove")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(sql, db)
        da2.Fill(dt2)

        Return dt2

        db.Close()
    End Function


End Class
