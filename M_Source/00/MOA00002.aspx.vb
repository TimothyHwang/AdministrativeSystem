Imports System.Data
Imports System.Data.sqlclient
Partial Class Source_00_MOA00002
    Inherits System.Web.UI.Page

    Public eformid As String = ""
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            '取得亂數的eformid
            Dim randstr As New C_Public
            eformid = randstr.randstr(10)

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA00001") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00002.aspx")
                Response.End()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub NextView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NextView.Click

        '新增資料
        Dim connstr, frm_chinese_name, frm_english_name, PRIMARY_TABLE As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        frm_chinese_name = Request.Form("frm_chinese_name")
        frm_english_name = Request.Form("frm_english_name")
        PRIMARY_TABLE = Request.Form("PRIMARY_TABLE")

        '表單名稱不可空白
        If frm_chinese_name = "" Or PRIMARY_TABLE = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單名稱或表單資料表不可空白');")
            Response.Write(" </script>")

        Else

            '新增TEMPEFORMS基本資料
            db.Open()
            Dim insCom As New SqlCommand("INSERT INTO TEMPEFORMS(eformid,frm_chinese_name,frm_english_name,organization_id,PRIMARY_TABLE) VALUES ('" & eformid & "','" & frm_chinese_name & "','" & frm_english_name & "','MOA','" & PRIMARY_TABLE & "')", db)
            insCom.ExecuteNonQuery()
            db.Close()

            '新增TEMPEFORMS開始基本資料
            db.Open()

            Dim sql1 = "insert into tempflow (eformid,eformrole,stepsid,steps,nextstep,major_step,pors,x,y,group_id,group_type,allhandle,canjump,canback,canadd,candist,candistcc,canaddatt,candelatt,canconti,bypass,opinonly,block,layer,tree,condi) values "
            sql1 += "('" & eformid & "','1','1','0','-1','1','S','300','200','0','1','0','0','0','0','0','0','1','1','0','0','0','0','0','0','0')"

            Dim instempflow1 As New SqlCommand(sql1, db)
            instempflow1.ExecuteNonQuery()
            db.Close()

            '新增TEMPEFORMS結束基本資料
            db.Open()

            Dim sql2 = "insert into tempflow (eformid,eformrole,stepsid,steps,nextstep,major_step,pors,x,y,group_id,group_type,allhandle,canjump,canback,canadd,candist,candistcc,canaddatt,candelatt,canconti,bypass,opinonly,block,layer,tree,condi) values "
            sql2 += "('" & eformid & "','1','2','-1','-1','1','S','300','100','-1','1','0','0','0','0','0','0','1','1','0','0','0','0','0','0','0')"

            Dim instempflow2 As New SqlCommand(sql2, db)
            instempflow2.ExecuteNonQuery()
            db.Close()

            Dim streformid As String = ""
            streformid = eformid & ",MOA," & frm_chinese_name & ",1"

            Server.Transfer("MOA00003.aspx?eformid=" + streformid)

        End If


    End Sub

    Protected Sub PreView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PreView.Click

        Server.Transfer("MOA00001.aspx")

    End Sub
End Class
