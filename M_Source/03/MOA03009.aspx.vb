Imports System.Data.sqlclient

Partial Class M_Source_03_MOA03009
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Dim connstr As String = ""
    Dim CarFlag As String = ""
    Dim user_id, org_uid As String
    Dim nRECDATE_Flag As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            nRECDATE_Flag = Request("nRECDATE_Flag")

            '新增連線
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷登入者是否為派車相關群組人員
            db.Open()
            Dim carCom As New SqlCommand("select * from ROLEGROUPITEM where employee_id = '" & user_id & "' AND (Group_Uid ='C6B45GAHD3')", db)
            Dim RdvCar = carCom.ExecuteReader()
            If RdvCar.read() Then
                CarFlag = "1"
            End If
            db.Close()

            If nRECDATE_Flag = "" Then
                '先設定起始日期
                Dim dt As Date = Now()
                If (nRECDATE1.Text = "") Then
                    nRECDATE1.Text = dt.AddDays(0).Date
                End If
            Else
                nRECDATE1.Text = nRECDATE_Flag
            End If

            If Not IsPostBack Then
                SqlDataSource2.SelectCommand = SQLALL(nRECDATE1.Text, DDLCar.SelectedValue)
            End If
        End If


    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "90px"

        Calendar1.SelectedDate = nRECDATE1.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        nRECDATE1.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        SqlDataSource2.SelectCommand = SQLALL(nRECDATE1.Text, DDLCar.SelectedValue)

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '隱藏eformsn
            e.Row.Cells(0).Visible = False

        End If

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        'Dim strPath As String = "MOA03001.aspx?eformid=j2mvKYe3l9&Read_Only=2&EFORMSN=" & GridView1.SelectedValue

        'Response.Write(" <script language='javascript'>")
        'Response.Write(" sPath = '" & strPath & "';")
        'Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=700px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
        'Response.Write(" showModalDialog(sPath,self,strFeatures);")
        'Response.Write(" </script>")

        'Server.Transfer("../00/MOA00020.aspx?x=MOA03001&y=j2mvKYe3l9&Read_Only=2&EFORMSN=" & GridView1.SelectedValue & "&CarFlag=1")

        Server.Transfer("MOA03001.aspx?eformid=j2mvKYe3l9&Read_Only=2&EFORMSN=" & GridView1.SelectedValue & "&nRECDATE_Flag=" & nRECDATE1.Text)

    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        SqlDataSource2.SelectCommand = SQLALL(nRECDATE1.Text, DDLCar.SelectedValue)

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(nRECDATE1.Text, DDLCar.SelectedValue) & strOrd

    End Sub

    Public Function SQLALL(ByVal nRECDATE1, ByVal CarSel)

        Dim sqlcom As String = "SELECT flowctl.appdate, flowctl.eformsn, flowctl.hddate,CONVERT (char(12), P_03.nARRDATE, 111) AS nARRDATE , P_03.PANAME, P_03.PAUNIT, P_03.nARRIVEPLACE, P_03.nREASON,flowctl.stepsid FROM flowctl INNER JOIN P_03 ON flowctl.eformsn = P_03.EFORMSN WHERE (flowctl.hddate IS NULL) AND (flowctl.gonogo = '?' OR flowctl.gonogo = 'R') AND (flowctl.empuid = '" & user_id & "') AND (flowctl.eformid = 'j2mvKYe3l9')"

        Dim sqlDate As String = ""
        Dim sqlCar As String = ""

        '報到日期
        If nRECDATE1 <> "" Then
            '報到日期搜尋
            sqlDate = "  AND nARRDATE='" & nRECDATE1 & "'"
        End If

        '車種
        If CarSel <> "" Then
            '報到日期搜尋
            sqlCar = "  AND nSTYLE='" & CarSel & "'"
        End If

        SQLALL = sqlcom & sqlDate & sqlCar

    End Function

    Protected Function FunStatus(ByVal str As String) As String

        '轉換表單狀態代號
        Dim tmpStr = Eval(str)

        If tmpStr = "19693" Then
            tmpStr = "調度一"
        ElseIf tmpStr = "9731" Then
            tmpStr = "調度二"
        End If

        FunStatus = tmpStr

    End Function

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then

            Dim tLItm As New ListItem("請選擇", "")

            '人員加請選擇
            DDLCar.Items.Insert(0, tLItm)

        End If

    End Sub
End Class
