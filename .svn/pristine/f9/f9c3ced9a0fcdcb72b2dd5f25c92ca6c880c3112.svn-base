Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_03_MOA03010
    Inherits System.Web.UI.Page

    Dim connstr, user_id, org_uid As String
    Dim streformsn As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        streformsn = Request.QueryString("eformsn")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '判斷選擇哪個填表人
            Try

                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                db.Open()
                Dim strPer As New SqlCommand("SELECT * FROM P_03 WHERE EFORMSN = '" & streformsn & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then

                    Lab_PWUNIT.Text = RdPer("PWUNIT").ToString
                    Lab_PWTITLE.Text = RdPer("PWTITLE").ToString
                    Lab_PWNAME.Text = RdPer("PWNAME").ToString
                    Lab_PAUNIT.Text = RdPer("PAUNIT").ToString
                    Lab_PANAME.Text = RdPer("PANAME").ToString
                    Lab_PATITLE.Text = RdPer("PATITLE").ToString

                    Lab_nAPPLYTIME.Text = RdPer("nAPPLYTIME")
                    Lab_nPHONE.Text = RdPer("nPHONE").ToString
                    Lab_nREASON.Text = RdPer("nREASON").ToString
                    Lab_nITEM.Text = RdPer("nITEM").ToString
                    Lab_nARRIVEPLACE.Text = RdPer("nARRIVEPLACE").ToString
                    Lab_nARRIVETO.Text = RdPer("nARRIVETO").ToString

                    Lab_nSTARTPOINT.Text = RdPer("nSTARTPOINT").ToString
                    Lab_nENDPOINT.Text = RdPer("nENDPOINT").ToString

                    Lab_nUSEMASTER.Text = RdPer("nUSEMASTER").ToString
                    Lab_nCARMASTER.Text = RdPer("nCARMASTER").ToString

                    Lab_nARRDATE.Text = RdPer("nARRDATE")
                    Lab_nSTHOUR.Text = RdPer("nSTHOUR")
                    Lab_nEDHOUR.Text = RdPer("nEDHOUR")

                    Lab_nSTATUS.Text = RdPer("nSTATUS").ToString
                    Lab_nSTYLE.Text = RdPer("nSTYLE").ToString

                    Lab_nUSEDATE.Text = RdPer("nUSEDATE")
                    Lab_nSTUSEHOUR.Text = RdPer("nSTUSEHOUR")
                    Lab_nSTUSEMIN.Text = RdPer("nSTUSEMIN")
                    Lab_nEDUSEDATE.Text = RdPer("nEDUSEDATE")
                    Lab_nEDUSEHOUR.Text = RdPer("nEDUSEHOUR")
                    Lab_nEDUSEMIN.Text = RdPer("nEDUSEMIN")

                End If
                db.Close()

            Catch ex As Exception

            End Try

        End If

    End Sub
End Class
