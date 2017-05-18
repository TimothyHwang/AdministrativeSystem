Imports System.Data.SqlClient
Imports System.Data

Partial Class M_Source_08_ConfidentialPhotoCopy
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim do_sql As New C_SQLFUN
        Dim connstr As String = do_sql.G_conn_string
        Dim dt As DataTable = New DataTable()
        Dim id As String = Request("id").ToString()
        Try
            Using conn As New SqlConnection(connstr)
                Dim da As New SqlDataAdapter("select * from P_0804 where Guid_ID = " + id, conn)
                da.Fill(dt)

                If dt.Rows.Count > 0 Then
                    'Top1unitName 申請單位全銜
                    lbTop1unitName1.Text = IIf(dt.Rows(0).Item(1) Is Nothing, "", dt.Rows(0).Item(1).ToString())
                    lbTop1unitName2.Text = lbTop1unitName1.Text
                    'Subject 主旨(頓號)簡由或(資料名稱)
                    lbSubject1.Text = IIf(dt.Rows(0).Item(2) Is Nothing, "", dt.Rows(0).Item(2).ToString())
                    lbSubject2.Text = lbSubject1.Text
                    'SignDateTime 發文時間
                    lbSignDateTime1.Text = IIf(dt.Rows(0).Item(3) Is Nothing, "", Convert.ToDateTime(dt.Rows(0).Item(3)).ToString("yyyy/MM/dd"))
                    lbSignDateTime2.Text = lbSignDateTime1.Text
                    'Security_No 發文字號
                    lbSecurity_No1.Text = IIf(dt.Rows(0).Item(4) Is Nothing, "", dt.Rows(0).Item(4).ToString())
                    lbSecurity_No2.Text = lbSecurity_No1.Text
                    'Security_Level 機密等級 (2:密 3:機密 4:極機密 5:絕對機密)
                    Dim strSecurityLevelNo As String = IIf(dt.Rows(0).Item(5) Is Nothing, "", dt.Rows(0).Item(5).ToString())
                    Select Case strSecurityLevelNo
                        Case "2"
                            lbSecurity_Level1.Text = "密"
                        Case "3"
                            lbSecurity_Level1.Text = "機密"
                        Case "4"
                            lbSecurity_Level1.Text = "極機密"
                        Case "5"
                            lbSecurity_Level1.Text = "絕對機密"
                        Case Else
                            lbSecurity_Level1.Text = ""
                    End Select
                    lbSecurity_Level2.Text = lbSecurity_Level1.Text
                    'Security_Type 機密屬性 (0:一般公務機密 1:國家機密 2:軍事機密 3:國防秘密 4:國家機密亦屬軍事機密 5:國家機密亦屬國防秘密)
                    Dim strSecurityTypeNo As String = IIf(dt.Rows(0).Item(6) Is Nothing, "", dt.Rows(0).Item(6).ToString())
                    Select Case strSecurityTypeNo
                        Case "0"
                            lbSecurity_Type1.Text = "一般公務機密"
                        Case "1"
                            lbSecurity_Type1.Text = "國家機密"
                        Case "2"
                            lbSecurity_Type1.Text = "軍事機密"
                        Case "3"
                            lbSecurity_Type1.Text = "國防秘密"
                        Case "4"
                            lbSecurity_Type1.Text = "國家機密亦屬軍事機密"
                        Case "5"
                            lbSecurity_Type1.Text = "國家機密亦屬國防秘密"
                        Case Else
                            lbSecurity_Type1.Text = ""
                    End Select
                    lbSecurity_Type2.Text = lbSecurity_Type1.Text
                    'Security_Range 保密期限/保密條件
                    lbSecurity_Range1.Text = IIf(dt.Rows(0).Item(7) Is Nothing, "", dt.Rows(0).Item(7).ToString())
                    lbSecurity_Range2.Text = lbSecurity_Range1.Text
                    'ProduceUnit 產製單位
                    lbProduceUnit1.Text = IIf(dt.Rows(0).Item(22) Is Nothing, "", dt.Rows(0).Item(22).ToString())
                    lbProduceUnit2.Text = lbProduceUnit1.Text
                    'AgreeTimeOrNumber 同意複(影)印時間/文號
                    lbAgreeTimeOrNumber1.Text = IIf(dt.Rows(0).Item(23) Is Nothing, "", dt.Rows(0).Item(23).ToString())
                    lbAgreeTimeOrNumber2.Text = lbAgreeTimeOrNumber1.Text
                    'AgreeSuperior 同意複(影)印權責長官級職姓名
                    lbAgreeSuperior1.Text = IIf(dt.Rows(0).Item(24) Is Nothing, "", dt.Rows(0).Item(24).ToString())
                    lbAgreeSuperior2.Text = lbAgreeSuperior1.Text
                    'Purpose 用途 (1:呈閱 2:分會、辦 3:作業用 4:歸檔 5:隨文分發 6:會議分發 7:(其他)Purpose_Other)
                    Dim strPurposeNo As String = IIf(dt.Rows(0).Item(8) Is Nothing, "", dt.Rows(0).Item(8).ToString())
                    If Not String.IsNullOrEmpty(strPurposeNo) Then
                        rblPurpose1.SelectedValue = strPurposeNo
                        rblPurpose2.SelectedValue = strPurposeNo
                        If strPurposeNo.Equals("7") Then
                            rbPurposeElse1.Checked = True
                            rbPurposeElse2.Checked = True
                            lbPurpose_Other1.Text = IIf(dt.Rows(0).Item(9) Is Nothing, "", dt.Rows(0).Item(9).ToString())
                            lbPurpose_Other2.Text = lbPurpose_Other1.Text
                        End If
                    End If
                    'Printer_Datetime 複印時間
                    lbPrinter_Datetime1.Text = IIf(dt.Rows(0).Item(17) Is Nothing, "", Convert.ToDateTime(dt.Rows(0).Item(17)).ToString("yyyy/MM/dd"))
                    lbPrinter_Datetime2.Text = lbPrinter_Datetime1.Text
                    'Printer_Num 複(影)印機浮水印暗記編號
                    lbPrinter_Num1.Text = IIf(dt.Rows(0).Item(10) Is Nothing, "", dt.Rows(0).Item(10).ToString())
                    lbPrinter_Num2.Text = lbPrinter_Num1.Text
                    'Ori_sheet 原件張數
                    lbOri_sheet1.Text = IIf(dt.Rows(0).Item(18) Is Nothing, "", dt.Rows(0).Item(18).ToString())
                    lbOri_sheet2.Text = lbOri_sheet1.Text
                    'Copy_sheet 每份複印張數
                    lbCopy_sheet1.Text = IIf(dt.Rows(0).Item(19) Is Nothing, "", dt.Rows(0).Item(19).ToString())
                    lbCopy_sheet2.Text = lbCopy_sheet1.Text
                    'Total_sheet 合計複印張數
                    lbTotal_sheet1.Text = IIf(dt.Rows(0).Item(20) Is Nothing, "", dt.Rows(0).Item(20).ToString())
                    lbTotal_sheet2.Text = lbTotal_sheet1.Text
                    'Sheet_ID 複(影)印張數流水號
                    lbSheet_ID1.Text = IIf(dt.Rows(0).Item(21) Is Nothing, "", dt.Rows(0).Item(21).ToString())
                    lbSheet_ID2.Text = lbSheet_ID1.Text
                    'Memo 附註
                    lbMemo1.Text = IIf(dt.Rows(0).Item(15) Is Nothing, "", dt.Rows(0).Item(15).ToString())
                    lbMemo2.Text = lbMemo1.Text
                End If
            End Using
        Catch ex As Exception
            dt = Nothing
        End Try
    End Sub
    
End Class
