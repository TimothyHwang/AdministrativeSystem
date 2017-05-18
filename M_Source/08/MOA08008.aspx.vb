Imports System.Drawing.Printing
Imports System.Data.SqlClient
Imports System.Data
Imports Microsoft.Win32

Partial Class M_Source_08_MOA08008
    Inherits System.Web.UI.Page
    Dim CP As New C_Public
    Dim aa As C_CheckFun
    Dim user_id, org_uid As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '將頁首頁尾取消
        Dim pageKey As RegistryKey = Registry.CurrentUser.OpenSubKey("software\microsoft\internet explorer\pagesetup", True)
        If Not pageKey Is Nothing Then
            Dim newHeader As String = ""
            Dim curHeader As Object = pageKey.GetValue("header")
            pageKey.SetValue("headerTemp", curHeader)
            pageKey.SetValue("header", newHeader)
            Dim newFooter As String = ""
            Dim curFooter As Object = pageKey.GetValue("footer")
            pageKey.SetValue("footerTemp", curFooter)
            pageKey.SetValue("footer", newFooter)
            pageKey.Close()
        End If

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else

            '判斷使用者是否仍有未送審或審核中之申請單
            Dim sql_function As New C_SQLFUN
            Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

            command.CommandType = Data.CommandType.Text
            command.CommandText = "SELECT * FROM P_0804 WITH(NOLOCK) WHERE Print_UserID = @Print_UserID AND Security_Status = 0"
            command.Parameters.Add(New SqlParameter("Print_UserID", SqlDbType.NVarChar, 10)).Value = user_id.Trim()
            Try
                command.Connection.Open()
                Dim adapter As New SqlDataAdapter
                Dim ds As New DataSet
                adapter.SelectCommand = command
                adapter.Fill(ds)
                adapter.Dispose()
                command.Dispose()
                If ds.Tables(0).Rows.Count > 0 Then
                    Response.Write("<script type='text/javascript'>alert('您仍有未送審之機密資訊複影印資料申請單，故無法再申請第二張。');window.location.href='MOA08007.aspx';</script>")
                End If
            Catch ex As Exception
                ClientScript.RegisterStartupScript(Me.GetType(), "SqlErrorMsg2", "alert('" + ex.Message + "');", True)
            Finally
                If command.Connection.State.Equals(ConnectionState.Open) Then
                    command.Connection.Close()
                End If
            End Try

            '判斷登入者權限
            If CP.LoginCheck(user_id, "MOA08007") <> "" Then
                CP.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA08007.aspx")
                Response.End()
            End If

            '找出登入者的一級單位
            Dim sAllName As String = CP.GetNo1ORGName(org_uid)
            If sAllName <> "0" Then
                lb_No1ORGName.Text = sAllName
            Else
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('取得一級單位全銜失敗，請重新操作或連絡資訊人員！');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            End If
        End If

        If (Not IsPostBack) Then
            lbSecurity_Level_2.Visible = True
            lbSecurity_Level_2.Text = "一般公務機密"
            '取得保密期限/保密條件 的 新/舊資料
            If (Not ViewState("Security_Desc") Is Nothing) Then
                tb_conditions.Text = ViewState("Security_Desc")
            End If

            Dim iYear As Integer = DateTime.Now.Year
            'Dim iBFYear As Integer = 2010
            Dim iAFYear As Integer = iYear + 30
            For i As Integer = iYear To iAFYear
                ddl_Year.Items.Add(i)
            Next
            ddl_Year.SelectedValue = iYear
            If (Not ViewState("Chose_Year") Is Nothing) Then
                ddl_Year.SelectedValue = ViewState("Chose_Year")
            End If
            Dim iMonth As Integer = DateTime.Now.Month
            For i As Integer = 1 To 12
                ddl_Month.Items.Add(i)
            Next
            ddl_Month.SelectedValue = iMonth
            If (Not ViewState("Chose_Month") Is Nothing) Then
                ddl_Month.SelectedValue = ViewState("Chose_Month")
            End If
            Dim iDay As Integer = Date.DaysInMonth(iYear, iMonth)
            For k As Integer = 1 To iDay
                ddl_Day.Items.Add(k)
            Next
            ddl_Day.SelectedValue = DateTime.Now.Day
            If (Not ViewState("Chose_Day") Is Nothing) Then
                ddl_Day.SelectedValue = ViewState("Chose_Day")
            End If

            If (Not ViewState("Chose_Change") Is Nothing) Then
                If (ViewState("Chose_Change") = "0") Then
                    RB11.Checked = True
                Else
                    RB12.Checked = True
                End If
            End If

            If (Not ViewState("ChoseSecurity") Is Nothing) Then
                If (ViewState("ChoseSecurity") = "0") Then
                    RBSecurity2_CheckedChanged(Nothing, Nothing)
                Else
                    RBSecurity1_CheckedChanged(Nothing, Nothing)
                End If
            End If

            Dim aySecurityLevel As String() = {"密", "機密", "極機密", "絕對機密"}
            For l As Integer = 0 To aySecurityLevel.Length - 1
                Dim item As New ListItem(aySecurityLevel(l), l + 2, True)
                ddl_Security_Level2.Items.Add(item)
            Next
            Dim Security_Level As String = ddl_Security_Level.SelectedValue
            Dim SecurityLevel2 As String = String.Empty
            For Each item As ListItem In ddl_Security_Level2.Items
                If (Security_Level = item.Value) Then
                    SecurityLevel2 = item.Value
                End If
            Next
            If (SecurityLevel2 <> String.Empty) Then
                ddl_Security_Level2.Items.Remove(ddl_Security_Level2.Items.FindByValue(SecurityLevel2))
            End If

        End If
    End Sub

    Protected Sub ImgSavaPrint_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgSavaPrint.Click
        '檢查輸入欄位正確與否
        Dim sSubject As String = tb_Subject.Text.Trim()
        If (sSubject.Length < 1) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg1", "alert('主旨、簡由請勿空白 !');", True)
            Return
        End If
        If (String.IsNullOrEmpty(lb_SignDT.Text)) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg1", "alert('請選擇發文時間 !');", True)
            Return
        End If
        Dim Security_No As String = tb_Subject.Text.Trim()
        If (tb_Security_No1.Text.Trim().Length < 1 Or tb_Security_No2.Text.Trim().Length < 1) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg2", "alert('發文字號不能有空白 !');", True)
            Return
        Else
            Security_No = tb_Security_No1.Text.Trim() + "字第" + tb_Security_No2.Text.Trim + "號"
        End If

        If (String.IsNullOrEmpty(lb_Security_Range.Text)) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg1", "alert('請選擇保密期限/保密條件 !');", True)
            Return
        End If

        Dim ichosePurpose As Integer = 0
        For i As Integer = 1 To 7
            Dim RB As RadioButton = CType(Me.FindControl("RB" + i.ToString()), RadioButton)
            If (RB.Checked) Then
                ichosePurpose = i
            End If
        Next
        If (ichosePurpose = 0) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg3", "alert('請選擇用途 !');", True)
            Return
        End If


        If (RB7.Checked) Then
            If (tb_Purpose_Other.Text.Trim().Length < 1) Then
                ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg4", "alert('用途點選[其他]時，請輸入其他用途說明 !');", True)
                Return
            End If
        End If

        If lb_Print_DT.Text.Trim().Length <= 0 Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg5", "alert('請輸入複印時間!');", True)
            Return
        End If

        If lb_Printer_Num.Text.Trim().Length <= 0 Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg6", "alert('請輸入複(影)印機浮水印暗記編號!');", True)
            Return
        End If
        If lb_Org_Num.Text.Trim().Length <= 0 Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg7", "alert('請輸入原件張數!');", True)
            Return
        ElseIf Not IsNumeric(lb_Org_Num.Text.Trim()) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg7", "alert('請輸入正確原件張數(數字)!');", True)
            Return
        End If

        If lb_PrintSet_Cnt.Text.Trim().Length <= 0 Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg8", "alert('請輸入每張複印張數!');", True)
            Return
        ElseIf Not IsNumeric(lb_PrintSet_Cnt.Text.Trim()) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg8", "alert('請輸入正確每張複印張數(數字)!');", True)
            Return
        End If

        If lb_PrintTotal.Text.Trim().Length <= 0 Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg9", "alert('請輸入每張合計複印張數!');", True)
            Return
        ElseIf Not IsNumeric(lb_PrintTotal.Text.Trim()) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "DataErrorMsg9", "alert('請輸入正確每張合計複印張數(數字)!');", True)
            Return
        End If


        Dim bl_InsertResult As Boolean = False
        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        command.CommandType = Data.CommandType.Text

        command.CommandText = "Insert Into P_0804(Top1unitName, Subject, SignDateTime, Security_No, Security_Level, " & _
                            "Security_Type,Security_Range,Purpose,Purpose_Other,Print_UserID,Memo,PAIDNO,Printer_Num, " & _
                            "Printer_Datetime,Ori_sheet,Copy_sheet,Total_sheet,Sheet_ID,ProduceUnit,AgreeTimeOrNumber,AgreeSuperior)" & _
                            "Values(@Top1unitName, @Subject, @SignDateTime, @Security_No, @Security_Level,@Security_Type," & _
                            "@Security_Range,@Purpose,@Purpose_Other,@Print_UserID,@Memo,@PAIDNO,@Printer_Num,@Printer_Datetime," & _
                            "@Ori_sheet,@Copy_sheet,@Total_sheet,@Sheet_ID,@ProduceUnit,@AgreeTimeOrNumber,@AgreeSuperior)"

        command.Parameters.Add(New SqlParameter("Top1unitName", SqlDbType.NVarChar, 50)).Value = lb_No1ORGName.Text.Trim()
        command.Parameters.Add(New SqlParameter("Subject", SqlDbType.NVarChar, 255)).Value = tb_Subject.Text.Trim()
        If (lb_SignDT.Text.Trim() = "") Then
            command.Parameters.Add(New SqlParameter("SignDateTime", SqlDbType.SmallDateTime)).Value = DBNull.Value
        Else
            command.Parameters.Add(New SqlParameter("SignDateTime", SqlDbType.SmallDateTime)).Value = DateTime.Parse(lb_SignDT.Text.Trim())
        End If

        command.Parameters.Add(New SqlParameter("Security_No", SqlDbType.NVarChar, 20)).Value = tb_Security_No1.Text.Trim() + "字第" + tb_Security_No2.Text.Trim + "號"
        command.Parameters.Add(New SqlParameter("Security_Level", SqlDbType.TinyInt)).Value = ddl_Security_Level.SelectedValue
        command.Parameters.Add(New SqlParameter("Security_Type", SqlDbType.TinyInt)).Value = IIf(ddl_Security_Type.Visible, ddl_Security_Type.SelectedValue, 0)
        command.Parameters.Add(New SqlParameter("Security_Range", SqlDbType.NVarChar, 100)).Value = lb_Security_Range.Text.Trim()
        command.Parameters.Add(New SqlParameter("Purpose", SqlDbType.TinyInt)).Value = ichosePurpose
        command.Parameters.Add(New SqlParameter("Purpose_Other", SqlDbType.NVarChar, 50)).Value = tb_Purpose_Other.Text.Trim()
        command.Parameters.Add(New SqlParameter("Print_UserID", SqlDbType.VarChar, 10)).Value = user_id
        command.Parameters.Add(New SqlParameter("Memo", SqlDbType.NVarChar, 255)).Value = tb_Memo.Value.Trim()
        command.Parameters.Add(New SqlParameter("PAIDNO", SqlDbType.VarChar, 10)).Value = user_id.Trim()

        command.Parameters.Add(New SqlParameter("Printer_Num", SqlDbType.VarChar, 6)).Value = lb_Printer_Num.Text.Trim()
        command.Parameters.Add(New SqlParameter("Printer_Datetime", SqlDbType.SmallDateTime)).Value = DateTime.Parse(lb_Print_DT.Text.Trim())
        command.Parameters.Add(New SqlParameter("Ori_sheet", SqlDbType.Int)).Value = lb_Org_Num.Text.Trim()
        command.Parameters.Add(New SqlParameter("Copy_sheet", SqlDbType.Int)).Value = lb_PrintSet_Cnt.Text.Trim()
        command.Parameters.Add(New SqlParameter("Total_sheet", SqlDbType.Int)).Value = lb_PrintTotal.Text.Trim()
        command.Parameters.Add(New SqlParameter("Sheet_ID", SqlDbType.VarChar, 30)).Value = lb_sheetID.Text.Trim()
        command.Parameters.Add(New SqlParameter("ProduceUnit", SqlDbType.NVarChar, 50)).Value = txtProduceUnit.Text.Trim()
        command.Parameters.Add(New SqlParameter("AgreeTimeOrNumber", SqlDbType.NVarChar, 50)).Value = txtAgreeTimeOrNumber.Text.Trim()
        command.Parameters.Add(New SqlParameter("AgreeSuperior", SqlDbType.NVarChar, 50)).Value = txtAgreeSuperior.Text.Trim()

        Try
            command.Connection.Open()
            command.ExecuteNonQuery()
            bl_InsertResult = True
            ddl_Security_Level.Visible = False
            lb_Security_Level.Visible = True
            ddl_Security_Type.Visible = False
            lbSecurity_Level_2.Visible = True
            ImgSavaPrint.Visible = False
            'lb_Print_DT.Text = ""          '複印時間
            'lb_Printer_Num.Text = ""       '複(影)印機浮水印暗記編號
            'lb_Org_Num.Text = ""           '原件張數
            'lb_PrintSet_Cnt.Text = ""      '每份複印張數
            'lb_PrintTotal.Text = ""        '合計複印張數
            'lb_sheetID.Text = ""           '複(影)印張數流水號
            'lb_Security_superior.Text = "" '保密軍官副署
            ImgDate1.Visible = False

        Catch ex As Exception
            If (TypeOf ex Is SqlException) Then
                If (CType(ex, SqlException).Number.Equals(2627)) Then
                    ClientScript.RegisterStartupScript(Me.GetType(), "SqlErrorMsg1", "alert('該筆資料已存在 !');", True)
                End If
            End If
            ClientScript.RegisterStartupScript(Me.GetType(), "SqlErrorMsg2", "alert('" + ex.Message + "');", True)
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
        End Try
        If (bl_InsertResult) Then
            'Dim js As String = "window.print();"
            'ClientScript.RegisterClientScriptBlock(Me.GetType(), "myscript", js, True)
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "myscript", "alert('資料新增成功');window.location.href='MOA08007.aspx';", True)
        End If
    End Sub

    Protected Sub ddl_Security_Level_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_Security_Level.SelectedIndexChanged
        If ddl_Security_Level.SelectedValue <> "2" Then
            ddl_Security_Type.Visible = True
            lbSecurity_Level_2.Visible = False
        Else
            lbSecurity_Level_2.Visible = True
            lbSecurity_Level_2.Text = "一般公務機密"
            ddl_Security_Type.Visible = False
        End If

        '機密等級已選的等級，於變更機密下拉式選單中不應出現
        ddl_Security_Level2.Items.Clear()
        Dim aySecurityLevel As String() = {"密", "機密", "極機密", "絕對機密"}
        For l As Integer = 0 To aySecurityLevel.Length - 1
            Dim item As New ListItem(aySecurityLevel(l), l + 2, True)
            ddl_Security_Level2.Items.Add(item)
        Next
        Dim Security_Level As String = ddl_Security_Level.SelectedValue
        Dim SecurityLevel2 As String = String.Empty
        For Each item As ListItem In ddl_Security_Level2.Items
            If (Security_Level = item.Value) Then
                SecurityLevel2 = item.Value
            End If
        Next
        If (SecurityLevel2 <> String.Empty) Then
            ddl_Security_Level2.Items.Remove(ddl_Security_Level2.Items.FindByValue(SecurityLevel2))
        End If
        lb_Security_Level.Text = ddl_Security_Level.SelectedItem.ToString()
        '若原本舊選擇的變更機密等級已於前面所選的一樣即不做選取動作，因為在下拉式選單中已不存在
        If (Not ViewState("Chose_ChangeSecurity") Is Nothing And ViewState("Chose_ChangeSecurity") <> Security_Level) Then
            ddl_Security_Level2.SelectedValue = ViewState("Chose_ChangeSecurity")
        End If
    End Sub

    Protected Sub ImgDate1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Div_grid.Visible = True
        Div_grid.Style("Top") = hidImgDate1Top.Value + "px"
        Div_grid.Style("left") = hidImgDate1Left.Value + "px"
        lb_SignDT.Text = DateTime.Now.ToString("yyyy/MM/dd")
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        lb_SignDT.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False
    End Sub
    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        lb_Print_DT.Text = Calendar2.SelectedDate.Date
        Div1.Visible = False
    End Sub
    Protected Sub btnClose1_Click(sender As Object, e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub lnkbtn_Security_Range_Click(sender As Object, e As System.EventArgs) Handles lnkbtn_Security_Range.Click
        Div_Security.Style("left") = hidSecurity_RangeLeft.Value + "px"
        Div_Security.Style("Top") = hidSecurity_RangeTop.Value + "px"
        Div_Security.Visible = True
    End Sub

    Protected Sub ddl_Month_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_Month.SelectedIndexChanged
        Dim iYear As Integer = Integer.Parse(ddl_Year.SelectedValue)
        Dim iMonth As Integer = Integer.Parse(ddl_Month.SelectedValue)
        Dim iDay As Integer = Date.DaysInMonth(iYear, iMonth)
        For k As Integer = 1 To iDay
            ddl_Day.Items.Add(k)
        Next
    End Sub

    Protected Sub Security_OK_Click(sender As Object, e As System.EventArgs) Handles Security_OK.Click
        Dim Security_Desc As String = String.Empty
        If (tb_conditions.Text.Trim() <> "") Then
            Security_Desc = tb_conditions.Text.Trim() + ","
            ViewState("Security_Desc") = Security_Desc
        End If
        If (RBSecurity1.Checked) Then
            ViewState("ChoseSecurity") = "1"
            ViewState("Chose_Year") = ddl_Year.SelectedValue
            ViewState("Chose_Month") = ddl_Month.SelectedValue
            ViewState("Chose_Day") = ddl_Day.SelectedValue
            If (RB11.Checked) Then
                ViewState("Chose_Change") = "0"
                Security_Desc += "本文件保密至" + ddl_Year.SelectedValue + "年" + ddl_Month.SelectedValue + "月" + ddl_Day.SelectedValue + "日解除密等"
            Else
                ViewState("Chose_Change") = "1"
                Security_Desc += "本文件保密至" + ddl_Year.SelectedValue + "年" + ddl_Month.SelectedValue + "月" + ddl_Day.SelectedValue + "日解除為：" + ddl_Security_Level2.SelectedItem.ToString()
            End If
        Else
            ViewState("ChoseSecurity") = "0" 'default Checked
            Security_Desc += "本文件永久保密"
        End If
        lb_Security_Range.Text = Security_Desc
        Div_Security.Visible = False
    End Sub

    Protected Sub Security_Close_Click(sender As Object, e As System.EventArgs) Handles Security_Close.Click
        Div_Security.Visible = False
    End Sub

    Protected Sub RBSecurity2_CheckedChanged(sender As Object, e As System.EventArgs) Handles RBSecurity2.CheckedChanged
        If (RBSecurity2.Checked) Then
            ddl_Year.Enabled = False
            ddl_Month.Enabled = False
            ddl_Day.Enabled = False
            RB11.Enabled = False
            RB12.Enabled = False
            ddl_Security_Level2.Enabled = False
        Else
            ddl_Year.Enabled = True
            ddl_Month.Enabled = True
            ddl_Day.Enabled = True
            RB11.Enabled = True
            RB12.Enabled = True
            ddl_Security_Level2.Enabled = True
        End If
    End Sub

    Protected Sub RBSecurity1_CheckedChanged(sender As Object, e As System.EventArgs) Handles RBSecurity1.CheckedChanged
        If (RBSecurity1.Checked) Then
            ddl_Year.Enabled = True
            ddl_Month.Enabled = True
            ddl_Day.Enabled = True
            RB11.Enabled = True
            RB12.Enabled = True
            ddl_Security_Level2.Enabled = True
        Else
            ddl_Year.Enabled = False
            ddl_Month.Enabled = False
            ddl_Day.Enabled = False
            RB11.Enabled = False
            RB12.Enabled = False
            ddl_Security_Level2.Enabled = False
        End If
    End Sub

    Protected Sub RB7_CheckedChanged(sender As Object, e As System.EventArgs) Handles RB7.CheckedChanged, RB1.CheckedChanged, RB2.CheckedChanged, RB3.CheckedChanged, RB4.CheckedChanged, RB5.CheckedChanged, RB6.CheckedChanged
        If (RB7.Checked) Then
            tb_Purpose_Other.Enabled = True
        Else
            tb_Purpose_Other.Text = ""
            tb_Purpose_Other.Enabled = False
        End If
    End Sub

    Protected Sub ddl_Security_Level2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_Security_Level2.SelectedIndexChanged
        ViewState("Chose_ChangeSecurity") = ddl_Security_Level2.SelectedValue
    End Sub

    Protected Sub ddl_Security_Type_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_Security_Type.SelectedIndexChanged
        lbSecurity_Level_2.Text = ddl_Security_Type.SelectedItem.ToString()
    End Sub

    Protected Sub ImgDate2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
        Div1.Visible = True
        Div1.Style("Top") = hidImgDate2Top.Value + "px"
        Div1.Style("left") = hidImgDate2Left.Value + "px"
        lb_Print_DT.Text = DateTime.Now.ToString("yyyy/MM/dd")
    End Sub

    Protected Sub Calendar2_DayRender(sender As Object, e As System.Web.UI.WebControls.DayRenderEventArgs) Handles Calendar2.DayRender
        If e.Day.Date < System.DateTime.Now.Date Then
            e.Day.IsSelectable = False
            e.Cell.Font.Strikeout = True
        End If
    End Sub

    Private Sub SqlDataReader(p1 As Object)
        Throw New NotImplementedException
    End Sub

End Class
