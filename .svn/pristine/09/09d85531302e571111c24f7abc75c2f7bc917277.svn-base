Imports Microsoft.VisualBasic


Public Class C_CheckFun

#Region "[isNumeric] 檢查輸入的String使否為 Int32的數字"
    Public Function isNumeric(ByVal str As String) As Boolean
        Dim NumberStyle As New System.Globalization.NumberStyles
        Dim result As Integer
        Return Int32.TryParse(str, NumberStyle, System.Globalization.CultureInfo.CurrentCulture, result)
    End Function

    Public Function CheckDataInt(ByVal filedvalue As String, ByVal mixInt As Int32, ByVal filed As String, ByVal msg As String) As Boolean
        CheckDataInt = False
        Dim tmpErr = ""
        Dim TransInteger As Integer = 0
        Try
            TransInteger = Integer.Parse(filedvalue)
        Catch ex As Exception
            tmpErr = filed + "您輸入的資料非數字型態資料，請重新確認！"
        End Try
        If tmpErr.Length > 0 Then
            Throw New Exception(tmpErr)
        Else
            If TransInteger < mixInt Then
                tmpErr = filed + "您輸入的數字" + msg
                Throw New Exception(tmpErr)
            End If
        End If
    End Function
#End Region
#Region "[AlertSussTranscation] 產生JavaScript Alert訊息並將視窗導往指定的URL"
    Public Sub AlertSussTranscation(ByVal alertstring As String, ByVal url As String)
        System.Web.HttpContext.Current.Response.Write("<Script language='javascript'>")
        System.Web.HttpContext.Current.Response.Write("alert('" + alertstring + "');")
        System.Web.HttpContext.Current.Response.Write("location.href ='" + url + "';")
        System.Web.HttpContext.Current.Response.Write("</Script>")
        HttpContext.Current.Response.Flush()
        HttpContext.Current.Response.End()
    End Sub
#End Region
#Region "檢查字串長度函中文"
    Public Function GetLength(ByVal str As String) As Integer
        Dim l, t, i, c, d As Int32

        l = Len(str)
        t = l
        For i = 1 To l
            c = Asc(Mid(str, i, 1))
            d = AscW(Mid(str, i, 1))
            If c <> d Then
                '中文漢字，長度再加1
                t = t + 1
            End If
        Next
        GetLength = t
    End Function
    Public Function CheckDataLen(ByVal str As String, ByVal maxlen As Int32, ByVal msg As String, ByVal nstr As Boolean) As String
        CheckDataLen = ""

        Dim Len As Integer = GetLength(str)
        Dim tmpErr = ""

        If str = "" And nstr Then
            tmpErr = msg + "不可空白"
            Throw New Exception(tmpErr)
        End If

        If Len > maxlen Then
            tmpErr = msg + "長度(" + Len.ToString + ")過長，不可超過" + maxlen.ToString
            Throw New Exception(tmpErr)
        End If
        CheckDataLen = tmpErr
    End Function
#End Region
#Region "計算日期"
    Public Function BetweenDay(ByVal sdt As Date, ByVal edt As Date) As Integer
        '傳回兩日期之起迄天數含頭尾
        Dim d1, d2, day As Int32
        d1 = (sdt.Year - 1) * 365 + sdt.DayOfYear
        d2 = (edt.Year - 1) * 365 + edt.DayOfYear
        day = d1 - d2
        If day < 0 Then
            day *= -1
        End If
        BetweenDay = day + 1
    End Function
#End Region

End Class
