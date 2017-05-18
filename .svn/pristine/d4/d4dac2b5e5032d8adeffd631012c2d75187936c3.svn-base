Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports System.Net.Mail

Public Class CMail
#Region "Fields"
    Protected _body As String
    Protected _sendFrom As String
    Protected _sendFromWhom As String
    ''' <summary>
    ''' 指定BCC地址；由於可以同時指定多人，所以使用String()型別
    ''' 預設值為service@YourDomain.com
    ''' </summary>
    ''' <remarks></remarks>
    Public bcc As String()
    ''' <summary>
    ''' 定郵件內容的編碼方式，預設值為UTF8
    ''' </summary>
    ''' <remarks></remarks>
    Public bodyEncoding As Encoding
    ''' <summary>
    ''' 指定SMTP Server
    ''' </summary>
    ''' <remarks></remarks>
    Public host As String
    ''' <summary>
    ''' 指定SMTP Server Port；預設值為25
    ''' </summary>
    ''' <remarks></remarks>
    Public hostPort As Integer
    ''' <summary>
    ''' 指定該郵件是否為HTML格式或純文字格式；
    ''' 預設值為False(純文字格式)
    ''' </summary>
    ''' <remarks></remarks>
    Public isBodyHtml As Boolean
    Protected mail As MailMessage
    Private snoo As Integer
    ''' <summary>
    ''' 指定樣式1
    ''' </summary>
    ''' <remarks></remarks>
    Public style1 As String
    ''' <summary>
    ''' 指定整體信件的樣式；預設值是 font-family:Verdana,Arial
    ''' </summary>
    ''' <remarks></remarks>
    Public styleBody As String
    ''' <summary>
    ''' 指定該郵件的標題
    ''' </summary>
    ''' <remarks></remarks>
    Public subject As String
    ''' <summary>
    ''' 指定郵件標題的編碼方式，預設值為 UTF8
    ''' </summary>
    ''' <remarks></remarks>
    Public subjectEncoding As Encoding
    ''' <summary>
    ''' 指定收件人地址；由於可以同時指定多人，所以使用String()型別
    ''' </summary>
    ''' <remarks></remarks>
    Public [to] As List(Of String)

#End Region

#Region "Property"
    Protected ReadOnly Property _bcc As MailAddress
        Get
            If ((Not bcc Is Nothing) AndAlso (bcc.Length > 0)) Then
                Return New MailAddress(Join(bcc, ","))
            End If
            Return Nothing
        End Get
    End Property

    Protected ReadOnly Property _to(ByVal pi As Integer) As MailAddress
        Get
            Return New MailAddress(Me.to.Item(pi))
        End Get
    End Property

    ''' <summary>
    ''' 指定該郵件的內容
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property body As String
        Get
            Return _body
        End Get
        Set(value As String)
            _body = ("<table style=" & styleBody & " border='0'><tr><td>" & value & "</td></tr></table>")
        End Set
    End Property

    Protected ReadOnly Property from As MailAddress
        Get
            Return New MailAddress(sendFrom, _sendFromWhom)
        End Get
    End Property

    ''' <summary>
    ''' 指定寄件人的地址
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property sendFrom As String
        Get
            Return _sendFrom
        End Get
        Set(value As String)
            _sendFrom = value
        End Set
    End Property

    Public Property sendFromWhom As string
        Get
            Return _sendFromWhom
        End Get
        Set(value As String)
            _sendFromWhom = value
        End Set
    End Property

#End Region

#Region "Method"
    Public Sub New()
        _sendFrom = ""
        _sendFromWhom = ""
        host = ""
        hostPort = &H19
        snoo = 0
        [to] = New List(Of String)()
        subject = Nothing
        subjectEncoding = Encoding.UTF8
        _body = Nothing
        bodyEncoding = Encoding.UTF8
        isBodyHtml = True
        styleBody = "font-family:Verdana,Arial; width:80%;font-size:0.8em;"
        style1 = ("background-color:#eeeeee;text-align:center;font-family:Verdana,Arial;font-size:0.9em;" & styleBody)
        mail = New MailMessage

    End Sub

    Public Sub send()
        If (Not Me.to Is Nothing) Then
            mail.From = from
            Dim a As String = mail.From.Address
            Dim u As String = mail.From.User
            Dim sendToCount As Integer = Me.to.Count - 1
            Dim pi As Integer = 0
            Do While (pi <= sendToCount)
                mail.To.Add(_to(pi))
                pi += 1
            Loop
            If Not IsNothing(_bcc) Then
                mail.Bcc.Add(_bcc)
            End If
            mail.IsBodyHtml = isBodyHtml
            mail.Subject = subject
            mail.SubjectEncoding = subjectEncoding
            mail.Body = body
            mail.BodyEncoding = bodyEncoding
            Call New SmtpClient(host, hostPort).Send(mail)
            snoo = 0
        End If
    End Sub
#End Region


End Class