'smtpmail - SMTP Mailer for vb.net
'
' Author: Nico Wehmöller
'
' Example usage:
'Dim mail As New smtpMail()
'mail.setCredentials("SMTP_USER", "SMTP_PASSWORT").setSMTPServer("SMTP_HOST", SMTP_PORT)
'Dim content = mail.create(FROM_MAIL,FROM_NAME, TO_MAIL, SUBJECT, MAIL_TEXT, isHTML[true|false])
'mail.send(content)


Imports System.Net.Mail
Public NotInheritable Class smtpMail
    Public Property subject As String
    Public Property body As String
    Public Property receiver As String

    Public Property smtp_user As String
    Public Property smtp_pass As String
    Public Property smtp_server As String
    Public Property smtp_port As String
    Public Property smtp_ssl As Boolean = True


    Public Function setCredentials(ByVal smtp_user As String, ByVal smtp_pass As String)
        Me.smtp_user = smtp_user
        Me.smtp_pass = smtp_pass

        Return Me
    End Function

    Public Function setSMTPServer(ByVal smtp_server As String, ByVal smtp_port As Integer, Optional ByVal smtp_ssl As Boolean = True)
        Me.smtp_server = smtp_server
        Me.smtp_port = smtp_port
        Me.smtp_ssl = smtp_ssl
        Return Me
    End Function

    Public Function create(ByVal mail_from As String, ByVal mail_from_name As String, ByVal mail_receiver As String, ByVal mail_subject As String, ByVal mail_body As String, Optional ByVal mail_isHTML As Boolean = False)
        Dim mail As New MailMessage()
        mail = New MailMessage()
        mail.From = New MailAddress(mail_from, mail_from_name)
        mail.To.Add(mail_receiver)
        mail.Subject = mail_subject
        mail.Body = mail_body
        mail.IsBodyHtml = mail_isHTML
        Return mail

    End Function

    Public Sub send(ByVal MailMessage As MailMessage)
        Try
            Dim smtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            smtpServer.UseDefaultCredentials = False
            smtpServer.Credentials = New Net.NetworkCredential(smtp_user, smtp_pass)
            smtpServer.Port = smtp_port
            smtpServer.EnableSsl = smtp_ssl
            smtpServer.Host = smtp_server
            smtpServer.Timeout = 2
            smtpServer.SendMailAsync(MailMessage)

        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & ex.StackTrace)
        End Try

    End Sub
End Class
