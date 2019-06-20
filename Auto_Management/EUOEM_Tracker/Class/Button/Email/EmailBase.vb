Namespace Button
    ''' <summary>
    ''' Abstract base class for Email button
    ''' </summary>
    MustInherit Class EmailBase
        Inherits ButtonBase

        Private Enum Position
            MYNAME
            PROJECT
            MODEL
            RELEASE
        End Enum

        Public Sub New(listinfo_in() As Information.InfoBase)
            MyBase.New(listinfo_in)
        End Sub

        Public Overrides Function DoFunctionality() As String
            Dim oolApp = CreateObject("Outlook.Application")
            Dim Email = oolApp.CreateItem(0)

            ' Fill TO of email
            Email.to = GetToList()
            ' Fill CC of email
            Email.CC = GetCCList()
            ' Fill Subject of email
            Email.Subject = "[" & listinfo(Position.PROJECT).GetValue & "][" & listinfo(Position.MODEL).GetValue & "_" & listinfo(Position.RELEASE).GetValue & "] " & GetTitleEmail()
            ' Fill body of email
            Email.HTMLBody = GetEmailBody()
            ' Fill footnote of email
            Email.HTMLBody = Email.HTMLBody & "Thank you.<BR><BR>"
            Email.HTMLBody = Email.HTMLBody & "Best regards,<BR><BR>"
            Email.HTMLBody = Email.HTMLBody & "<B>" + listinfo(Position.MYNAME).GetValue + " (Mr.)<BR>"
            Email.HTMLBody = Email.HTMLBody & "RBVH/ESS1</B><BR>"
            Email.HTMLBody = Email.HTMLBody & "</BODY></HTML>"

            ' Show the email
            Email.display()

            Return Nothing
        End Function

        ''' <summary>
        ''' Get TO information for email
        ''' </summary>
        ''' <returns>String contain TO information</returns>
        MustOverride Function GetToList()

        ''' <summary>
        ''' Get CC information for email
        ''' </summary>
        ''' <returns>String contain CC information, default is nothing</returns>
        Public Overridable Function GetCCList()
            Return Nothing
        End Function

        ''' <summary>
        ''' Get title information for email
        ''' </summary>
        ''' <returns>String contain title information</returns>
        MustOverride Function GetTitleEmail()

        ''' <summary>
        ''' Get body information for email
        ''' </summary>
        ''' <returns>HTML content for body information</returns>
        MustOverride Function GetEmailBody()
    End Class
End Namespace
