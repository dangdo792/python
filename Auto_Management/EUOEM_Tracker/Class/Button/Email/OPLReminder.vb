Namespace Button
    ''' <summary>
    ''' OPL remind email
    ''' </summary>
    Class OPLReminderEmail
        Inherits EmailBase

        Private Enum Position
            MYNAME
            PROJECT
            MODEL
            RELEASE
            TASK_ID
            MODULE_NAME
            MODULE_OWNER
            PM
            TEAMLEAD
            OPL
        End Enum

        Private ProjectCoor As String

        Sub New(myname_in As Information.CheckNull,
               project_in As Information.CheckNull,
                model_in As Information.CheckNull,
                release_in As Information.CheckNull,
                taskid_in As Information.TaskID,
                modulename_in As Information.CheckNull,
                mo_in As Information.CheckNull,
                pm_in As Information.CheckNull,
                teamlead_in As Information.CheckNull,
                opl_in As Information.CheckNull,
                ProjectCoor_in As String)
            MyBase.New({myname_in,
                       project_in,
                       model_in,
                       release_in,
                       taskid_in,
                       modulename_in,
                       mo_in,
                       pm_in,
                       teamlead_in,
                       opl_in})
            ProjectCoor = ProjectCoor_in
        End Sub

        Public Overrides Function GetToList()
            Return listinfo(Position.MODULE_OWNER).GetValue
        End Function

        Public Overrides Function GetCCList()
            Return listinfo(Position.PM).GetValue & "; " & listinfo(Position.TEAMLEAD).GetValue & "; " & ProjectCoor
        End Function

        Public Overrides Function GetTitleEmail() As Object
            Return "OPL clarifications in " & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & ": A Gentle Reminder"
        End Function

        Public Overrides Function GetEmailBody() As Object
            Dim MailBody As String = Nothing

            ' Greeting to MO
            Dim MOName() As String = Split(listinfo(Position.MODULE_OWNER).GetValue, " ")
            If MOName.Length > 1 Then
                MailBody = MailBody & "<font face='Arial' size=2>Hello " & MOName(1) & ",<BR><BR>"
            Else
                MailBody = MailBody & "<font face='Arial' size=2>Hello " & MOName(0) & ",<BR><BR>"
            End If

            ' Fill main body
            MailBody = MailBody & "&nbsp;&nbsp;&nbsp;&nbsp;With Respect to the Email on OPL clarification dated." _
                & " I request you to clarify the open points for the " + "<B>" + listinfo(Position.MODULE_NAME).GetValue + "</B>" + "<BR><BR>"
            MailBody = MailBody & "Link: " + " <a href=" + listinfo(Position.OPL).GetValue + ">" + listinfo(Position.OPL).GetValue + "</a><BR><BR>"
            MailBody = MailBody & "Please clarify the OPLs to enable us to deliver the task on agreed delivery date " + "<B>" _
                + DateTime.Now.AddDays(2).ToString("dd-MM-yyyy") + "</B>." + "<BR><BR>"
            MailBody = MailBody & "Awaiting for your response,<BR><BR>"
            Return MailBody
        End Function

    End Class

End Namespace
