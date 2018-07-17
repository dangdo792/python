Namespace Button
    ''' <summary>
    ''' Observation mail class
    ''' </summary>
    Class ObservationEmail
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
        End Enum

        Private ProjectCoor As String = Nothing

        Sub New(myname_in As Information.CheckNull,
                project_in As Information.CheckNull,
                model_in As Information.CheckNull,
                release_in As Information.CheckNull,
                taskid_in As Information.TaskID,
                modulename_in As Information.CheckNull,
                mo_in As Information.CheckNull,
                pm_in As Information.CheckNull,
                teamlead As Information.CheckNull,
                projectcoor_in As String)
            MyBase.New({myname_in,
                       project_in,
                       model_in,
                       release_in,
                       taskid_in,
                       modulename_in,
                       mo_in,
                       pm_in,
                       teamlead})
            ProjectCoor = projectcoor_in
        End Sub

        Public Overrides Function GetToList()
            Return listinfo(Position.MODULE_OWNER).GetValue
        End Function

        Public Overrides Function GetCCList()
            Return listinfo(Position.PM).GetValue & "; " & listinfo(Position.TEAMLEAD).GetValue & "; " & ProjectCoor
        End Function

        Public Overrides Function GetTitleEmail() As Object
            Return listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & ": Observation"
        End Function

        Public Overrides Function GetEmailBody() As Object
            Dim MailBody As String = Nothing
            MailBody = MailBody & "<font face='Arial' size=2>Hello " & listinfo(Position.MODULE_OWNER).GetValue & ",<BR><BR><BR><BR><BR><BR>"
            Return MailBody
        End Function

    End Class

End Namespace