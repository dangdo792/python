Namespace Button
    ''' <summary>
    ''' Email for request reviewer assignment
    ''' </summary>
    Class ReviewerAssignEmail
        Inherits EmailBase

        Private Enum Position
            MYNAME
            PROJECT
            MODEL
            RELEASE
            TASK_ID
            MODULE_NAME
            TEAMLEAD
            ELOC
        End Enum

        Public Sub New(myname_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       model_in As Information.CheckNull,
                       release_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       teamlead As Information.CheckNull,
                       eloc_in As Information.CheckNull)
            MyBase.New({myname_in,
                       project_in,
                       model_in,
                       release_in,
                       taskid_in,
                       modulename_in,
                       teamlead,
                       eloc_in})
        End Sub

        Public Overrides Function GetToList() As Object
            Return listinfo(Position.TEAMLEAD).GetValue
        End Function

        Public Overrides Function GetTitleEmail() As Object
            Dim title As String = Nothing
            title = listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & ":"
            title = title & listinfo(Position.ELOC).GetValue & "_1: FullPhaseReview_Request"
            Return title
        End Function

        Public Overrides Function GetEmailBody() As Object
            Dim MailBody As String = Nothing
            MailBody = MailBody & "<font face='Arial' size=2>Hello leader,<BR><BR>"
            MailBody = MailBody & "Kindly help me to assign reviewer for this task.<BR><BR><BR>"
            Return MailBody
        End Function

    End Class

End Namespace
