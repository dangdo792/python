Namespace Button
    ''' <summary>
    ''' Request MKS checkin email
    ''' </summary>
    Class MKSCheckinEmail
        Inherits EmailBase

        Private SysType As String

        Private Enum Position
            MYNAME
            PROJECT
            MODEL
            RELEASE
            TASK_ID
            MODULE_NAME
            MKS_CHECKER
            SUB_REV
            MODULE_PATH
            REVISION
            RESULT_PATH
            STATEMENT
            DECISION
            DEFECT_ID
        End Enum

        Sub New(myname_in As Information.CheckNull,
                project_in As Information.CheckNull,
                model_in As Information.CheckNull,
                release_in As Information.CheckNull,
                taskid_in As Information.TaskID,
                modulename_in As Information.CheckNull,
                mkschecker_in As Information.CheckNull,
                SubRev_in As Information.NoCheck,
                ModulePath_in As Information.ModulePath,
                Revision_in As Information.CheckNull,
                ResultPath_in As Information.CheckNull,
                Statement_in As Information.CheckNull,
                Decisions_in As Information.NoCheck,
                Defectid_in As Information.NoCheck,
                SysType_in As String)
            MyBase.New({myname_in,
                       project_in,
                       model_in,
                       release_in,
                       taskid_in,
                       modulename_in,
                       mkschecker_in,
                       SubRev_in,
                       ModulePath_in,
                       Revision_in,
                       ResultPath_in,
                       Statement_in,
                       Decisions_in,
                       Defectid_in})
            SysType = SysType_in
        End Sub

        Public Overrides Function GetToList() As Object
            Return listinfo(Position.MKS_CHECKER).GetValue
        End Function

        Public Overrides Function GetTitleEmail() As Object
            Return listinfo(Position.TASK_ID).GetValue + "_" + listinfo(Position.MODULE_NAME).GetValue + ": MKS Check In Request"
        End Function

        Public Overrides Function GetEmailBody() As Object
            Dim MailBody As String = Nothing
            MailBody = MailBody & "<font face='Bosch Office Sans' size=2>Hi checker,<BR><BR>"
            MailBody = MailBody & "Please help me to do check-in MKS:<BR><BR>"
            MailBody = MailBody & "o&nbsp;&nbsp;&nbsp;&nbsp;Project: <B>" & SysType & "</B>" + "<BR>"
            MailBody = MailBody & "o&nbsp;&nbsp;&nbsp;&nbsp;Sub-project revision: " + listinfo(Position.SUB_REV).GetValue + "<BR>"
            MailBody = MailBody & "o&nbsp;&nbsp;&nbsp;&nbsp;Path to module: " + listinfo(Position.MODULE_PATH).GetValue + "<BR>"
            MailBody = MailBody & "o&nbsp;&nbsp;&nbsp;&nbsp;Revision of Module: " + listinfo(Position.REVISION).GetValue + "<BR><BR>"
            MailBody = MailBody & "Description : " + "<BR>"
            MailBody = MailBody & "&nbsp;&nbsp;&nbsp;&nbsp;" + listinfo(Position.RESULT_PATH).GetValue + "<BR>"

            If InStr(listinfo(Position.DECISION).GetValue, "") <> 0 Then
                MailBody = MailBody & "&nbsp;&nbsp;&nbsp;&nbsp;C1: Statement: " + listinfo(Position.STATEMENT).GetValue + "%; Decisions: " _
                    + listinfo(Position.DECISION).GetValue + "% " + "<B> Dead code/No requirement/Unreachable code </B><BR><BR>"
            Else
                If listinfo(Position.STATEMENT).GetValue = "100" Then
                    MailBody = MailBody & "&nbsp;&nbsp;&nbsp;&nbsp; " + "C1: 100%" + " </B><BR><BR> "
                Else
                    MailBody = MailBody & "&nbsp;&nbsp;&nbsp;&nbsp; " + "C1: Statement " & listinfo(Position.STATEMENT).GetValue & "%" _
                        + "<B> Dead code/No requirement/Unreachable code </B><BR><BR>"
                End If

            End If

            ' Give defect information if it have
            If listinfo(Position.DEFECT_ID).GetValue <> "" Then
                MailBody = MailBody & "<B> Uncover requirement ( raised defect " + listinfo(Position.DEFECT_ID).GetValue + " ) </B><BR><BR>"
            End If
            Return MailBody
        End Function
    End Class

End Namespace
