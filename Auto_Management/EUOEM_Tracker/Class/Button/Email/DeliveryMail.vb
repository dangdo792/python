Imports System.Text.RegularExpressions

Namespace Button
    ''' <summary>
    ''' Delivery email
    ''' </summary>
    Class DeliveryEmail
        Inherits EmailBase

        Private TaskCoor As String

        Private Enum Position
            MYNAME
            PROJECT
            MODEL
            RELEASE
            TASK_ID
            MODULE_NAME
            MODULE_OWNER
            MODULE_PATH
            REVISION
            RESULT_PATH
            STATEMENT
            DECISION
            DEFECT_ID
            PACKAGE_ID
            RS
            RS_BL
            TS
            TS_BL
            SD
            SD_BL
            OPL
            COVER_PATH
            BRANCH
            SANDBOX
            ELOC
            PM
            TEAMLEAD
            REVIEWER
        End Enum

        Sub New(myname_in As Information.CheckNull,
                project_in As Information.CheckNull,
                model_in As Information.CheckNull,
                release_in As Information.CheckNull,
                taskid_in As Information.TaskID,
                modulename_in As Information.CheckNull,
                mo_in As Information.CheckNull,
                ModulePath_in As Information.ModulePath,
                Revision_in As Information.CheckNull,
                ResultPath_in As Information.CheckNull,
                Statement_in As Information.CheckNull,
                Decisions_in As Information.NoCheck,
                Defectid_in As Information.NoCheck,
                packageid_in As Information.NoCheck,
                rs_in As Information.CheckNull,
                rsbl_in As Information.DoorBaseline,
                ts_in As Information.CheckNull,
                tsbl_in As Information.DoorBaseline,
                sd_in As Information.NoCheck,
                sdbl_in As Information.NoCheck,
                OPL_in As Information.NoCheck,
                coverpath_in As Information.NoCheck,
                branch_in As Information.NoCheck,
                sandbox_in As Information.CheckNull,
                eloc_in As Information.CheckNull,
                pm_in As Information.CheckNull,
                teamlead_in As Information.CheckNull,
                reviewer_in As Information.CheckNull,
                TaskCoor_in As String)
            MyBase.New({myname_in,
                       project_in,
                       model_in,
                       release_in,
                       taskid_in,
                       modulename_in,
                       mo_in,
                       ModulePath_in,
                       Revision_in,
                       ResultPath_in,
                       Statement_in,
                       Decisions_in,
                       Defectid_in,
                       packageid_in,
                       rs_in,
                       rsbl_in,
                       ts_in,
                       tsbl_in,
                       sd_in,
                       sdbl_in,
                       OPL_in,
                        coverpath_in,
                        branch_in,
                        sandbox_in,
                        eloc_in,
                        pm_in,
                        teamlead_in,
                        reviewer_in})
            TaskCoor = TaskCoor_in
        End Sub

        Public Overrides Function GetToList() As Object
            Return listinfo(Position.MODULE_OWNER).GetValue
        End Function

        Public Overrides Function GetCCList() As Object
            Return listinfo(Position.PM).GetValue & "; " & listinfo(Position.TEAMLEAD).GetValue & "; " & TaskCoor & "; " & listinfo(Position.REVIEWER).GetValue & "; " & listinfo(Position.MYNAME).GetValue
        End Function

        Public Overrides Function GetTitleEmail() As Object
            Return "Delivery " + listinfo(Position.TASK_ID).GetValue + "_" + listinfo(Position.MODULE_NAME).GetValue + " RTRT Module test"
        End Function

        Public Overrides Function GetEmailBody() As Object
            Dim MailBody As String = Nothing
            MailBody = MailBody & "<font face='Arial' size=2>"

            MailBody = MailBody & "&#60;Change Package ID>" + "&nbsp;&nbsp;" + listinfo(Position.PACKAGE_ID).GetValue + "<BR><BR>"
            MailBody = MailBody & "****************************************************************************************************************************************************************<BR><BR>"
            MailBody = MailBody & "&#60;SWMRS>" + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & " <a href=" + listinfo(Position.RS).GetValue + ">" _
                + listinfo(Position.RS).GetValue + "</a>" + "&nbsp;&nbsp;" + " Baseline: " + listinfo(Position.RS_BL).GetValue + "</a><BR><BR>"
            If Not String.IsNullOrEmpty(listinfo(Position.SD).GetValue) Then
                MailBody = MailBody & "&#60;SWMSD>" + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & " <a href=" + listinfo(Position.SD).GetValue + ">" _
                    + listinfo(Position.SD).GetValue + "</a>" + "&nbsp;&nbsp;" + " Baseline: " + listinfo(Position.SD_BL).GetValue + "</a><BR><BR>"
            End If

            MailBody = MailBody & "&#60;SWMTS>" + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + " <a href=" + listinfo(Position.TS).GetValue + ">" _
                + listinfo(Position.TS).GetValue + "</a>" + "&nbsp;&nbsp;" + " Baseline: " + listinfo(Position.TS_BL).GetValue + "</a><BR><BR>"
            MailBody = MailBody & "&#60;Module Test Result> <a href=" + listinfo(Position.RESULT_PATH).GetValue + ">" _
                + listinfo(Position.RESULT_PATH).GetValue + "</a><BR><BR>"
            If listinfo(Position.OPL).GetValue = "" Then
                MailBody = MailBody & "&#60;OPL link>" + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & "No" + "<BR>"
            Else
                MailBody = MailBody & "&#60;OPL link>" + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" _
                    & "Yes. It has been clarified and closed. You can find information about OPL in the following document: " _
                    + " <a href=" + listinfo(Position.OPL).GetValue + ">" + listinfo(Position.OPL).GetValue + "</a>" + "<BR>"
            End If

            If listinfo(Position.DECISION).GetValue = "" Then
                MailBody = MailBody & "&#60;Code Coverage>" & "&nbsp;&nbsp;" & "C1:100%" + "<BR>"
            Else
                MailBody = MailBody & "&#60;Code Coverage>" & "&nbsp;&nbsp;" & "Statement: " & listinfo(Position.STATEMENT).GetValue & "%; " _
                    + "Decisions: " + listinfo(Position.DECISION).GetValue + "%" + "<BR>"
            End If

            If listinfo(Position.COVER_PATH).GetValue = "" Then
                MailBody = MailBody & "&#60;Code Coverage Exception> No." + "<BR>"
            Else
                MailBody = MailBody & "&#60;Code Coverage Exception> Yes. You can find information about code coverage exception in the following document: " _
                    + " <a href=" + listinfo(Position.COVER_PATH).GetValue + ">" + listinfo(Position.COVER_PATH).GetValue + "</a>" + "<BR>"
            End If

            If listinfo(Position.DEFECT_ID).GetValue = "" Then
                MailBody = MailBody & "&#60;Defect CQA>" + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + "No." + "</a><BR>"
            Else
                MailBody = MailBody & "&#60;Defect CQA>" + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + listinfo(Position.DEFECT_ID).GetValue + "</a><BR>"
            End If
            MailBody = MailBody & "******************************************************************************************************************<BR><BR>"

            MailBody = MailBody & "-File Name: " + listinfo(Position.MODULE_PATH).GetValue + "<BR>"
            If Regex.Matches(listinfo(Position.REVISION).GetValue, "^[0-9.]+$").Count = 0 Then
                MailBody = MailBody & "-Tag name: " + Char.IsLetter(listinfo(Position.REVISION).GetValue) + "<BR>"
                MailBody = MailBody & "-Features branch: " + listinfo(Position.BRANCH).GetValue + "<BR>"
            Else
                MailBody = MailBody & "-CP: " + listinfo(Position.SANDBOX).GetValue + "<BR>"
                MailBody = MailBody & "-Revision: " + listinfo(Position.REVISION).GetValue + "<BR>"
            End If
            MailBody = MailBody & "-Total lines of code: " + "" + "<BR>"
            MailBody = MailBody & "-ELOC (tested lines): " + listinfo(Position.ELOC).GetValue + "<BR>"
            MailBody = MailBody & "-No. of requirements tested: 0<BR>"

            MailBody = MailBody & "******************************************************************************************************************<BR>"
            MailBody = MailBody & "</BODY></HTML><BR><BR>"

            Dim Nprod_Hyberlink = "<a href=" & Chr(34) & "https://rb-wam.bosch.com/tracker02/browse/" & listinfo(Position.TASK_ID).GetValue & Chr(34) & ">" & listinfo(Position.TASK_ID).GetValue & "</a>"
            Dim Defect_Hyberlink = "<a href=" & Chr(34) & "https://rb-wam.bosch.com/tracker02/browse/" & listinfo(Position.DEFECT_ID).GetValue & Chr(34) & ">" & listinfo(Position.DEFECT_ID).GetValue & "</a>"
            Dim MOName() As String = Split(listinfo(Position.MODULE_OWNER).GetValue, " ")
            If MOName.Length > 1 Then
                MailBody = MailBody & "<font face='Arial' size=2>Hello " & MOName(1) & ",<BR><BR>"
            Else
                MailBody = MailBody & "<font face='Arial' size=2>Hello " & MOName(0) & ",<BR><BR>"
            End If

            MailBody = MailBody & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The Testing of " _
                + Nprod_Hyberlink + " RTRT module test has been completed and related information are updated accordingly.<BR><BR>"
            MailBody = MailBody & "&#60;JIRA Status>: In Progress<BR><BR>"

            If listinfo(Position.DECISION).GetValue = "" Then
                MailBody = MailBody & "&#60;Code Coverage>: " + "C1:100%" + "<BR><BR>"
            Else
                MailBody = MailBody & "&#60;Code Coverage>: " + "Statement: " & listinfo(Position.STATEMENT).GetValue & "%, " + "Decisions: " _
                    + listinfo(Position.DECISION).GetValue + "%" + "<BR><BR>"
            End If

            If listinfo(Position.COVER_PATH).GetValue <> "" Then
                MailBody = MailBody & "&#60;Code Coverage Exception>: <a href=" + listinfo(Position.COVER_PATH).GetValue + ">" _
                    + listinfo(Position.COVER_PATH).GetValue + "</a><BR><BR>"
            End If

            If listinfo(Position.DEFECT_ID).GetValue <> "" Then
                MailBody = MailBody & "There is a failed test case which is mentioned as below JIRA <BR>&#60;Defect JIRA>: " _
                    + Defect_Hyberlink + ". Please take care to close this JIRA.<BR><BR>"
            End If

            MailBody = MailBody & "Could you please cross check delivery artifacts and close this task accordingly?<BR><BR>"

            MailBody = MailBody & "[Delivery-Module Test]" + "[" & listinfo(Position.PROJECT).GetValue & "_" & listinfo(Position.MODEL).GetValue & "]" _
                + " " + listinfo(Position.MODULE_NAME).GetValue & "<BR><BR><BR><BR>"
            Return MailBody
        End Function

        Overrides Function AdditionCondition() As Boolean
            If Not String.IsNullOrEmpty(listinfo(Position.DECISION).GetValue) And String.IsNullOrEmpty(listinfo(Position.COVER_PATH).GetValue) Then
                additional_errorMsg = "There has a coverage but coverage path is empty." & vbNewLine & "Please update code coverage path"
                Return False
            Else
                additional_errorMsg = Nothing
                Return True
            End If
        End Function
    End Class

End Namespace
