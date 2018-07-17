Imports System.Text.RegularExpressions
Imports Scripting

Namespace Button
    Class FillHeader
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            SANDBOX_FOLDER
            PROJECT
            MODEL
            RELEASE
            TASK_ID
            MODULE_NAME
            REVISION
            SANDBOX
            BRANCH
            MODULE_PATH
            OLD_TASK
        End Enum

        Public Sub New(explorer_in As Information.CheckNull,
                       sfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       model_in As Information.CheckNull,
                       release_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       revision_in As Information.NoCheck,
                       sandbox_in As Information.CheckNull,
                       branch_in As Information.NoCheck,
                       modulepath_in As Information.ModulePath,
                       oldtask_in As Information.NoCheck)
            MyBase.New({explorer_in,
                       sfd_in,
                       project_in,
                       model_in,
                       release_in,
                       taskid_in,
                       modulename_in,
                       revision_in,
                       sandbox_in,
                       branch_in,
                       modulepath_in,
                       oldtask_in})
        End Sub

        Private Script_File_Dir As String

        Public Overrides Function DoFunctionality() As String

            Dim fso As FileSystemObject
            Dim TS As TextStream
            Dim Final As String = Nothing
            fso = New FileSystemObject

            TS = fso.OpenTextFile(Script_File_Dir, IOMode.ForReading)
            Dim listlines() = Split(TS.ReadAll, vbCrLf)

            Dim comment As String = Nothing
            If Right(listinfo(Position.MODULE_PATH).GetValue, 2) = ".c" Then
                comment = "-- "
            ElseIf Right(listinfo(Position.MODULE_PATH).GetValue, 4) = ".cpp" Then
                comment = "// "
            End If

            listlines(1) = comment & "Customer = " & listinfo(Position.PROJECT).GetValue
            listlines(2) = comment & "Model = " & listinfo(Position.MODEL).GetValue
            listlines(3) = comment & "File_Name = " & listinfo(Position.MODULE_PATH).GetValue

            Dim Task_ID_Comment As String = Nothing
            If InStr(LCase(listinfo(Position.TASK_ID).GetValue), "nprod") = 0 Then
                Task_ID_Comment = "Jira_Issue"
            Else
                Task_ID_Comment = "nprod_number_info"
            End If

            If listinfo(Position.OLD_TASK).GetValue = "" Then
                listlines(4) = comment & Task_ID_Comment & " = " & listinfo(Position.TASK_ID).GetValue
            Else
                listlines(4) = comment & Task_ID_Comment & " = " & listinfo(Position.OLD_TASK).GetValue & ", " & listinfo(Position.TASK_ID).GetValue
            End If
            If Regex.Matches(listinfo(Position.REVISION).GetValue, "^[0-9.]+$").Count = 0 Then
                listlines(5) = comment & "Commit = " & listinfo(Position.REVISION).GetValue
            Else
                listlines(5) = comment & "MKS_File_Revision = " & listinfo(Position.REVISION).GetValue
            End If

            If Not String.IsNullOrEmpty(listinfo(Position.BRANCH).GetValue) Then
                listlines(6) = comment & "Features_Branch = " & listinfo(Position.BRANCH).GetValue
            Else
                listlines(6) = comment & "MKS_version_CP = " & listinfo(Position.SANDBOX).GetValue
            End If


            listlines(7) = comment & "Release_info = " & listinfo(Position.RELEASE).GetValue

            If Right(listinfo(Position.MODULE_PATH).GetValue, 2) = ".c" Then
                listlines(0) = "--********************************************************************************"
                listlines(8) = "--********************************************************************************"
                If Regex.Matches(listinfo(Position.REVISION).GetValue, "^[0-9.]+$").Count = 0 And Not String.IsNullOrEmpty(listinfo(Position.BRANCH).GetValue) Then
                    listlines(9) = "HEADER " & listinfo(Position.MODULE_NAME).GetValue + ".c, " & listinfo(Position.BRANCH).GetValue + ", " + listinfo(Position.REVISION).GetValue
                Else
                    listlines(9) = "HEADER " & listinfo(Position.MODULE_NAME).GetValue + ".c, " & listinfo(Position.SANDBOX).GetValue + ", " + listinfo(Position.REVISION).GetValue
                End If
                listlines(10) = ""
            End If

            Final = Join(listlines, vbCrLf)

            TS.Close()

            TS = fso.OpenTextFile(Script_File_Dir, IOMode.ForWriting, True)
            TS.Write(Final)
            TS.Close()

            TS = Nothing
            fso = Nothing

            Return Nothing
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim path As String = GetFullPath()

            ' Get the path of script
            Dim extension As String = Nothing
            If Microsoft.VisualBasic.Right(listinfo(Position.MODULE_PATH).GetValue, 2) = ".c" Then
                extension = ".ptu"
            ElseIf Microsoft.VisualBasic.Right(listinfo(Position.MODULE_PATH).GetValue, 4) = ".cpp" Then
                extension = ".otd"
            End If
            Script_File_Dir = listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue _
                    & Left(listinfo(Position.MODULE_PATH).GetValue, InStrRev(listinfo(Position.MODULE_PATH).GetValue, "\") - 1) & "\_test\" _
                    & listinfo(Position.MODULE_NAME).GetValue & "\" & listinfo(Position.MODULE_NAME).GetValue & extension

            ' Check if the path is exist or not
            Dim DirObject As Information.CheckPathExist = New Information.CheckPathExist(Script_File_Dir)
            Dim isValidFlag = DirObject.IsValid()
            additional_errorMsg = DirObject.GetErrorMsg()

            If Not isValidFlag And String.IsNullOrEmpty(listinfo(Position.REVISION).GetValue) Then
                additional_errorMsg = "Revision are empty to fill script. Please put."
                isValidFlag = False
            End If

            Return isValidFlag
        End Function

        Function GetFullPath()
            Dim getsandpath As New Button.GotoSandBox(listinfo(Position.EXPLORER), listinfo(Position.SANDBOX_FOLDER), listinfo(Position.SANDBOX), listinfo(Position.MODULE_PATH))
            Return getsandpath.GetFullPath
        End Function

    End Class

End Namespace
