Imports System.IO
Imports System.IO.Compression

Namespace Button
    Class ExportScriptReport
        Inherits ButtonBase

        Private Enum Position
            WORKING_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
            SANDBOX_FOLDER
            MODULE_PATH
            SANDBOX
        End Enum

        Private SubFolder() As String = {"Documents", "Emails", "Scripts", "Reports"}

        Public Sub New(wfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       sfd_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath,
                       sandbox_in As Information.CheckNull)
            MyBase.New({wfd_in,
                       project_in,
                       taskid_in,
                       modulename_in,
                       sfd_in,
                       modulepath_in,
                       sandbox_in})
        End Sub

        Private TaskFolderPath As String = Nothing
        Private TestFolderPath As String = Nothing

        Private fullmodulename As String
        Private testreportpath As String

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing

            'Copy 1 file report
            copyfile(testreportpath & "\test_report.html", TaskFolderPath & "\Reports\test_report.html")

            'Copy 3 files script
            copyfile(TestFolderPath & "\ipg.cop", TaskFolderPath & "\Scripts\ipg.cop")
            copyfile(TestFolderPath & "\test.mk", TaskFolderPath & "\Scripts\test.mk")
            copyfile(TestFolderPath & "\test_" & fullmodulename, TaskFolderPath & "\Scripts\test_" & fullmodulename)

            Dim startPath As String = listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue + "\_test\" & LCase(listinfo(Position.MODULE_NAME).GetValue)
            Dim zipPath As String = TaskFolderPath & "\Scripts\" & LCase(listinfo(Position.MODULE_NAME).GetValue) & ".zip"



            Dim pathObj As New Information.CheckPathExist(zipPath)
            Dim isValid As Boolean = pathObj.IsValid()
            'Check if zip file is exist
            If isValid Then
                'Deleted old zip file
                My.Computer.FileSystem.DeleteFile(zipPath)
            End If

            Try
                ZipFile.CreateFromDirectory(startPath, zipPath)
            Catch ex As Exception
                ErrorMsg = ex.Message
            End Try



            Return ErrorMsg
        End Function



        Overrides Function AdditionCondition() As Boolean
            'Check working path is valid
            TaskFolderPath = listinfo(Position.WORKING_FOLDER).GetValue + "\" & UCase(listinfo(Position.PROJECT).GetValue) & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue
            Dim pathObj As New Information.CheckPathExist(TaskFolderPath & "\Scripts")
            Dim isValid As Boolean = pathObj.IsValid()
            additional_errorMsg = pathObj.GetErrorMsg() & vbNewLine & vbNewLine

            pathObj = New Information.CheckPathExist(TaskFolderPath & "\Reports")
            isValid = isValid AndAlso pathObj.IsValid()
            additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

            If isValid Then
                'Check 1 file report is valid
                testreportpath = listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue + "\_test\" & LCase(listinfo(Position.MODULE_NAME).GetValue) & "\Cantata\results"
                pathObj = New Information.CheckPathExist(testreportpath & "\test_report.html")
                isValid = isValid AndAlso pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

                'Check 3 files path is valid
                TestFolderPath = listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue & "\_test\" & LCase(listinfo(Position.MODULE_NAME).GetValue) & "\Cantata\tests\test_" & LCase(listinfo(Position.MODULE_NAME).GetValue)
                pathObj = New Information.CheckPathExist(TestFolderPath & "\ipg.cop")
                isValid = isValid AndAlso pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

                pathObj = New Information.CheckPathExist(TestFolderPath & "\test.mk")
                isValid = isValid AndAlso pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

                Dim aPath() As String
                aPath = Split(listinfo(Position.MODULE_PATH).GetValue, "\")
                fullmodulename = aPath(UBound(aPath))
                If InStr(fullmodulename, ".inl") <> 0 Then fullmodulename = fullmodulename.Replace(".inl", ".cpp")
                pathObj = New Information.CheckPathExist(TestFolderPath & "\test_" & fullmodulename)
                isValid = isValid AndAlso pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine
            End If
            Return isValid
        End Function

        Public Sub copyfile(sourcepath As String, despath As String)
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            If Not fso.FileExists(despath) Then
                Call fso.CopyFile(sourcepath, despath, True)
            End If
        End Sub
    End Class

End Namespace
