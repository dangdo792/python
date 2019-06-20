Imports Scripting

Namespace Button
    Class AutoImportDOORsDoc
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            TOOL_FOLDER
            PROJECT
            MODEL
            TASK_ID
            MODULE_NAME
            MODULE_PATH
            RELEASE
        End Enum

        Private Result_Path As String = Nothing

        Public Sub New(explorer_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                       tfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       model_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath,
                       release_in As Information.CheckNull,
                       resultpath_in As String)
            MyBase.New({explorer_in,
                       wfd_in,
                       tfd_in,
                       project_in,
                       model_in,
                       taskid_in,
                       modulename_in,
                       modulepath_in,
                       release_in})
            Result_Path = resultpath_in
        End Sub

        Private Auto_Import_Door_Path As String
        Dim Testcase_Design_Path As String

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing
            Dim First_Testcase_ID As String = InputBox("Please Insert First Testcase ID")
            If Not String.IsNullOrEmpty(First_Testcase_ID) Then
                Dim oShell = CreateObject("WScript.Shell")
                oShell.Run("cmd /K " + Auto_Import_Door_Path + " -s " + Testcase_Design_Path + " -f " + First_Testcase_ID + " -r " & listinfo(Position.MODEL).GetValue & "_" & listinfo(Position.RELEASE).GetValue + " -l " + Result_Path + " & exit", 0, True)
                oShell = Nothing
            End If
            Return ErrorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim LinkObject As Information.CheckPathExist
            Dim IsValid As Boolean = True
            'Check Explorer is exist or not
            Auto_Import_Door_Path = listinfo(Position.TOOL_FOLDER).GetValue + "\AutoToolChain1.8\Auto_Import_DOORS\Auto_Import_DOORS.pl"
            LinkObject = New Information.CheckPathExist(Auto_Import_Door_Path)
            IsValid = LinkObject.IsValid()
            additional_errorMsg = LinkObject.GetErrorMsg() & vbNewLine
            If IsValid Then
                ' Check if the path is exist or not 
                Testcase_Design_Path = GetFullPath() & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_testcase_design" & ".xls"
                LinkObject = New Information.CheckPathExist(Testcase_Design_Path)
                IsValid = LinkObject.IsValid()
                additional_errorMsg = LinkObject.GetErrorMsg()
            End If

            Return IsValid
        End Function

        Function GetFullPath()
            Dim getdocpath As New Button.GotoDocument(listinfo(Position.EXPLORER), listinfo(Position.WORKING_FOLDER), listinfo(Position.PROJECT), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
            Return getdocpath.GetFullPath
        End Function

    End Class
End Namespace
