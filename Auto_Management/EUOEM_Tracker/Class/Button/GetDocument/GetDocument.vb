Imports Microsoft.Office.Interop.Excel

Namespace Button
    Class GetDocument
        Inherits ButtonBase

        Private Enum Position
            WORKING_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
            OLD_TASK
            MY_NAME
            FILES_TEMPLATE_DIR
        End Enum

        Private SubFolder() As String = {"Documents", "Emails", "Scripts", "Reports"}

        Public Sub New(wfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       oldtask_in As Information.NoCheck,
                       myname_in As Information.CheckNull,
                       filestemplatedir_in As Information.CheckNull)
            MyBase.New({wfd_in,
                       project_in,
                       taskid_in,
                       modulename_in,
                       oldtask_in,
                       myname_in,
                       filestemplatedir_in})
        End Sub

        Public TaskFolderPath As String

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing

            ' Download files
            Dim taskid_double As Double
            Double.TryParse(Search_f(listinfo(Position.TASK_ID).GetValue, "\d+"), taskid_double)

            TaskFolderPath = listinfo(Position.WORKING_FOLDER).GetValue & "\" & UCase(listinfo(Position.PROJECT).GetValue) & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue

            'Create Folder
            Dim SubFolderPath As String = Nothing

            For i = 0 To 3 : SubFolderPath = TaskFolderPath & "\" & SubFolder(i) : createfolder(SubFolderPath) : Next
            If listinfo(Position.OLD_TASK).GetValue <> "" Then SubFolderPath = TaskFolderPath & "\" & listinfo(Position.OLD_TASK).GetValue : createfolder(TaskFolderPath)

            Dim templateDocPath As String = listinfo(Position.FILES_TEMPLATE_DIR).GetValue

            If InStr(templateDocPath, "https://") <> 0 Then
                Call ILM_DownloadItemInFolder(templateDocPath, TaskFolderPath & "\Documents", MainF.OutputTextBox)
            Else
                'Check Template documents is existed
                ErrorMsg = ErrorMsg & IsFileExist(templateDocPath & "\" & "TaskIssue_UnitName_Review_Checklist.xlsx")
                ErrorMsg = ErrorMsg & IsFileExist(templateDocPath & "\" & "TaskIssue_UnitName_test_analysis.xls")
                ErrorMsg = ErrorMsg & IsFileExist(templateDocPath & "\" & "UnitName_OPL.xls")
                ErrorMsg = ErrorMsg & IsFileExist(templateDocPath & "\" & "TaskIssue_UnitName_CodeCoverage_Exception.xls")
                If Not String.IsNullOrEmpty(ErrorMsg) Then
                    GC.Collect()
                    ErrorMsg = "Those of document templates aren't existed." & vbNewLine & ErrorMsg & vbNewLine
                    ErrorMsg = ErrorMsg & "Please check whether they exist in?: " & templateDocPath
                    Return ErrorMsg
                    Exit Function
                End If

                'Copy Document
                copyfile(templateDocPath & "\" & "TaskIssue_UnitName_Review_Checklist.xlsx",
                            TaskFolderPath & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_Review_Checklist.xlsx")
                copyfile(templateDocPath & "\" & "TaskIssue_UnitName_test_analysis.xls",
                            TaskFolderPath & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_test_analysis.xls")
                copyfile(templateDocPath & "\" & "UnitName_OPL.xls",
                            TaskFolderPath & "\" & "Documents" & "\" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "_OPL.xls")
                copyfile(templateDocPath & "\" & "TaskIssue_UnitName_CodeCoverage_Exception.xls",
                            TaskFolderPath & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "_CodeCoverage_Exception.xls")
            End If

            ' Fill test analysis file
            Dim testanalysisWB =
                New ExcelHandle(TaskFolderPath & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_test_analysis.xls")

            Dim WB As Workbook
            Dim WS As Worksheet
            WB = testanalysisWB.Get_WB
            WS = WB.Worksheets("Revision History")
            WS.Range("C17").Value = listinfo(Position.MY_NAME).GetValue
            WS.Range("D17").Value = CStr(Now().Date)

            WS = WB.Worksheets("Test_analysis_RBT")
            If listinfo(Position.OLD_TASK).GetValue <> "" Then
                WS.Range("C2").Value = listinfo(Position.OLD_TASK).GetValue
            Else
                'WS.Range("F:I").EntireColumn.Delete()
                WS.Range("C2").Value = "N.A"
            End If


            ' Close all workbooks
            testanalysisWB.CloseWB()

            'Kill Object
            testanalysisWB = Nothing
            GC.Collect()

            Return Nothing
        End Function

        Overrides Function AdditionCondition() As Boolean
            ' Check if excel editable or not
            Dim temp As ExcelHandle = New ExcelHandle("")
            If temp.CannotUse() Then
                additional_errorMsg = "You are editing cell. Please check and release cell." & vbNewLine
                Return False
            Else
                additional_errorMsg = Nothing
                Return True
            End If
            temp.CloseWB()
            temp = Nothing
            GC.Collect()
        End Function

        Public Function IsFileExist(ByVal FileDir As String) As String
            Dim ErrorMsg As String = Nothing
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            If Not fso.FileExists(FileDir) Then
                ErrorMsg = reg.Search_f(FileDir, "\\(?:.(?!\\))+$")
                If ErrorMsg IsNot Nothing Then
                    ErrorMsg = "- " & ErrorMsg.Replace("\", "") & vbNewLine
                End If
            End If
            Return ErrorMsg
        End Function

        Public Sub createfolder(ByVal folderpath As String)
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            Try
                If Not IO.Directory.Exists(folderpath) Then
                    IO.Directory.CreateDirectory(folderpath)
                End If
            Catch ex As Exception
                MainF.OutputTextBox.Text = ex.Message
            End Try

        End Sub

        Public Sub copyfile(sourcepath As String, despath As String)
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            If Not fso.FileExists(despath) Then
                Call fso.CopyFile(sourcepath, despath)
            End If
        End Sub
    End Class

End Namespace
