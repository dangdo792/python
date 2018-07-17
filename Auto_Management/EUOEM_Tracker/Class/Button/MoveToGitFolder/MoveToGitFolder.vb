
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq

Namespace Button
    Class MoveToGitFolder
        Inherits ButtonBase

        Dim IsRevTab As Boolean

        ''' <summary>
        ''' The position of first 7 field is a must, do not adapt it
        ''' </summary>
        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
            SANDBOX_FOLDER
            SANDBOX
            REVIEW_FOLDER
        End Enum


        Public Sub New(explorer_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       sfd_in As Information.CheckNull,
                       sandbox_in As Information.CheckNull,
                       rfd_in As Information.CheckNull,
                       IsRevTab_in As Boolean)

            MyBase.New({explorer_in,
                        wfd_in,
                        project_in,
                        taskid_in,
                        modulename_in,
                        sfd_in,
                        sandbox_in,
                        rfd_in})
            IsRevTab = IsRevTab_in
        End Sub

        Private ZipFolderPath As String
        Private TestFolderPath As String

        Public Overrides Function DoFunctionality() As String
            Dim errormsg As String = Nothing
            'Check and create folder _test
            createfolder(TestFolderPath & "\_test")


            Dim zipPath As String = ZipFolderPath
            Dim extractPath As String = TestFolderPath & "\_test"

            Dim aPath() As String
            aPath = Split(zipPath, "\")

            Dim abc = extractPath & "\" & aPath(UBound(aPath)).Replace(".zip", "")


            Dim path As String = extractPath & "\" & aPath(UBound(aPath)).Replace(".zip", "")
            'Check if folder is exist
            If IO.Directory.Exists(path) Then
                'Ask user to overwrite
                Dim result = MsgBox(path & " is exist." & vbNewLine & "Would you like to overwrite?", vbOKCancel)
                If result = vbOK Then
                    'Deleted old folder
                    System.IO.Directory.Delete(path, True)
                Else
                    Return Nothing
                End If

            End If

            Try
                ZipFile.ExtractToDirectory(ZipFolderPath, TestFolderPath & "\_test")
            Catch ex As Exception
                errormsg = ex.Message
            End Try

            Return errormsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim IsValid As Boolean = True
            'Check Zip file is exist
            If IsRevTab = True Then
                Dim GotoRev As New Button.GotoReview(listinfo(Position.EXPLORER), listinfo(Position.REVIEW_FOLDER), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
                ZipFolderPath = GotoRev.GetFullPath & "\" & "Scripts" & "\" & listinfo(Position.MODULE_NAME).GetValue & ".zip"
            Else
                Dim GotoDoc As New Button.GotoDocument(listinfo(Position.EXPLORER), listinfo(Position.WORKING_FOLDER), listinfo(Position.PROJECT), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
                ZipFolderPath = GotoDoc.GetFullPath & "\" & "Scripts" & "\" & listinfo(Position.MODULE_NAME).GetValue & ".zip"
            End If
            Dim isExistObject As Information.CheckPathExist = New Information.CheckPathExist(ZipFolderPath)
            IsValid = isExistObject.IsValid()
            additional_errorMsg = isExistObject.GetErrorMsg()
            If IsValid Then
                'Check test folder path is exist
                TestFolderPath = listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue
                isExistObject = New Information.CheckPathExist(TestFolderPath)
                IsValid = isExistObject.IsValid()
                additional_errorMsg = isExistObject.GetErrorMsg()
            End If
            Return IsValid
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

    End Class
End Namespace
