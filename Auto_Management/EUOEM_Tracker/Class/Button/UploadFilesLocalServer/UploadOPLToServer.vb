Imports System.Text.RegularExpressions

Namespace Button
    Public Class UploadOPLToServer
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            TASK_ID
            PROJECT
            MODULE_NAME
            MODULE_PATH
            OPL_LINK
        End Enum
        Private MyOplPath As String
        Private ResultPath As String

        Public OPLFullPath As String
        Public Sub New(explorerpath_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                      taskid_in As Information.TaskID,
                       project_in As Information.CheckNull,
                       modulename_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath,
                       ResultPath_in As String)
            MyBase.New({explorerpath_in,
                        wfd_in,
                       taskid_in,
                       project_in,
                       modulename_in,
                       modulepath_in})
            ResultPath = ResultPath_in
        End Sub

        Public Overrides Function DoFunctionality() As String
            Dim errorMsg As String = Nothing

            Dim component As String = Nothing
            Dim regex As Regex = New Regex("(?<=src\\)[^\\]+")
            Dim match As Match = regex.Match(listinfo(Position.MODULE_PATH).GetValue)
            If match.Success Then
                component = match.Value
            Else
                regex = New Regex("(?<=component\\)[^\\]+")
                match = regex.Match(listinfo(Position.MODULE_PATH).GetValue)
                If match.Success Then
                    component = match.Value
                End If
            End If

            Dim result = InputBox("Please put your component: ", , UCase(component))
            If String.IsNullOrEmpty(result) Then
                Return "No component to create OPL Path"
            Else
                component = result
            End If


            OPLFullPath = ResultPath & "\" & UCase(component) & "\" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "\OPL"
            createfolder(OPLFullPath)
            copyfile(MyOplPath, OPLFullPath & "\" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "_OPL.xls")
            Return errorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            'Check My local OPL file is exist.
            Dim IsValid As Boolean = True
            additional_errorMsg = Nothing
            Dim getdocpath As New Button.GotoDocument(listinfo(Position.EXPLORER),
                                                      listinfo(Position.WORKING_FOLDER),
                                                      listinfo(Position.PROJECT),
                                                      listinfo(Position.TASK_ID),
                                                      listinfo(Position.MODULE_NAME))
            MyOplPath = getdocpath.GetFullPath() & "\" & "Documents" & "\" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "_OPL.xls"
            Dim isExistObject As Information.CheckPathExist = New Information.CheckPathExist(MyOplPath)
            IsValid = isExistObject.IsValid()
            additional_errorMsg = isExistObject.GetErrorMsg()
            If IsValid Then
                'Check Result Link is exist
                isExistObject = New Information.CheckPathExist(ResultPath)
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

        Public Sub copyfile(sourcepath As String, despath As String)
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            If Not fso.FileExists(despath) Then
                Call fso.CopyFile(sourcepath, despath)
            End If
        End Sub
    End Class
End Namespace
