Imports System.Text.RegularExpressions

Namespace Button
    Public Class UploadResultToServer
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            TASK_ID
            PROJECT
            MODULE_NAME
            MODULE_PATH
        End Enum
        Private fullmodulename As String

        Private ResultPath As String

        Private MyChecklistPath As String
        Private MyipgcopPath As String
        Private MytestmkPath As String
        Private MytestfilePath As String
        Private MyreportPath As String


        Public ResultFullPath As String
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

            'Try to fine component name
            Dim component As String = Nothing
            Dim regex As Regex = New Regex("(?<=src\\)[^\\]+")
            Dim match As Match = regex.Match(listinfo(Position.MODULE_PATH).GetValue)
            If match.Success Then
                component = match.Value
            Else
                'not found. Try again with another way
                regex = New Regex("(?<=component\\)[^\\]+")
                match = regex.Match(listinfo(Position.MODULE_PATH).GetValue)
                If match.Success Then
                    component = match.Value
                Else
                    'not found. Try again with another way
                    regex = New Regex("(?<=dc_interfaces\\)[^\\]+")
                    match = regex.Match(listinfo(Position.MODULE_PATH).GetValue)
                    If match.Success Then
                        component = match.Value
                    Else
                        'Can't found to suggest user => user had to put by manual
                    End If
                End If
            End If

            Dim result = InputBox("Please put your component: ", , UCase(component))
            If String.IsNullOrEmpty(result) Then
                Return "No component to create OPL Path"
            Else
                component = result
            End If

            ResultFullPath = ResultPath & "\" & UCase(component) & "\" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "\" & listinfo(Position.TASK_ID).GetValue
            createfolder(ResultFullPath & "\Documents")
            createfolder(ResultFullPath & "\Scripts")
            createfolder(ResultFullPath & "\Reports")

            copyfile(MyChecklistPath, ResultFullPath & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_Review_Checklist.xlsx")
            copyfile(MyreportPath, ResultFullPath & "\Reports\test_report.html")

            copyfile(MyipgcopPath, ResultFullPath & "\Scripts\ipg.cop")
            copyfile(MytestmkPath, ResultFullPath & "\Scripts\test.mk")
            copyfile(MytestfilePath, ResultFullPath & "\Scripts\test_" & fullmodulename)
            Return errorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            'Check Result root path is exist
            Dim IsValid As Boolean = True
            Dim pathObj As Information.CheckPathExist = New Information.CheckPathExist(ResultPath)
            IsValid = pathObj.IsValid()
            additional_errorMsg = pathObj.GetErrorMsg()
            If IsValid Then
                'Check 1 file checklist is valid
                Dim getdocpath As New Button.GotoDocument(listinfo(Position.EXPLORER),
                                                      listinfo(Position.WORKING_FOLDER),
                                                          listinfo(Position.PROJECT),
                                                          listinfo(Position.TASK_ID),
                                                          listinfo(Position.MODULE_NAME))
                MyChecklistPath = getdocpath.GetFullPath() & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_Review_Checklist.xlsx"
                pathObj = New Information.CheckPathExist(MyChecklistPath)
                IsValid = IsValid And pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

                'Check 1 file report Is valid
                MyreportPath = getdocpath.GetFullPath() & "\Reports\test_report.html"
                pathObj = New Information.CheckPathExist(MyreportPath)
                IsValid = IsValid And pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

                'Check 1 file ipgcop is valid
                MyipgcopPath = getdocpath.GetFullPath() & "\Scripts\ipg.cop"
                pathObj = New Information.CheckPathExist(MyipgcopPath)
                IsValid = IsValid And pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

                'Check 1 file testmk is valid
                MytestmkPath = getdocpath.GetFullPath() & "\Scripts\test.mk"
                pathObj = New Information.CheckPathExist(MytestmkPath)
                IsValid = IsValid And pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine

                'Check 1 file testfile is valid
                Dim aPath() As String
                aPath = Split(listinfo(Position.MODULE_PATH).GetValue, "\")
                fullmodulename = aPath(UBound(aPath))
                If InStr(fullmodulename, ".inl") <> 0 Then fullmodulename = fullmodulename.Replace(".inl", ".cpp")
                MytestfilePath = getdocpath.GetFullPath() & "\Scripts\test_" & fullmodulename
                pathObj = New Information.CheckPathExist(MytestfilePath)
                IsValid = IsValid And pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg() & vbNewLine & vbNewLine
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
