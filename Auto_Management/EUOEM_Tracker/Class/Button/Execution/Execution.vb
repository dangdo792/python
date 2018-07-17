Imports System.IO
Imports System.Text.RegularExpressions
Imports Scripting

Namespace Button

    ''' <summary>
    ''' Auto Execution button
    ''' </summary>
    Class AutoExecution
        Inherits ButtonBase

        ''' <summary>
        ''' Diab: the bmk information
        ''' </summary>
        Private BMKName As String
        Private flist_file As String
        Private run_Cantata_execution_file As String

        Private testlocationpath As String

        Private IsCMakeBuilt As Boolean

        Private Enum Position
            SANDBOX_FOLDER
            TOOL_FOLDER
            SANDBOX
            MODULE_PATH
            C_MACRO
            CPP_MACRO
            SYSBPLUS_MACRO
        End Enum

        Public Sub New(sfd_in As Information.CheckNull,
                       tfd_in As Information.CheckNull,
                       sandbox_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath,
                       cmacro_in As Information.CheckNull,
                       cppmacro_in As Information.CheckNull,
                       sysbplus_in As Information.CheckNull)
            MyBase.New({sfd_in,
                       tfd_in,
                       sandbox_in,
                       modulepath_in,
                       cmacro_in,
                       cppmacro_in,
                       sysbplus_in})
        End Sub

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing

            Dim fso As FileSystemObject
            Dim TS As TextStream
            Dim Final As String = Nothing
            Dim listlines() As String
            fso = New FileSystemObject

            'Try to Get Comoponent
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

            Dim build_folder As String = Nothing
            Dim result2 = InputBox("Please put your build folder: ", , "build_cmake")
            If String.IsNullOrEmpty(result2) Then
                Return "No Build folder to execute cantata"
            Else
                build_folder = result2
            End If

            'Update flist_file with module path
            TS = fso.OpenTextFile(flist_file, IOMode.ForReading)
            listlines = Split(TS.ReadAll, vbCrLf)
            listlines(0) = "$(ROOT)" & listinfo(Position.MODULE_PATH).GetValue.Replace("\", "/")

            Final = Join(listlines, vbCrLf)

            TS.Close()

            TS = fso.OpenTextFile(flist_file, IOMode.ForWriting, True)
            TS.Write(Final)
            TS.Close()

            'Update cantata_auto_execution bat file
            TS = fso.OpenTextFile(run_Cantata_execution_file, IOMode.ForReading)

            Dim FileContent As String = TS.ReadAll

            FileContent = reg.Replace_f(FileContent, "SET BUILT_FOLDER_GEN5_LOCATION=.*$", "SET BUILT_FOLDER_GEN5_LOCATION=""" & testlocationpath & "\" & build_folder & """")
            FileContent = reg.Replace_f(FileContent, "SET BMK_NAME=.*$", "SET BMK_NAME=""" & BMKName & """")
            FileContent = reg.Replace_f(FileContent, "SET COMPONENT=.*$", "SET COMPONENT=""" & UCase(component) & """")

            TS.Close()

            TS = fso.OpenTextFile(run_Cantata_execution_file, IOMode.ForWriting, True)
            TS.Write(FileContent)
            TS.Close()

            ' Create shell object

            Dim cmd As String = "cmd /k cd /d " & listinfo(Position.TOOL_FOLDER).GetValue & "\cantata_tool_external\Single_run & " & listinfo(Position.TOOL_FOLDER).GetValue & "\cantata_tool_external\Single_run\run_Cantata_execution.bat"

            Shell(cmd, vbNormalFocus)


            Return Nothing
        End Function

        Overrides Function AdditionCondition() As Boolean
            'Check module test location is valid
            testlocationpath = listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue
            Dim pathObj As New Information.CheckPathExist(testlocationpath)
            Dim isValid As Boolean = pathObj.IsValid()
            additional_errorMsg = pathObj.GetErrorMsg()
            If isValid Then
                'Check Module path is valid
                Dim modulepath As String = listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue + "\" + listinfo(Position.MODULE_PATH).GetValue
                pathObj = New Information.CheckPathExist(modulepath)
                isValid = pathObj.IsValid()
                additional_errorMsg = pathObj.GetErrorMsg()
                If isValid Then
                    'Check build path is valie
                    Dim build_path As String = listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue + "\build"
                    pathObj = New Information.CheckPathExist(build_path)
                    isValid = pathObj.IsValid()
                    additional_errorMsg = pathObj.GetErrorMsg()
                    If isValid Then

                        'Sear .hpp file path in sandbox 
                        Dim files() = Directory.GetFiles(testlocationpath, "cmake_gen.bat", SearchOption.TopDirectoryOnly)
                        If files.Length <> 0 Then
                            'Built with Cmacke
                            BMKName = ""
                        Else
                            'Built with makefile -> ' Search BMK
                            If Directory.GetDirectories(build_path).Length > 0 Then
                                BMKName = Directory.GetDirectories(build_path)(0).ToString
                                BMKName = Directory.GetDirectories(build_path)(0).ToString.Replace(build_path + "\", String.Empty)
                            Else
                                additional_errorMsg = "There is no 'BMK Name' for this task " & vbNewLine & "Please go to " & listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue + "\project\sw\build for check"
                                isValid = False
                            End If
                        End If


                    End If
                End If
            End If

            'Check flist_file is valid
            flist_file = listinfo(Position.TOOL_FOLDER).GetValue & "\cantata_tool_external\Template_folder\flist_file.txt"
            pathObj = New Information.CheckPathExist(flist_file)
            isValid = isValid AndAlso pathObj.IsValid()
            additional_errorMsg = additional_errorMsg & vbNewLine & pathObj.GetErrorMsg()

            'Check cantata auto execution is valid
            run_Cantata_execution_file = listinfo(Position.TOOL_FOLDER).GetValue & "\cantata_tool_external\Single_run\run_Cantata_execution.bat"
            pathObj = New Information.CheckPathExist(run_Cantata_execution_file)
            isValid = isValid AndAlso pathObj.IsValid()
            additional_errorMsg = additional_errorMsg & vbNewLine & pathObj.GetErrorMsg()
            Return isValid
        End Function
    End Class

End Namespace
