Imports System.IO
Imports Scripting

Namespace Button

    Public Class GetHashCommit
        Inherits ButtonBase

        Private Enum Position
            SANDBOX_FOLDER
            SANDBOX
            MODULE_PATH
        End Enum

        Public Sub New(sfd_in As Information.CheckNull,
                       sandbox_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath)
            MyBase.New({sfd_in,
                       sandbox_in,
                       modulepath_in})
        End Sub

        Private testlocationpath As String

        Public commit As String
        Public hash As String

        Public SpecSourceFlag As Boolean
        Public SpecSourcePath As String

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing
            Dim modulepath As String

            If SpecSourceFlag = True Then
                modulepath = SpecSourcePath
            Else
                modulepath = listinfo(Position.MODULE_PATH).GetValue()
            End If

            'Check Module path is valid
            Dim gitpath As String = listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue + "\.git"
            Dim pathObj As New Information.CheckPathExist(gitpath)
            Dim IsValid As Boolean = pathObj.IsValid()
            additional_errorMsg = pathObj.GetErrorMsg() & vbNewLine & vbNewLine

            'Check is full sandbox
            If IsValid Then
                Dim proc As New Process
                proc.StartInfo.FileName = "cmd.exe"
                proc.StartInfo.Arguments = "cmd /k cd /d " + testlocationpath + "&git rev-parse HEAD"
                proc.StartInfo.CreateNoWindow = True
                proc.StartInfo.UseShellExecute = False
                proc.StartInfo.RedirectStandardOutput = True
                proc.Start()
                proc.WaitForExit()
                'Get commit of repo (<=> Sandbox)
                commit = proc.StandardOutput.ReadLine
                If commit IsNot Nothing Then
                    Dim abc = Mid(modulepath, 2, Len(modulepath)).Replace("\", "/")
                    proc.StartInfo.Arguments = "cmd /k cd /d " + testlocationpath + "&git rev-parse " & commit & ":" & Mid(modulepath, 2, Len(modulepath)).Replace("\", "/")
                    proc.Start()
                    proc.WaitForExit()
                    'Get hash of file (<=> Revision)
                    hash = proc.StandardOutput.ReadLine
                End If
                additional_errorMsg = Nothing
            Else 'Not full sandbox
                Dim revisiontextpath = listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue + "\revision.txt"
                pathObj = New Information.CheckPathExist(revisiontextpath)
                IsValid = pathObj.IsValid()
                additional_errorMsg = additional_errorMsg & pathObj.GetErrorMsg()
                'check is it a built sandbox
                If IsValid Then
                    Dim debugabc = Mid(modulepath, 2, Len(modulepath)).Replace("\", "/")
                    Dim fso As FileSystemObject
                    Dim TS As TextStream
                    Dim Final As String = Nothing
                    Dim listlines() As String
                    fso = New FileSystemObject
                    TS = fso.OpenTextFile(revisiontextpath, IOMode.ForReading)
                    listlines = Split(TS.ReadAll, vbCrLf)
                    If listlines(0) IsNot Nothing Then commit = listlines(0).Replace("commit", "").Trim
                    For Each line In listlines
                        If InStr(line, Mid(modulepath, 2, Len(modulepath)).Replace("\", "/")) <> 0 Then
                            hash = line.Replace(Mid(modulepath, 2, Len(modulepath)).Replace("\", "/"), "").Trim
                        End If
                    Next
                    additional_errorMsg = Nothing
                    TS.Close()
                    fso = Nothing
                    GC.Collect()
                End If
            End If
            Return additional_errorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            ''Check module test location is valid
            testlocationpath = listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue
            Dim pathObj As New Information.CheckPathExist(testlocationpath)
            Dim isValid As Boolean = pathObj.IsValid()
            additional_errorMsg = pathObj.GetErrorMsg()
            If isValid Then
                Dim sourcecodepath = listinfo(Position.SANDBOX_FOLDER).GetValue + "\" + listinfo(Position.SANDBOX).GetValue + listinfo(Position.MODULE_PATH).GetValue
                pathObj = New Information.CheckPathExist(sourcecodepath)
                isValid = pathObj.IsValid()
                additional_errorMsg = pathObj.GetErrorMsg()
            End If
            Return isValid
        End Function
    End Class
End Namespace
