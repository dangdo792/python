Namespace Button
    Public MustInherit Class AutoGenBase
        Inherits ButtonBase

        ''' <summary>
        ''' The position of first 3 field is a must, do not adapt it
        ''' </summary>
        Private Enum Position
            TOOL_FOLDER
            MODULE_PATH
            TASK_ID
        End Enum

        Public Sub New(listinfo_in() As Information.InfoBase)
            MyBase.New(listinfo_in)
        End Sub

        Private Auto_Gen_dir As String
        Private Testcase_Design_Path As String

        Public cmdoutput As String

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing
            Dim M_Type As String = Nothing
            If Right(listinfo(Position.MODULE_PATH).GetValue, 4) = ".cpp" Then : M_Type = "cpp"
            ElseIf Right(listinfo(Position.MODULE_PATH).GetValue, 2) = ".c" Then : M_Type = "c"
            End If
            Dim oShell
            oShell = CreateObject("WScript.Shell")
            oShell.Run("cmd /k cd /d " + listinfo(Position.TOOL_FOLDER).GetValue + "\Code & perl " + Auto_Gen_dir + " -i " + Testcase_Design_Path + " -t " + M_Type + " & exit", 0, True)

            'Process to get output cmd
            Dim proc As New Process
            proc.StartInfo.FileName = "cmd.exe"
            proc.StartInfo.Arguments = "cmd /k cd /d " + listinfo(Position.TOOL_FOLDER).GetValue + "\Code & perl " + Auto_Gen_dir + " -i " + Testcase_Design_Path + " -t " + M_Type + " & exit"
            proc.StartInfo.CreateNoWindow = True
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.Start()
            proc.WaitForExit()

            cmdoutput = proc.StandardOutput.ReadToEnd

            oShell = Nothing
            Return ErrorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim LinkObject As Information.CheckPathExist
            Dim IsValid As Boolean = True
            'Check Explorer is exist or not
            Auto_Gen_dir = listinfo(Position.TOOL_FOLDER).GetValue & "\Code\AutoGen.pl"
            LinkObject = New Information.CheckPathExist(Auto_Gen_dir)
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

        ''' <summary>
        ''' Get full path of testcase file
        ''' </summary>
        ''' <returns>String full testcase path</returns>
        Public MustOverride Function GetFullPath() As String
    End Class
End Namespace
