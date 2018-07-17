Imports System.Text.RegularExpressions
Imports EUOEM_Tracker.Information
Imports Newtonsoft.Json.Linq
Imports Scripting


Namespace Button
    Public MustInherit Class CreateJIRABase
        Inherits ButtonBase

        Public JIRATicketID As String

        Private updateJirafile As String

        Protected project As String
        Protected issuetype As String
        Protected assignee As String
        Protected labels As String
        Protected components As String
        Protected moduleowner As String
        Protected fullmodulename As String

        Private Enum Position
            TASK_ID
            USER
            PASSWORD
            MODULE_NAME
            TOOL_PATH
            MODULE_PATH
        End Enum

        Public Sub New(listinfo_in() As Information.InfoBase)
            MyBase.New(listinfo_in)
        End Sub

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing

            Dim taskid As String = listinfo(Position.TASK_ID).GetValue
            Dim username As String = listinfo(Position.USER).GetValue
            Dim password As String = listinfo(Position.PASSWORD).GetValue

            Dim strURLJ As String : strURLJ = "https://rb-tracker.bosch.com/tracker08/rest/api/2/issue/" & taskid
            Dim winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
            winHttp.SetAutoLogonPolicy(0)
            winHttp.Open("GET", strURLJ)

            Try
                winHttp.Send
            Catch ex As Exception
                Return "Can't access to JIRA" & vbCrLf
                Exit Function
            End Try

            Dim success As Boolean
            success = winHttp.waitForResponse(5)
            If Not success Then Return "Can't access to JIRA. Please check access right" : Exit Function

            Dim json As JObject
            Try
                json = JObject.Parse(winHttp.responseText)
            Catch ex As Exception
                Return "Can't access to JIRA. Please check access right" & vbCrLf
                Exit Function
            End Try

            Dim jsonToken As JToken = json.SelectToken("fields")
            project = jsonToken.SelectToken("project").SelectToken("key").ToString
            issuetype = jsonToken.SelectToken("issuetype").SelectToken("name")
            assignee = jsonToken.SelectToken("assignee").SelectToken("name")
            labels = jsonToken.SelectToken("labels").SelectToken("[0]")
            components = jsonToken.SelectToken("components").SelectToken("[0]").SelectToken("name")
            moduleowner = jsonToken.SelectToken("reporter").SelectToken("displayName")

            Dim aPath() As String
            aPath = Split(listinfo(Position.MODULE_PATH).GetValue, "\")
            fullmodulename = aPath(UBound(aPath))

            'Create file template
            Dim fso As New FileSystemObject
            Dim TS As TextStream = fso.OpenTextFile(updateJirafile, IOMode.ForWriting, True)
            Dim Final As String = GetJsonData()



            'Dim listlines(12) As String
            'listlines(0) = "{" & vbNewLine
            'listlines(1) = Chr(34) & "fields" & Chr(34) & ": " & "{" & vbNewLine
            'listlines(2) = Chr(34) & "project" & Chr(34) & ": {" & Chr(34) & "key" & Chr(34) & ": " & Chr(34) & project & Chr(34) & "}," & vbNewLine
            'listlines(3) = Chr(34) & "issuetype" & Chr(34) & ": {" & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & "Task" & Chr(34) & "}," & vbNewLine
            'listlines(4) = Chr(34) & "summary" & Chr(34) & ":" & Chr(34) & "TCDS&TS Review for " & fullmodulename & Chr(34) & "," & vbNewLine
            'listlines(5) = Chr(34) & "assignee" & Chr(34) & ": {" & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & assignee & Chr(34) & "}," & vbNewLine
            'listlines(6) = Chr(34) & "labels" & Chr(34) & ": [" & Chr(34) & labels & "_Review" & Chr(34) & "]," & vbNewLine
            'listlines(7) = Chr(34) & "components" & Chr(34) & ":[{" & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & components & Chr(34) & "}]," & vbNewLine
            'listlines(8) = Chr(34) & "description" & Chr(34) & ":" & Chr(34) & "Please help me review this task.\n ILM link: \n" & Chr(34) & "}}" & vbNewLine

            'Dim Final1 = Join(listlines, vbCrLf)

            TS.Write(Final)
            TS.Close()


            Dim debugabc As String = "cd " & listinfo(Position.TOOL_PATH).GetValue & "\CQAJira"

            'Call cmd
            Dim proc As New System.Diagnostics.Process()
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.FileName = listinfo(Position.TOOL_PATH).GetValue & "\CQAJira\run.bat"
            proc.StartInfo.WorkingDirectory = listinfo(Position.TOOL_PATH).GetValue & "\CQAJira"
            proc.StartInfo.Arguments = "JIRA POST"
            proc.StartInfo.CreateNoWindow = True
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
            proc.Start()



            Dim ProcOutput As String = proc.StandardOutput.ReadToEnd

            Dim regex As Regex = New Regex("(?<=key"":"").*(?="","")")
            Dim match As Match = regex.Match(ProcOutput)
            If match.Success Then
                JIRATicketID = match.Value
            End If

            proc.Close()

            Return ErrorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            'Check UploadJson File is exist
            updateJirafile = listinfo(Position.TOOL_PATH).GetValue & "\CQAJira\updateJira.json"
            Dim pathObj As New Information.CheckPathExist(updateJirafile)
            Dim isValid As Boolean = pathObj.IsValid()
            additional_errorMsg = pathObj.GetErrorMsg()
            If isValid Then
                'Check run.bat file is exist
                Dim runbatfile As String = listinfo(Position.TOOL_PATH).GetValue & "\CQAJira\run.bat"
                pathObj = New Information.CheckPathExist(runbatfile)
                isValid = pathObj.IsValid()
                additional_errorMsg = pathObj.GetErrorMsg()
            End If
            Return isValid
        End Function

        MustOverride Function GetJsonData() As String

    End Class
End Namespace