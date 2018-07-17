Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

Namespace Button
    Public Class GetCQorJIRA
        Inherits ButtonBase

        Public project As String = Nothing
        Public modulename As String = Nothing
        Public modulepath As String = Nothing
        Public moduleowner As String = Nothing
        Public taskauthor As String = Nothing

        Private Enum Position
            TASK_ID
        End Enum

        Public Sub New(taskid_in As Information.TaskID)
            MyBase.New({taskid_in})
        End Sub

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing
            Dim taskid As String = listinfo(Position.TASK_ID).GetValue

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

            Dim json As JObject
            Try
                json = JObject.Parse(winHttp.responseText)
            Catch ex As Exception
                Return "Can't access to JIRA. Please check access right" & vbCrLf
                Exit Function
            End Try

            Dim jsonToken As JToken = json.SelectToken("fields")
            project = jsonToken.SelectToken("project").SelectToken("key").ToString
            moduleowner = jsonToken.SelectToken("reporter").SelectToken("displayName")
            taskauthor = jsonToken.SelectToken("assignee").SelectToken("displayName")

            Dim regex As Regex = New Regex("^.*(?=[(])")
            Dim match As Match = regex.Match(moduleowner)
            If match.Success Then
                moduleowner = match.Value
            End If
            If Not String.IsNullOrEmpty(moduleowner) Then moduleowner.Trim()


            regex = New Regex("^.*(?=[(])")
            match = regex.Match(taskauthor)
            If match.Success Then
                taskauthor = match.Value
            End If
            If Not String.IsNullOrEmpty(taskauthor) Then taskauthor.Trim()


            regex = New Regex("\w+[/].*(?=[|])")
            match = regex.Match(jsonToken.SelectToken("description"))
            If match.Success Then
                modulepath = match.Value
                modulepath = modulepath.Replace("/", "\")
            Else
                regex = New Regex("\w+\\.*(?=[|])")
                match = regex.Match(jsonToken.SelectToken("description"))
                If match.Success Then
                    modulepath = match.Value
                End If
            End If
            modulepath = "\" & modulepath

            regex = New Regex("\\([^\\]*)(?=[.])")
            match = regex.Match(modulepath)
            If match.Success Then
                modulename = Mid(match.Value, 2)
            End If

            Return ErrorMsg
        End Function
    End Class

End Namespace
