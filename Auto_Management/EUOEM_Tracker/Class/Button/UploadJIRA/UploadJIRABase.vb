Imports System.Text.RegularExpressions
Imports EUOEM_Tracker.Information
Imports Newtonsoft.Json.Linq
Imports Scripting
Namespace Button
    Public MustInherit Class UploadJIRABase
        Inherits ButtonBase


        Protected updateJirafile As String

        Protected project As String
        Protected issuetype As String
        Protected assignee As String
        Protected labels As String
        Protected components As String
        Protected moduleowner As String
        Protected fullmodulename As String
        Protected decription As String

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
            decription = jsonToken.SelectToken("description")

            Dim aPath() As String
            aPath = Split(listinfo(Position.MODULE_PATH).GetValue, "\")
            fullmodulename = aPath(UBound(aPath))

            Dim HeaderFlag As Boolean = False
            Dim DescriptionUpload As String = Nothing

            'Write content to UploadJson file to update data
            ErrorMsg = CreateJiraFileContent()

            If String.IsNullOrEmpty(ErrorMsg) Then
                'Call cmd
                Dim proc As New Process()
                proc.StartInfo.UseShellExecute = False
                proc.StartInfo.FileName = "cmd.exe"
                proc.StartInfo.Arguments = "cmd /k cd " + listinfo(Position.TOOL_PATH).GetValue & "\CQAJira & run.bat JIRA PUT " & taskid
                proc.StartInfo.CreateNoWindow = True
                proc.StartInfo.RedirectStandardOutput = True
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal
                proc.Start()
            End If

            Return ErrorMsg
        End Function

        MustOverride Function CreateJiraFileContent() As String


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
    End Class

    Public Class UploadJIRAStartTask
        Inherits UploadJIRABase

        Private Enum Position
            TASK_ID
            USER
            PASSWORD
            MODULE_NAME
            TOOL_PATH
            MODULE_PATH
        End Enum

        Public Sub New(taskid_in As TaskID,
                       user_in As CheckNull,
                       password_in As CheckNull,
                       modulename_in As CheckNull,
                       toolpath_in As CheckNull,
                       modulepath_in As ModulePath)
            MyBase.New({taskid_in, user_in, password_in, modulename_in, toolpath_in, modulepath_in})
        End Sub

        Public Overrides Function CreateJiraFileContent() As String
            Dim ErrorMsg As String = Nothing
            If InStr(decription, "{panel:title=Delivery Information}") = 0 Then
                decription = decription.Replace(vbNewLine, "\n")
                'Write data to file to update content
                Dim fso As New FileSystemObject
                Dim TS As TextStream = fso.OpenTextFile(updateJirafile, IOMode.ForWriting, True)
                Dim Final As String = "{
                                    ""fields"": {
                                    ""description"": """ & decription & "\n{panel:title=Delivery Information}\n*File info*\n||File_Name||File_Hash||Code_Coverage||Tested_ELOC||\n|" & fullmodulename & "| |Function: /| |\n{panel}""}}"
                TS.Write(Final)
                TS.Close()
            Else
                ErrorMsg = "start task content alrealdy exist."
            End If
            Return ErrorMsg
        End Function
    End Class

    Public Class UploadJIRADeliveryTask
        Inherits UploadJIRABase

        Private Enum Position
            TASK_ID
            USER
            PASSWORD
            MODULE_NAME
            TOOL_PATH
            MODULE_PATH
            HASH
            STATEMENT
            DECISIONS
            ELOC
            COMMIT
            RESULT_PATH
        End Enum

        Public Sub New(taskid_in As TaskID,
                       user_in As CheckNull,
                       password_in As CheckNull,
                       modulename_in As CheckNull,
                       toolpath_in As CheckNull,
                       modulepath_in As ModulePath,
                       hash_in As NoCheck,
                       statement_in As NoCheck,
                       decision_in As NoCheck,
                       eloc_in As NoCheck,
                       commit_in As NoCheck,
                       resultpath_in As NoCheck)
            MyBase.New({taskid_in,
                       user_in,
                       password_in,
                       modulename_in,
                       toolpath_in,
                       modulepath_in,
                       hash_in,
                       statement_in,
                       decision_in,
                       eloc_in,
                       commit_in,
                       resultpath_in})
        End Sub

        Public Overrides Function CreateJiraFileContent() As String
            Dim ErrorMsg As String = Nothing


            Dim regex As Regex = New Regex("[{]panel:title=Delivery Information[}]([^~]*)", RegexOptions.Multiline)
            decription = regex.Replace(decription, "")

            Dim statement As String
            Dim decisions As String
            If InStr(listinfo(Position.STATEMENT).GetValue, "100") <> 0 And String.IsNullOrEmpty(listinfo(Position.DECISIONS).GetValue) Then
                statement = "100"
                decisions = "100"
            Else
                statement = listinfo(Position.STATEMENT).GetValue
                decisions = listinfo(Position.DECISIONS).GetValue
            End If


            decription = decription.Replace(vbNewLine, "\n").Replace(vbCr, "\n").Replace(vbLf, "\n")
            decription = decription & "{panel:title=Delivery Information}\n*File info*\n||File_Name||File_Hash||Code_Coverage||Tested_ELOC||\n|" & fullmodulename & "|" & listinfo(Position.HASH).GetValue & "|Function: /, Statement: " & statement & "%, Decision: " & decisions & "%|" & listinfo(Position.ELOC).GetValue & "/" & listinfo(Position.ELOC).GetValue & "|\n*Other info*\n||Commit|" & listinfo(Position.COMMIT).GetValue & "|\n||Requirement_baseline|NA|\n||Link to delivery folder|" & listinfo(Position.RESULT_PATH).GetValue.Replace("\", "\\") & "\n||Link to report|[Reports|" & listinfo(Position.RESULT_PATH).GetValue.Replace("\", "\\") & "\\" & listinfo(Position.TASK_ID).GetValue & "\\Reports\\test_report.html" & "]|\n{panel}"

            decription = decription.Replace(vbNewLine, "\n")
            'Write data to file to update content
            Dim fso As New FileSystemObject
            Dim TS As TextStream = fso.OpenTextFile(updateJirafile, IOMode.ForWriting, True)
            Dim Final As String = "{
                                    ""fields"": {
                                    ""description"": """ & decription & """}}"
            TS.Write(Final)
            TS.Close()

            Return ErrorMsg
        End Function
    End Class

End Namespace

