Imports EUOEM_Tracker.Information

Namespace Button
    Public Class CreateJIRAPDC
        Inherits CreateJIRABase

        Private Enum Position
            TASK_ID
            USER
            PASSWORD
            MODULE_NAME
            TOOL_PATH
            MODULE_PATH
            ILM_LINK
        End Enum

        Public Sub New(taskid_in As TaskID,
                       user_in As CheckNull,
                       password_in As CheckNull,
                       modulename_in As CheckNull,
                       toolpath_in As CheckNull,
                       modulepath_in As ModulePath,
                       ilmlink_in As CheckNull)
            MyBase.New({taskid_in, user_in, password_in, modulename_in, toolpath_in, modulepath_in, ilmlink_in})
        End Sub

        Public Overrides Function GetJsonData() As String
            Dim jsoncontent As String = Nothing

            jsoncontent = "{
                            ""fields"": {
                            ""project"":{""key"": """ & project & """},
                            ""issuetype"": {""name"": ""Task""},
                            ""summary"": ""PDC for " & fullmodulename & """,
                            ""assignee"": {""name"": """ & assignee & """},
                            ""labels"":[""" & labels & "_Review""],
                            ""components"":[{""name"":""" & components & """}],
                            ""description"": ""Please help me PDC this task.\nILM link: [" & listinfo(Position.ILM_LINK).GetValue & "]""}}"

            Return jsoncontent
        End Function
    End Class
End Namespace