Imports EUOEM_Tracker.Information

Namespace Button
    Public Class CreateJIRAOPL
        Inherits CreateJIRABase

        Private Enum Position
            TASK_ID
            USER
            PASSWORD
            MODULE_NAME
            TOOL_PATH
            MODULE_PATH
            OPL_LINK
        End Enum

        Public Sub New(taskid_in As TaskID,
                       user_in As CheckNull,
                       password_in As CheckNull,
                       modulename_in As CheckNull,
                       toolpath_in As CheckNull,
                       modulepath_in As ModulePath,
                       opllink_in As CheckNull)
            MyBase.New({taskid_in, user_in, password_in, modulename_in, toolpath_in, modulepath_in, opllink_in})
        End Sub

        Public Overrides Function GetJsonData() As String
            Dim jsoncontent As String = Nothing

            Dim parent As String = Nothing
            Dim subtask As String = Nothing
            If InStr(LCase(components), "sit") Then
                Dim parenttask = InputBox("This task is from Sit component." & vbNewLine & "Please Put parent to create subtask: ", , "PJDC-15414")
                parent = """parent"": {""key"": """ & parenttask & """},"
                subtask = ",""subtask"": ""true"""
            Else

            End If

            jsoncontent = "{
                            ""fields"": {
                            ""project"":{""key"": """ & project & """},
                            " & parent & "
                            ""issuetype"": {""name"": """ & issuetype & """" & subtask & "},
                            ""summary"": ""SW_UVE - OPL for " & fullmodulename & """,
                            ""assignee"": {""name"": """ & assignee & """},
                            ""labels"":[""" & labels & "_OPL""],
                            ""components"":[{""name"":""" & components & """}],
                            ""description"": ""Hello " & moduleowner & ",\n\n Please find OPL for module " & listinfo(Position.MODULE_NAME).GetValue & " in below location :\n" & listinfo(Position.OPL_LINK).GetValue.Replace("\", "\\") & """}}"

            Return jsoncontent
        End Function
    End Class
End Namespace
