Imports EUOEM_Tracker.Information

Namespace Button
    Public Class CreateJIRADefect
        Inherits CreateJIRABase

        Private Enum Position
            TASK_ID
            USER
            PASSWORD
            MODULE_NAME
            TOOL_PATH
            MODULE_PATH
            HASH
        End Enum

        Public Sub New(taskid_in As TaskID,
                       user_in As CheckNull,
                       password_in As CheckNull,
                       modulename_in As CheckNull,
                       toolpath_in As CheckNull,
                       modulepath_in As ModulePath,
                       hash_in As CheckNull)
            MyBase.New({taskid_in, user_in, password_in, modulename_in, toolpath_in, modulepath_in, hash_in})
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
                            ""summary"": ""[SW_UVE_DEFECT] " & fullmodulename & " (Process: Software Unit Verification)" & """,
                            ""assignee"": {""name"": """ & assignee & """},
                            ""labels"":[""" & labels & "_DEFECT""],
                            ""components"":[{""name"":""" & components & """}],
                            ""description"": ""{panel:title=Defect Information}\n||File_name|" & fullmodulename & "|\n||File_hash|" & listinfo(Position.HASH).GetValue & "|\n||Detailed_description||\n||Observed_result_or_behavior|the incorrect implementation|\n||Expected_result_or_behavior|the correct implementation|\n||Requirement_violated|N.A|\n||Defect_type|Data controlling|\n{panel} " & "\n\n""}}"

            Return jsoncontent
        End Function
    End Class
End Namespace