Imports System.IO
Imports EUOEM_Tracker.Information
Imports Newtonsoft.Json.Linq

Namespace ReadJson
    'Read json file base
    Public MustInherit Class ReadJsonBase
        Public listinfo() As Information.InfoBase

        Public IsValid As Boolean = True
        Public ErrorMsg As String = Nothing

        Public JsonFileName As String


        Public Sub New(listinfo_in() As Information.InfoBase)
            listinfo = listinfo_in
        End Sub

        Public Function Execute() As String
            Dim LinkObject As Information.CheckPathExist
            LinkObject = New Information.CheckPathExist(GetJsonFile)
            IsValid = LinkObject.IsValid()
            ErrorMsg = LinkObject.GetErrorMsg() & vbNewLine
            If IsValid = True Then
                Dim jsonfile As String = File.ReadAllText(GetJsonFile)
                If jsonfile = "" Then
                    Return Nothing
                    Exit Function
                End If
                If InStr(GetJsonFile, "Project_Config.json") Then
                    Dim json As JObject = JObject.Parse(jsonfile)
                    Dim countelementjson As Integer = json.SelectToken("fields").Count
                    If countelementjson <> 0 Then
                        For i = 0 To countelementjson - 1
                            If CheckCondition(json, i) Then
                                Return Getvalue(json.SelectToken("fields").SelectToken("[" & i & "]"))
                            End If
                        Next
                    End If
                ElseIf InStr(GetJsonFile, "Review_Config.json") Then
                    Dim json As JObject = JObject.Parse(jsonfile)
                    Dim countelementjson As Integer = json.SelectToken("fieldsReview").Count
                    If countelementjson <> 0 Then
                        For i = 0 To countelementjson - 1
                            If CheckCondition(json, i) Then
                                Return Getvalue(json.SelectToken("fieldsReview").SelectToken("[" & i & "]"))
                            End If
                        Next
                    End If
                End If
            End If

                Return Nothing
        End Function

        MustOverride Function GetJsonFile() As String
        MustOverride Function CheckCondition(ByVal json_in As Object, ByVal index As Integer) As Boolean
        MustOverride Function Getvalue(ByVal node_in As Object) As String
    End Class
    'Read Json reviewer_config file
    Public Class ReadJsonReviewerFile
        Inherits ReadJsonBase

        Private Enum Position
            REVIEWER
            JSON_PATH
        End Enum

        Public Sub New(reviewer_in As NoCheck, jsonpath_in As CheckNull)
            MyBase.New({reviewer_in, jsonpath_in})
        End Sub

        Public Overrides Function GetJsonFile() As String
            If Right(listinfo(Position.JSON_PATH).GetValue, 1) = "\" Then
                JsonFileName = listinfo(Position.JSON_PATH).GetValue & "Review_Config.json"
                Return JsonFileName
            Else
                JsonFileName = listinfo(Position.JSON_PATH).GetValue & "\Review_Config.json"
                Return JsonFileName
            End If
        End Function

        Public Overrides Function CheckCondition(json_in As Object, index As Integer) As Boolean
            Return (json_in.SelectToken("fieldsReview").SelectToken("[" & index & "]").SelectToken("Reviewer").ToString = Trim(LCase(listinfo(Position.REVIEWER).GetValue)))
        End Function

        Public Overrides Function Getvalue(node_in As Object) As String
            Return node_in.SelectToken("Reviewer_ID")
        End Function
    End Class
    'Read Json project_config file
    Public Class ReadJsonProjectConfigFile
        Inherits ReadJsonBase

        Private Enum Position
            PROJECT
            JSON_PATH
        End Enum

        Public Sub New(project_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New({project_in, jsonpath_in})
        End Sub

        Public Overrides Function GetJsonFile() As String
            If Right(listinfo(Position.JSON_PATH).GetValue, 1) = "\" Then
                JsonFileName = listinfo(Position.JSON_PATH).GetValue & "Project_Config.json"
                Return JsonFileName
            Else
                JsonFileName = listinfo(Position.JSON_PATH).GetValue & "\Project_Config.json"
                Return JsonFileName
            End If
        End Function

        Public Overrides Function CheckCondition(json_in As Object, index As Integer) As Boolean
            Return (json_in.SelectToken("fields").SelectToken("[" & index & "]").SelectToken("Project_Model").ToString = listinfo(Position.PROJECT).GetValue)
        End Function

        Public Overrides Function Getvalue(node_in As Object) As String
            Return Nothing
        End Function
    End Class

    Public Class ReadJsonSysType
        Inherits ReadJsonProjectConfigFile

        Public Sub New(project_in As CheckNull, model_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New(project_in, jsonpath_in)
        End Sub

        Public Overrides Function Getvalue(node_in As Object) As String
            Return node_in.SelectToken("System_Type")
        End Function
    End Class

    Public Class ReadJsonProjectCoor
        Inherits ReadJsonProjectConfigFile

        Public Sub New(project_in As CheckNull, model_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New(project_in, jsonpath_in)
        End Sub

        Public Overrides Function Getvalue(node_in As Object) As String
            Return node_in.SelectToken("Task_Coor")
        End Function
    End Class

    Public Class ReadJsonResultPath
        Inherits ReadJsonProjectConfigFile

        Public Sub New(project_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New(project_in, jsonpath_in)
        End Sub

        Public Overrides Function Getvalue(node_in As Object) As String
            Return node_in.SelectToken("Result_Path")
        End Function
    End Class
End Namespace


