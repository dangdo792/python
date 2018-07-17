Imports System.IO
Imports EUOEM_Tracker.Information
Imports Newtonsoft.Json.Linq

Namespace ReadJson
    'Read json file base
    Public MustInherit Class ReadJsonBase
        Public listinfo() As Information.InfoBase

        Public IsValid As Boolean = True
        Public ErrorMsg As String = Nothing


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
                Dim json As JObject = JObject.Parse(jsonfile)
                For i = 0 To json.SelectToken("field").Count - 1
                    If CheckCondition(json, i) Then
                        Return Getvalue(json.SelectToken("field").SelectToken("[" & i & "]"))
                    End If
                Next
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
                Return listinfo(Position.JSON_PATH).GetValue & "Review_Config.json"
            Else
                Return listinfo(Position.JSON_PATH).GetValue & "\Review_Config.json"
            End If
        End Function

        Public Overrides Function CheckCondition(json_in As Object, index As Integer) As Boolean
            Return (json_in.SelectToken("field").SelectToken("[" & index & "]").SelectToken("Reviewer").ToString = Trim(LCase(listinfo(Position.REVIEWER).GetValue)))
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
            MODEL
            JSON_PATH
        End Enum

        Public Sub New(project_in As CheckNull, model_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New({project_in, model_in, jsonpath_in})
        End Sub

        Public Overrides Function GetJsonFile() As String
            If Right(listinfo(Position.JSON_PATH).GetValue, 1) = "\" Then
                Return listinfo(Position.JSON_PATH).GetValue & "Project_Config.json"
            Else
                Return listinfo(Position.JSON_PATH).GetValue & "\Project_Config.json"
            End If
        End Function

        Public Overrides Function CheckCondition(json_in As Object, index As Integer) As Boolean
            Return (json_in.SelectToken("field").SelectToken("[" & index & "]").SelectToken("Project_Model").ToString = listinfo(Position.PROJECT).GetValue & "_" & listinfo(Position.MODEL).GetValue)
        End Function

        Public Overrides Function Getvalue(node_in As Object) As String
            Return Nothing
        End Function
    End Class

    Public Class ReadJsonSysType
        Inherits ReadJsonProjectConfigFile

        Public Sub New(project_in As CheckNull, model_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New(project_in, model_in, jsonpath_in)
        End Sub

        Public Overrides Function Getvalue(node_in As Object) As String
            Return node_in.SelectToken("System_Type")
        End Function
    End Class

    Public Class ReadJsonProjectCoor
        Inherits ReadJsonProjectConfigFile

        Public Sub New(project_in As CheckNull, model_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New(project_in, model_in, jsonpath_in)
        End Sub

        Public Overrides Function Getvalue(node_in As Object) As String
            Return node_in.SelectToken("Task_Coor")
        End Function
    End Class

    Public Class ReadJsonResultPath
        Inherits ReadJsonProjectConfigFile

        Public Sub New(project_in As CheckNull, model_in As CheckNull, jsonpath_in As CheckNull)
            MyBase.New(project_in, model_in, jsonpath_in)
        End Sub

        Public Overrides Function Getvalue(node_in As Object) As String
            Return node_in.SelectToken("Result_Path")
        End Function
    End Class
End Namespace


