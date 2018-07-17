Imports EUOEM_Tracker.Information
Imports Newtonsoft.Json
Imports Scripting

Namespace WriteJson

    Class WriteJsonBase
        Public listinfo() As InfoBase
        Public ErrorMsg As String
        Public rootobj As RootObject
        Public jsonFilePath As String
        Public FileName As String

        Private Enum Position
            JSON_PATH
        End Enum

        Public Sub New(listinfo_in() As InfoBase, FileName_in As String)
            listinfo = listinfo_in
            FileName = FileName_in
        End Sub

        Public Function execute()
            Dim LinkObject As CheckPathExist
            Dim IsValid As Boolean = True
            jsonFilePath = GetFilePath()
            LinkObject = New CheckPathExist(jsonFilePath)
            IsValid = LinkObject.IsValid()
            ErrorMsg = LinkObject.GetErrorMsg()
            If IsValid Then
                Dim dataResults
                Dim content As String = IO.File.ReadAllText(jsonFilePath)
                rootobj = JsonConvert.DeserializeObject(content, GetType(RootObject))
                dataResults = Dofunctionality()
                WriteJsonFile()
                Return dataResults
                Exit Function
            End If
            Return Nothing
        End Function

        Overridable Function GetFilePath() As String
            Return listinfo(Position.JSON_PATH).GetValue & "\" & FileName
        End Function

        Overridable Function Dofunctionality() As Object
            Return Nothing
        End Function

        Overridable Sub WriteJsonFile()
            Dim serializerSettings As New JsonSerializerSettings
            serializerSettings.NullValueHandling = NullValueHandling.Ignore
            Dim jsonData As String = JsonConvert.SerializeObject(rootobj, serializerSettings)

            Dim fso As FileSystemObject
            Dim TS As TextStream
            fso = New FileSystemObject
            TS = fso.OpenTextFile(jsonFilePath, IOMode.ForWriting, True)
            TS.Write(jsonData)
            TS.Close()
        End Sub
    End Class


    Class LoadJson
        Inherits WriteJsonBase

        Public Sub New(jsonpath_in As CheckNull, FileName_in As String)
            MyBase.New({jsonpath_in}, FileName_in)
        End Sub

        Overrides Function Dofunctionality()
            If FileName = "Project_Config.json" Then
                Return rootobj.fields.Select(Function(data) New fields With {.Project_Model = data.Project_Model,
                                                                   .Task_Coor = data.Task_Coor,
                                                                   .Result_Path = data.Result_Path,
                                                                   .System_Type = data.System_Type}).ToList
            ElseIf FileName = "Review_Config.json" Then
                Return rootobj.fieldsReview.Select(Function(data) New fieldsReview With {.Reviewer = data.Reviewer,
                                                                                        .Reviewer_ID = data.Reviewer_ID}).ToList
            Else
                Return Nothing
            End If
        End Function

        Overrides Sub WriteJsonFile()

        End Sub
    End Class

    Class AddJson
        Inherits WriteJsonBase

        Private ProMod As String
        Private ProCoor As String
        Private ProResultPath As String
        Private SysType As String
        Private Reviewer As String
        Private Reviewer_ID As String

        Public Sub New(jsonpath_in As CheckNull,
                        FileName_in As String,
                       ProMod_in As String,
                       ProCoor_in As String,
                       ProResultPath_in As String)
            MyBase.New({jsonpath_in}, FileName_in)
            ProMod = ProMod_in
            ProCoor = ProCoor_in
            ProResultPath = ProResultPath_in
        End Sub

        Overrides Function Dofunctionality()
            If FileName = "Project_Config.json" Then
                rootobj.fields.Add(New fields() With {.Project_Model = ProMod,
                                                    .Task_Coor = ProCoor,
                                                    .Result_Path = ProResultPath,
                                                    .System_Type = SysType})
            ElseIf FileName = "Review_Config.json" Then
                rootobj.fieldsReview.Add(New fieldsReview() With {.Reviewer = Reviewer,
                                                                .Reviewer_ID = Reviewer_ID})
            End If
            Return Nothing
        End Function
    End Class

    Class RemoveJson
        Inherits WriteJsonBase

        Dim drvctl As MetroFramework.Controls.MetroGrid

        Public Sub New(jsonpath_in As CheckNull,
                       FileName_in As String,
                       drvctl_in As MetroFramework.Controls.MetroGrid)
            MyBase.New({jsonpath_in}, FileName_in)
            drvctl = drvctl_in
        End Sub

        Overrides Function Dofunctionality()
            Dim RowsSelected() As DataGridViewRow = Nothing
            Dim SelRowIndex As Integer = 0
            For Each EachSelRow As DataGridViewRow In drvctl.SelectedRows
                ReDim Preserve RowsSelected(drvctl.SelectedRows.Count - 1)
                RowsSelected(SelRowIndex) = EachSelRow
                SelRowIndex = SelRowIndex + 1
            Next
            If RowsSelected Is Nothing Then
                Return Nothing
                Exit Function
            End If
            For Each r As DataGridViewRow In RowsSelected
                If FileName = "Project_Config.json" Then
                    Dim findindex = FindFirst(r.Cells("Project_Model").Value.ToString)
                    rootobj.fields.RemoveAt(findindex)
                ElseIf FileName = "Review_Config.json" Then
                    Dim findindex = FindFirst(r.Cells("Reviewer").Value.ToString)
                    rootobj.fieldsReview.RemoveAt(findindex)
                End If


            Next
            Return Nothing
        End Function

        Function FindFirst(SearchString As String) As Integer

            If FileName = "Project_Config.json" Then
                Dim StringList As List(Of fields) = rootobj.fields
                Dim I As Integer
                For I = 0 To StringList.Count - 1
                    If StringList(I).Project_Model.Contains(SearchString) Then Exit For
                Next
                Return If(I < StringList.Count, I, -1)
            ElseIf FileName = "Review_Config.json" Then
                Dim StringList As List(Of fieldsReview) = rootobj.fieldsReview
                Dim I As Integer
                For I = 0 To StringList.Count - 1
                    If StringList(I).Reviewer.Contains(SearchString) Then Exit For
                Next
                Return If(I < StringList.Count, I, -1)
            End If
            Return Nothing
        End Function


    End Class



    Public Class fields
        Public Property Project_Model() As String
        Public Property Task_Coor() As String
        Public Property Result_Path() As String
        Public Property System_Type() As String
    End Class
    Public Class fieldsReview
        Public Property Reviewer() As String
        Public Property Reviewer_ID() As String
    End Class
    Public Class RootObject
        <JsonProperty(PropertyName:="fields")>
        Public Property fields() As List(Of fields)
        <JsonProperty(PropertyName:="fieldsReview")>
        Public Property fieldsReview() As List(Of fieldsReview)
    End Class
End Namespace
