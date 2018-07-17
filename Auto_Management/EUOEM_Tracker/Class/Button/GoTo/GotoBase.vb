Namespace Button
    MustInherit Class GotoBase
        ' The EXPLORER information must be put at 0 index of the array
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
        End Enum

        Public Sub New(listinfo_in() As Information.InfoBase)
            MyBase.New(listinfo_in)
        End Sub

        Public Overrides Function DoFunctionality() As String
            Call Shell(listinfo(Position.EXPLORER).GetValue & " " & GetFullPath(), vbNormalFocus)
            Return Nothing
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim IsValid As Boolean = True
            additional_errorMsg = Nothing
            Dim path As String = GetFullPath()

            additional_errorMsg = IsFileExist(listinfo(Position.EXPLORER).GetValue)
            If Not String.IsNullOrEmpty(additional_errorMsg) Then
                additional_errorMsg = "- Explorer isn't existed. Please check" & vbNewLine & additional_errorMsg & vbNewLine & vbNewLine
                IsValid = False
            End If

            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            If Not fso.FolderExists(path) Then
                additional_errorMsg = additional_errorMsg & "- Folder doesn't exist. Please check: " & vbNewLine & path
                IsValid = False
            End If

            Return IsValid
        End Function

        ''' <summary>
        ''' Get full path to go to 
        ''' </summary>
        ''' <returns>string of full path</returns>
        MustOverride Function GetFullPath() As String


        Public Function IsFileExist(ByVal FileDir As String) As String
            Dim ErrorMsg As String = Nothing
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            If Not fso.FileExists(FileDir) Then
                ErrorMsg = FileDir
            End If
            Return ErrorMsg
        End Function
    End Class

End Namespace
