Namespace Information
    ''' <summary>
    ''' Check if the folder is exist or not
    ''' </summary>
    Public Class CheckPathExist
        Inherits InfoBase

        Public Sub New(infovalue As String)
            MyBase.New(infovalue)
        End Sub

        ''' <summary>
        ''' Override function get valid status of information
        ''' Error message contain field name information
        ''' </summary>
        ''' <returns>True if information is not null, False otherwise</returns>
        Public Overrides Function IsValid() As Boolean
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")

            If Not fso.FolderExists(GetValue()) And Not fso.FileExists(GetValue()) Then
                errorMsg = "Path doesn't exist." & vbNewLine & "Please check: " & vbNewLine & GetValue()
                IsValid = False
            Else
                errorMsg = Nothing
                IsValid = True
            End If
        End Function
    End Class

End Namespace
