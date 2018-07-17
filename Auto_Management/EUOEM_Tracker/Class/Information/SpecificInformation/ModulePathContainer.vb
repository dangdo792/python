Namespace Information
    ''' <summary>
    ''' Module path information container class
    ''' </summary>
    Public Class ModulePath
        Inherits InfoBase

        Public Sub New(infovalue As String)
            MyBase.New(infovalue)
        End Sub

        ''' <summary>
        ''' Override function get valid status of information
        ''' </summary>
        ''' <returns>True if information is not null and end with ".c" or ".cpp", False otherwise</returns>
        Public Overrides Function IsValid() As Boolean
            If info_value = "" Then
                errorMsg = "Module path is empty"
                IsValid = False
            ElseIf LCase(Right(info_value, 2)) <> ".c" And LCase(Right(info_value, 4)) <> ".cpp" And LCase(Right(info_value, 4)) <> ".inl" Then
                errorMsg = "Type of source code is incorrect. It should be .c or .cpp "
                IsValid = False
            Else
                errorMsg = Nothing
                IsValid = True
            End If
        End Function
    End Class

End Namespace
