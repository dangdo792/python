Namespace Information
    ''' <summary>
    ''' Class for check if information is null or not
    ''' </summary>
    Public Class CheckNull
        Inherits InfoBase

        ''' <summary>
        ''' field_name: Contain the name of field, use to generate error message
        ''' </summary>
        Private field_name As String

        Public Sub New(field As String, infovalue As String)
            MyBase.New(infovalue)
            field_name = field
        End Sub

        ''' <summary>
        ''' Override function get valid status of information
        ''' Error message contain field name information
        ''' </summary>
        ''' <returns>True if information is not null, False otherwise</returns>
        Public Overrides Function IsValid() As Boolean
            If info_value = "" Then
                errorMsg = field_name & " is empty"
                IsValid = False
            Else
                errorMsg = Nothing
                IsValid = True
            End If
        End Function
    End Class

End Namespace
