Namespace Information
    ''' <summary>
    ''' The abstract base class for information container
    ''' </summary>
    Public Class InfoBase
        Public info_value As String
        Public errorMsg As String

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="infovalue"></param>
        Public Sub New(infovalue As String)
            info_value = infovalue
            errorMsg = Nothing
        End Sub

        ''' <summary>
        ''' Get the valid status of information
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function IsValid() As Boolean
            IsValid = True
            errorMsg = Nothing
        End Function

        ''' <summary>
        ''' Get error message of information
        ''' </summary>
        ''' <returns></returns>
        Public Function GetErrorMsg() As String
            GetErrorMsg = errorMsg
        End Function

        ''' <summary>
        ''' Get the raw value of information
        ''' </summary>
        ''' <returns></returns>
        Public Function GetValue() As String
            GetValue = info_value
        End Function

        ''' <summary>
        ''' Set the raw value of information
        ''' </summary>
        ''' <param name="info_value_in"></param>
        Public Sub SetValue(info_value_in As String)
            info_value = info_value_in
        End Sub
    End Class
End Namespace
