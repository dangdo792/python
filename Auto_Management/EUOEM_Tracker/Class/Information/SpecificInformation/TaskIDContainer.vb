Imports System.Text.RegularExpressions

Namespace Information
    ''' <summary>
    ''' Task ID container class
    ''' </summary>
    Public Class TaskID
        Inherits InfoBase

        Public Sub New(infovalue As String)
            MyBase.New(infovalue)
        End Sub

        ''' <summary>
        ''' Override function get valid status of information
        ''' </summary>
        ''' <returns>True if information not null and have number, False otherwise</returns>
        Public Overrides Function IsValid() As Boolean
            If info_value = "" Then
                errorMsg = "Task ID is empty"
                IsValid = False
            ElseIf Regex.Matches(info_value, "[0-9]+").Count = 0 Then
                errorMsg = "Task ID need to have number"
                IsValid = False
            Else
                errorMsg = Nothing
                IsValid = True
            End If
        End Function
    End Class

End Namespace