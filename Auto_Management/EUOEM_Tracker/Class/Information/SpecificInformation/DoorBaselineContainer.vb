Imports System.Text.RegularExpressions

Namespace Information
    ''' <summary>
    ''' Door baseline number information container
    ''' </summary>
    Public Class DoorBaseline
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
        ''' <returns>True if information is not null and have at least 2 number/ dot charater (".") , False otherwise</returns>
        Public Overrides Function IsValid() As Boolean
            If info_value = "" Then
                errorMsg = field_name & " is empty"
                IsValid = False
            ElseIf Regex.Matches(info_value, "^[0-9.]{2,}$").Count = 0 Then
                errorMsg = "Baseline must be a number"
                IsValid = False
            Else
                errorMsg = Nothing
                IsValid = True
            End If
        End Function
    End Class

End Namespace
