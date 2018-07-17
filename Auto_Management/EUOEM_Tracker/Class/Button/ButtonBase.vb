Namespace Button
    Public MustInherit Class ButtonBase

        ''' <summary>
        ''' listinfo(): Array of information container object
        ''' </summary>
        Public listinfo() As Information.InfoBase
        ''' <summary>
        ''' additional_errorMsg: Additional error message.
        ''' It is set in AdditionCondition() function and generate to final error message at function GetErrorMsg()
        ''' </summary>
        Public additional_errorMsg As String

        Public Sub New(listinfo_in() As Information.InfoBase)
            listinfo = listinfo_in
            additional_errorMsg = Nothing
        End Sub

        ''' <summary>
        ''' Main function, all button should call this function
        ''' </summary>
        ''' <returns></returns>
        Public Function Execute() As String
            ' Check if all preconditions are fulfilled 
            If IsValid() Then
                ' Execute main functionality of this button
                Return DoFunctionality()
            Else
                ' Return error message
                Return GetErrorMsg()
            End If
        End Function

        ''' <summary>
        ''' Precondition check
        ''' </summary>
        ''' <returns>True if all data is valid and additional condition is fulfilled, False otherwise</returns>
        Function IsValid() As Boolean
            Dim flag As Boolean = True
            ' Get validity status of all input data
            For Each item In listinfo
                flag = flag And item.IsValid
            Next

            ' Check addtiontional condition
            If flag = True Then
                flag = flag And AdditionCondition()
            End If

            ' Return the status
            Return flag
        End Function

        ''' <summary>
        ''' Perform more validity condition check
        ''' Default is True
        ''' </summary>
        ''' <returns>True if fulfilled, False otherwise</returns>
        Overridable Function AdditionCondition() As Boolean
            Return True
        End Function

        ''' <summary>
        ''' Main function, must be overrided by each button
        ''' </summary>
        MustOverride Function DoFunctionality() As String

        ''' <summary>
        ''' Get final error message
        ''' </summary>
        ''' <returns>String contain all error message</returns>
        Function GetErrorMsg() As String
            Dim msg As String = Nothing

            ' Get error message from each input information object
            For Each item In listinfo
                If item.errorMsg <> Nothing Then
                    msg = msg & item.errorMsg & vbNewLine
                End If
            Next

            ' Get additional error message into
            msg = msg & additional_errorMsg
            Return msg
        End Function

    End Class

End Namespace
