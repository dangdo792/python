Imports VBScript_RegExp_55

Public Module reg
    Function Replace_f(ByVal strInput As String, ByVal strPattern As String, ByVal strReplace As String)
        Dim regEx As New RegExp
        With regEx
            .Global = True
            .Multiline = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With
        If regEx.Test(strInput) Then
            Replace_f = regEx.Replace(strInput, strReplace)
        Else
            Replace_f = strInput
        End If
    End Function

    Function Search_f(ByVal strInput As String, ByVal strPattern As String) As String
        Dim regEx As New RegExp
        With regEx
            .Global = True
            .Multiline = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With
        If regEx.Test(strInput) Then
            Search_f = regEx.Execute(strInput)(0).value.ToString
        Else
            Search_f = ""
        End If

    End Function
End Module
