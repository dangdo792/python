Namespace Button
    Public Class GotoCQOrJIRA
        Inherits ButtonBase

        Private Enum Position
            TASK_ID
            USER
            PASSWORD
        End Enum

        Private FireFoxPath As String = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"

        Public Sub New(taskid_in As Information.TaskID,
                       user_in As Information.CheckNull,
                       password_in As Information.CheckNull)
            MyBase.New({taskid_in,
                       user_in,
                       password_in})
        End Sub

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing
            Dim taskid As String = listinfo(Position.TASK_ID).GetValue
            Dim username As String = listinfo(Position.USER).GetValue
            Dim password As String = listinfo(Position.PASSWORD).GetValue
            Dim Link As String = Nothing

            If InStr(taskid, "nprod") <> 0 Then
                Link = "http://rb-cq-da.de.bosch.com/cqweb/#/CQ2012prod/nprod/RECORD/" & taskid & "&noframes=true&format=HTML&recordType=Action&username=" & username '& "&password=" & password & ""
            Else
                Link = "https://rb-tracker.bosch.com/tracker08/browse/" & listinfo(Position.TASK_ID).GetValue
            End If
            Call Shell(FireFoxPath & " - url" & " " & Link, vbMaximizedFocus)
            Return ErrorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim IsValid As Boolean = True
            additional_errorMsg = IsFileExist(FireFoxPath)
            If Not String.IsNullOrEmpty(additional_errorMsg) Then
                IsValid = False
            End If
            Return IsValid
        End Function

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
