Imports Scripting

Namespace Button
    Public Class GetReviewScriptInfo
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            REVIEW_FOLDER
            TASK_ID
            MODULE_NAME
        End Enum

        Public Sub New(explorer_in As Information.CheckNull,
                       rfd_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull)
            MyBase.New({explorer_in,
                         rfd_in,
                       taskid_in,
                       modulename_in})
        End Sub

        Private Script_File_Dir As String

        Public project As String
        Public model As String
        Public release As String
        Public oldtask As String
        Public modulepath As String
        Public branch As String
        Public sandbox As String
        Public revision As String


        Public Overrides Function DoFunctionality() As String
            Dim fso As FileSystemObject
            Dim TS As TextStream
            Dim TempS As String
            Dim Final As String = Nothing
            fso = New FileSystemObject
            TS = fso.OpenTextFile(Script_File_Dir, IOMode.ForReading)

            Do Until TS.AtEndOfStream
                TempS = TS.ReadLine
                If TS.Line = 3 Then
                    project = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                ElseIf TS.Line = 4 Then
                    model = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                ElseIf TS.Line = 5 Then
                    modulepath = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                ElseIf TS.Line = 6 Then
                    If InStr(TempS, ",") <> 0 Then
                        Dim Nprod_Temp As String = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                        oldtask = Trim(Left(Nprod_Temp, InStr(Nprod_Temp, ",") - 1))
                    End If
                ElseIf TS.Line = 7 Then
                    revision = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                ElseIf TS.Line = 8 Then
                    If InStr(TempS, "Feature_Branch") Then
                        branch = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                    Else
                        sandbox = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                    End If
                ElseIf TS.Line = 9 Then
                    release = Trim(Right(TempS, Len(TempS) - InStr(TempS, "=")))
                End If
                Final = Final & TempS & vbCrLf
            Loop
            Return Nothing
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim path As String = GetFullPath()
            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")
            If Not fso.FolderExists(path) Then
                additional_errorMsg = "Folder doesn't exist." & vbNewLine & "Please check: " & path
                Return False
            Else
                Script_File_Dir = GetFullPath() & "\" & "Scripts"
                If fso.FileExists(Script_File_Dir & "\" & listinfo(Position.MODULE_NAME).GetValue & ".ptu") Then
                    Script_File_Dir = Script_File_Dir & "\" & listinfo(Position.MODULE_NAME).GetValue & ".ptu"
                ElseIf fso.FileExists(Script_File_Dir & "\" & listinfo(Position.MODULE_NAME).GetValue & ".otd") Then
                    Script_File_Dir = Script_File_Dir & "\" & listinfo(Position.MODULE_NAME).GetValue & ".otd"
                End If
                If Not fso.FileExists(Script_File_Dir) Then
                    additional_errorMsg = "Script file doesn't exist" & vbNewLine & "Please check: " & Script_File_Dir
                    Return False
                Else
                    additional_errorMsg = Nothing
                    Return True
                End If

            End If
        End Function

        Function GetFullPath()
            Dim getrevpath As New Button.GotoReview(listinfo(Position.EXPLORER), listinfo(Position.REVIEW_FOLDER), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
            Return getrevpath.GetFullPath
        End Function

    End Class
End Namespace
