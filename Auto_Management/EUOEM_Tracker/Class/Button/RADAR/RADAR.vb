Imports Scripting

Namespace Button
    Class CreateRADAR
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            TOOL_FOLDER
            WORKING_FOLDER
            PROJECT
            MODEL
            RELEASE
            TASK_ID
            MODULE_NAME
            MODULE_PATH
            RS
            TS
        End Enum

        Private ReviewerID As String
        Private LeaderID As String

        Public Sub New(explorer_in As Information.CheckNull,
                       tfd_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       model_in As Information.CheckNull,
                       release_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath,
                       rs_in As Information.CheckNull,
                       ts_in As Information.CheckNull,
                       ReviewerID_in As String,
                       LeaderID_in As String)

            MyBase.New({explorer_in,
                       tfd_in,
                       wfd_in,
                       project_in,
                       model_in,
                       release_in,
                       taskid_in,
                       modulename_in,
                       modulepath_in,
                       rs_in,
                       ts_in})
            ReviewerID = ReviewerID_in
            LeaderID = LeaderID_in
        End Sub

        Private RADAR_File As String
        Private SourcePath As String

        Public Overrides Function DoFunctionality() As String

            Dim fso As Object
            fso = CreateObject("scripting.filesystemobject")

            Dim taskid_double As Double
            Double.TryParse(Search_f(listinfo(Position.TASK_ID).GetValue, "\d+"), taskid_double)

            Dim DesPath = GetFullPath() + "\"
            Call fso.CopyFile(SourcePath, DesPath)

            Dim content1 As String = Nothing

            Dim Ticket_Name As String = Nothing
            Dim Email_Name As String

            Ticket_Name = content1 & "EI-300010_" + listinfo(Position.PROJECT).GetValue + "_MT_00" + CStr(taskid_double) + "_TestScript_Review_1" + vbLf
            Email_Name = content1 & "[" & listinfo(Position.PROJECT).GetValue & "]" & "[" & listinfo(Position.MODEL).GetValue & "_" & listinfo(Position.RELEASE).GetValue & "]" + " " + listinfo(Position.TASK_ID).GetValue + "_" + listinfo(Position.MODULE_NAME).GetValue + " : TestScript Review" + vbLf

            Dim TS As TextStream
            Dim TempS As String
            Dim Final As String = Nothing
            TS = fso.OpenTextFile(DesPath & RADAR_File, IOMode.ForReading)
            Do Until TS.AtEndOfStream
                TempS = TS.ReadLine
                If (InStr(TempS, "REVIEW_LEADER")) Then
                    TempS = Replace(TempS, "REVIEW_LEADER", LeaderID)
                ElseIf (InStr(TempS, "NPROD_NUMBER")) Then
                    TempS = Replace(TempS, "NPROD_NUMBER", listinfo(Position.TASK_ID).GetValue)
                ElseIf (InStr(TempS, "NAME_OF_TICKET")) Then
                    TempS = Replace(TempS, "NAME_OF_TICKET", Ticket_Name)
                ElseIf (InStr(TempS, "REQUIREMENT_LINK_INFORMATION")) Then
                    TempS = Replace(TempS, "REQUIREMENT_LINK_INFORMATION", listinfo(Position.RS).GetValue)
                ElseIf (InStr(TempS, "TESTSPEC_LINK_INFORMATION")) Then
                    TempS = Replace(TempS, "TESTSPEC_LINK_INFORMATION", listinfo(Position.TS).GetValue)
                ElseIf (InStr(TempS, "REVIEWER_NAME")) Then
                    TempS = Replace(TempS, "REVIEWER_NAME", ReviewerID)
                ElseIf (InStr(TempS, "MODULE_NAME")) Then
                    TempS = Replace(TempS, "MODULE_NAME", listinfo(Position.MODULE_NAME).GetValue)
                End If
                Final = Final & TempS & vbCrLf
            Loop

            Dim listLines = Split(Final, vbCrLf)

            TS.Close()

            TS = fso.OpenTextFile(DesPath & RADAR_File, IOMode.ForWriting, True)
            TS.Write(Final)
            TS.Close()

            TS = Nothing
            fso = Nothing

            Dim NotePad_Dir = "C:\Program Files (x86)\Notepad++\notepad++.exe"
            Dim RADARText_Dir = DesPath & RADAR_File
            Call Shell(NotePad_Dir & " " & RADARText_Dir, vbNormalFocus)

            Return Nothing
        End Function

        ''' <summary>
        ''' Check if document folder is valid or not
        ''' </summary>
        ''' <returns>True if Document folder exist, False otherwise</returns>
        Overrides Function AdditionCondition() As Boolean
            Dim LinkObject As Information.CheckPathExist
            Dim IsValid As Boolean = True
            'Check Notepad is exist or not
            LinkObject = New Information.CheckPathExist("C:\Program Files (x86)\Notepad++\notepad++.exe")
            IsValid = LinkObject.IsValid()
            additional_errorMsg = LinkObject.GetErrorMsg() & vbNewLine
            If IsValid Then
                'Check document folder is exist or not
                Dim path As String = GetFullPath()
                LinkObject = New Information.CheckPathExist(path)
                IsValid = LinkObject.IsValid()
                additional_errorMsg = LinkObject.GetErrorMsg() & vbNewLine
                If IsValid = True Then
                    ' Check RADAR template file is exist or not 
                    If Right(listinfo(Position.MODULE_PATH).GetValue, 2) = ".c" Then
                        RADAR_File = "Create_RADAR_Template_For_C.txt"
                    ElseIf Right(listinfo(Position.MODULE_PATH).GetValue, 4) = ".cpp" Then
                        RADAR_File = "Create_RADAR_Template_For_CPP.txt"
                    End If
                    SourcePath = listinfo(Position.TOOL_FOLDER).GetValue & "\" & RADAR_File
                    LinkObject = New Information.CheckPathExist(SourcePath)
                    IsValid = LinkObject.IsValid()
                    additional_errorMsg = LinkObject.GetErrorMsg()
                End If
            End If
            Return IsValid

        End Function

        ''' <summary>
        ''' Get full path of document
        ''' </summary>
        ''' <returns>String full path </returns>
        Function GetFullPath()
            Dim getdocpath As New Button.GotoDocument(listinfo(Position.EXPLORER), listinfo(Position.WORKING_FOLDER), listinfo(Position.PROJECT), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
            Return getdocpath.GetFullPath
        End Function
    End Class

End Namespace
