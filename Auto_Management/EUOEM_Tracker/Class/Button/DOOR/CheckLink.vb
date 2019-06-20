Imports Scripting

Namespace Button
    Class CheckLinkDOORsDoc
        Inherits ButtonBase

        Private Enum Position
            TOOL_FOLDER
            RELEASE
            RS
            TS
            MODEL
        End Enum

        Public Sub New(tfd_in As Information.CheckNull,
                       release_in As Information.CheckNull,
                       rs_in As Information.CheckNull,
                       ts_in As Information.CheckNull,
                       model_in As Information.CheckNull)
            MyBase.New({tfd_in,
                       release_in,
                       rs_in,
                       ts_in,
                       model_in})
        End Sub

        Private Path As String

        Public Overrides Function DoFunctionality() As String
            Dim ErrorMsg As String = Nothing

            Dim ID_SWRS As String
            Dim ID_SWTS As String

            Dim arr_Data() As String
            Dim str_Data As String
            Dim fso As New FileSystemObject

            ID_SWRS = Replace_f(listinfo(Position.RS).GetValue, ".+[OV]-(\d+).+", "$1")
            ID_SWTS = Replace_f(listinfo(Position.TS).GetValue, ".+[OV]-(\d+).+", "$1")

            str_Data = fso.GetFile(Path).OpenAsTextStream.ReadAll
            arr_Data = Split(str_Data, vbNewLine)
            Dim Counter_Flag As Integer
            Counter_Flag = 0
            Dim txt_file As Object
            txt_file = fso.OpenTextFile(Path, IOMode.ForWriting)

            For i = 0 To UBound(arr_Data) - 1
                If InStr(arr_Data(i), "string TestResultStatus =") <> 0 Then
                    arr_Data(i) = Replace_f(arr_Data(i), "(.+=).+", "$1" + " ""Test Result Status " & listinfo(Position.MODEL).GetValue & " " & listinfo(Position.RELEASE).GetValue + """")
                End If
                If InStr(arr_Data(i), "string RequirementStatus =") <> 0 Then
                    arr_Data(i) = Replace_f(arr_Data(i), "(.+=).+", "$1" + " ""Status_" & listinfo(Position.MODEL).GetValue & """")
                End If
                If InStr(arr_Data(i), "int ModuleReqAbsnumber =") <> 0 Then
                    arr_Data(i) = Replace_f(arr_Data(i), "(=\W+)\w+", "$1" + ID_SWRS)
                    Counter_Flag = 1
                End If
                If InStr(arr_Data(i), "int ModuleTestAbsnumber =") <> 0 Then
                    arr_Data(i) = Replace_f(arr_Data(i), "(=\W+)\w+", "$1" + ID_SWTS)
                    Counter_Flag = 2
                End If
                txt_file.WriteLine(arr_Data(i))
            Next i
            txt_file.Close()

            If Counter_Flag <> 2 Or ID_SWRS = "" Or ID_SWTS = "" Then
                txt_file = fso.OpenTextFile(Path, IOMode.ForWriting)
                txt_file.Write(str_Data)
            Else
                Shell("C:\Program Files (x86)\Notepad++\notepad++.exe " + Path, vbNormalFocus)
            End If

            Return ErrorMsg
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim LinkObject As Information.CheckPathExist
            Dim IsValid As Boolean = True
            'Check Notepad is exist or not
            LinkObject = New Information.CheckPathExist("C:\Program Files (x86)\Notepad++\notepad++.exe")
            IsValid = LinkObject.IsValid()
            additional_errorMsg = LinkObject.GetErrorMsg() & vbNewLine
            If IsValid Then
                ' Check if the path is exist or not 
                Path = listinfo(Position.TOOL_FOLDER).GetValue + "\CheckLinkingFeatureInDOORS_CCT.txt"
                LinkObject = New Information.CheckPathExist(Path)
                IsValid = LinkObject.IsValid()
                additional_errorMsg = LinkObject.GetErrorMsg()
            End If
            Return IsValid
        End Function

    End Class

End Namespace
