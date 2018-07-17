Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq

Namespace Button
    ''' <summary>
    ''' Fill Coverage button
    ''' </summary>
    Class FillCodeCoverage
        Inherits FillBase

        ''' <summary>
        ''' The position of first 7 field is a must, do not adapt it
        ''' </summary>
        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            TASK_ID
            MODULE_NAME
            MY_NAME
            STATEMENT
            DECISIONS
            MODULE_PATH
            HASH
            REVIEW_FOLDER
            PROJECT
        End Enum

        Public Sub New(explorer_in As Information.CheckNull,
                       myname_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       statement_in As Information.CheckNull,
                       decisions_in As Information.NoCheck,
                       modulepath_in As Information.ModulePath,
                       Hash_in As Information.CheckNull,
                      rfd_in As Information.CheckNull,
                       project_in As Information.CheckNull)

            MyBase.New({explorer_in,
                        wfd_in,
                        taskid_in,
                        modulename_in,
                        myname_in,
                        statement_in,
                        decisions_in,
                        modulepath_in,
                        Hash_in,
                        rfd_in,
                        project_in})
        End Sub

        Private placetask As String
        Private ProjectFullName As String

        Public Overrides Function DoFunctionality() As String


            Dim CodeCoverageWB = New ExcelHandle(placetask)

            Dim WB As Workbook
            Dim WS As Worksheet
            WB = CodeCoverageWB.Get_WB
            WS = WB.Worksheets("Version history")
            WS.Range("E11").Value = listinfo(Position.MY_NAME).GetValue
            WS.Range("C11").Value = Now()

            WS = WB.Worksheets("Code_Coverage_Exception")
            WS.Range("C3").Value = ProjectFullName
            WS.Range("C4").Value = listinfo(Position.TASK_ID).GetValue
            WS.Range("C5").Value = listinfo(Position.MODULE_NAME).GetValue
            WS.Range("C6").Value = listinfo(Position.HASH).GetValue
            WS.Range("D7").Value = WS.Range("D7").Replace("<RAFIVEDAI-1674>", listinfo(Position.TASK_ID).GetValue)
            WS.Range("C7").Characters(Start:=54, Length:=listinfo(Position.TASK_ID).GetValue.Length + 1).Font.FontStyle = "Bold"
            WS.Range("D10").Value = "N.A"
            WS.Range("D11").Value = listinfo(Position.STATEMENT).GetValue & " %"
            WS.Range("D12").Value = listinfo(Position.DECISIONS).GetValue & " %"

            WS.Range("C15:D15").Value = ""
            Return Nothing
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim IsValid As Boolean = True

            Dim strURLJ As String : strURLJ = "https://rb-tracker.bosch.com/tracker08/rest/api/2/issue/" & listinfo(Position.TASK_ID).GetValue
            Dim winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
            winHttp.SetAutoLogonPolicy(0)
            winHttp.Open("GET", strURLJ)
            Try
                winHttp.Send
            Catch ex As Exception
                Return "Can't access to JIRA" & vbCrLf
                Exit Function
            End Try

            Dim json As JObject
            Try
                json = JObject.Parse(winHttp.responseText)
            Catch ex As Exception
                Return "Can't access to JIRA. Please check access right" & vbCrLf
                Exit Function
            End Try
            ProjectFullName = json.SelectToken("fields").SelectToken("project").SelectToken("name")

            ' Check if excel editable or not
            Dim temp As ExcelHandle = New ExcelHandle("")
            If temp.CannotUse() Then
                additional_errorMsg = "You are editing cell. Please check and release cell." & vbNewLine
                IsValid = False
            Else

                Dim GotoDoc As New Button.GotoDocument(listinfo(Position.EXPLORER), listinfo(Position.WORKING_FOLDER), listinfo(Position.PROJECT), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
                placetask = GotoDoc.GetFullPath & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "_CodeCoverage_Exception.xls"

                Dim pathObj As New Information.CheckPathExist(placetask)
                IsValid = pathObj.IsValid()
                additional_errorMsg = pathObj.GetErrorMsg() & vbNewLine & vbNewLine
            End If
            Return IsValid
        End Function

        Public Overrides Function GetFullPath() As String
            Return Nothing
        End Function
    End Class

End Namespace