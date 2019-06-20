Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq

Namespace Button
    Class FillChecklist
        Inherits FillBase

        Const SCRIPT_ROW_MAX = 5

        ''' <summary>
        ''' The position of first 7 field is a must, do not adapt it
        ''' </summary>
        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
            MY_NAME
            REVIEWER
            LEADER
            MODULE_PATH
            OPL
            RS
            RS_BL
            SD
            SD_BL
            TS
            TS_BL
            STATEMENT
            DECISIONS
            DEFECT_ID
            OLD_TASK
            SANDBOX_FOLDER
            SANDBOX
        End Enum


        Public Sub New(explorer_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       myname_in As Information.CheckNull,
                       reviewer_in As Information.NoCheck,
                       leader_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath,
                       opl_in As Information.NoCheck,
                       rs_in As Information.NoCheck,
                       rsbl_in As Information.NoCheck,
                       sd_in As Information.NoCheck,
                       sdbl_in As Information.NoCheck,
                       ts_in As Information.NoCheck,
                       tsbl_in As Information.NoCheck,
                       statement_in As Information.CheckNull,
                       decisions_in As Information.NoCheck,
                       defectid_in As Information.NoCheck,
                       oldtask_in As Information.NoCheck,
                       sfd_in As Information.CheckNull,
                       sandbox_in As Information.CheckNull)

            MyBase.New({explorer_in,
                        wfd_in,
                        project_in,
                        taskid_in,
                        modulename_in,
                        myname_in,
                        reviewer_in,
                        leader_in,
                        modulepath_in,
                        opl_in,
                        rs_in,
                        rsbl_in,
                        sd_in,
                        sdbl_in,
                        ts_in,
                        tsbl_in,
                        statement_in,
                        decisions_in,
                        defectid_in,
                        oldtask_in,
                        sfd_in,
                        sandbox_in})
        End Sub

        Private ProjectFullName As String

        Public Overrides Function DoFunctionality() As String
            Dim errormsg As String = Nothing

            Dim checklistWB = New ExcelHandle(GetFullPath() & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_Review_Checklist.xlsx")

            Dim WB = checklistWB.Get_WB

            Dim WS1 As Worksheet = WB.Worksheets("Version history")
            Dim WS2 As Worksheet = WB.Worksheets("Review data")
            Dim WS3 As Worksheet = WB.Worksheets("TestAnalysis")
            Dim WS4 As Worksheet = WB.Worksheets("TC&TS review criteria")
            Dim WS5 As Worksheet = WB.Worksheets("Pre-Delivery Check")
            Dim WS6 As Worksheet = WB.Worksheets("Findings")

            WS1.Range("C11:C14").Value = CStr(Now().Date)
            WS1.Range("F11").Value = listinfo(Position.MY_NAME).GetValue
            WS1.Range("F12:F13").Value = listinfo(Position.REVIEWER).GetValue
            WS1.Range("F14").Value = listinfo(Position.LEADER).GetValue
            WS1.Range("F11:F14").Font.Color = Color.Black
            WS1.Range("C11:C14").Font.Color = Color.Black

            WS2.Range("B2").Value = ProjectFullName
            WS2.Range("B2").Font.Color = Color.Blue
            WS2.Range("B8:B9").Value = ""
            WS2.Range("C8:C9").Value = ""
            WS2.Range("D8:D9").Value = ""
            If String.IsNullOrEmpty(listinfo(Position.OPL).GetValue) Then
                WS2.Range("B8").Value = "test_analysis.xls" 'listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_test_analysis.xls"
                WS2.Range("C8").Value = "n.a"
            Else
                WS2.Range("B8").Value = "test_analysis.xls" 'listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_test_analysis.xls"
                WS2.Range("B9").Value = UCase(listinfo(Position.MODULE_NAME).GetValue) & "_OPL.xls"
                WS2.Range("C8:C9").Value = "n.a"
            End If
            WS2.Range("B8:C10").Font.Color = Color.Blue

            'Sear .hpp file path in sandbox 
            Dim hppfilepath As String = Nothing
            Dim sandboxPath As String = listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue
            Dim files() = Directory.GetFiles(sandboxPath, listinfo(Position.MODULE_NAME).GetValue & ".hpp", SearchOption.AllDirectories)
            For Each file As String In files
                If InStr(file, sandboxPath & "\build") = 0 Then
                    hppfilepath = file.Replace(sandboxPath, "")
                End If
            Next
            WS2.Range("B14").Value = Mid(hppfilepath, 2, Len(hppfilepath)).Replace("\", "/")

            'Get Hash of hpp file
            Dim getHashCommit As New Button.GetHashCommit(listinfo(Position.SANDBOX_FOLDER), listinfo(Position.SANDBOX), listinfo(Position.MODULE_PATH))
            getHashCommit.SpecSourceFlag = True
            getHashCommit.SpecSourcePath = hppfilepath
            errormsg = getHashCommit.Execute()

            WS2.Range("C14").Value = getHashCommit.hash
            WS2.Range("C14").Font.Color = Color.Blue
            WS2.Range("D14").Value = ""
            WS2.Range("B14").Font.Color = Color.Blue
            WS2.Range("B19").Value = "ipg.cop"
            WS2.Range("C19:C22").Value = "n.a"
            WS2.Range("B19:C22").Font.Color = Color.Blue
            Dim aPath() As String
            aPath = Split(listinfo(Position.MODULE_PATH).GetValue, "\")
            WS2.Range("B21").Value = "test_" & aPath(UBound(aPath))
            WS2.Range("B35:B36").Value = "'"

            WS2.Range("B40").Value = listinfo(Position.MY_NAME).GetValue
            WS2.Range("B40").Font.Color = Color.Black
            WS2.Range("C43").Value = listinfo(Position.REVIEWER).GetValue
            WS2.Range("C44").Value = listinfo(Position.REVIEWER).GetValue
            WS2.Range("C45").Value = listinfo(Position.LEADER).GetValue
            WS2.Range("D43:D45").Font.Color = Color.Black
            WS2.Range("B40").Font.Color = Color.Blue
            WS2.Range("B43:C45").Font.Color = Color.Blue
            WS2.Range("C40").Font.Color = Color.Black

            If String.IsNullOrEmpty(listinfo(Position.RS).GetValue) Then
                WS2.Range("A32:C32").Value = ""
                WS4.Range("D2:D3").Value = "N.A"
                WS4.Range("E2:E3").Value = "RQM is not available for DA core."
                WS4.Range("D10").Value = "N.A"
                WS4.Range("E10").Value = "RQM is not available for DA core."
            End If

            If String.IsNullOrEmpty(listinfo(Position.OLD_TASK).GetValue) Then
                WS3.Range("D2:D3").Value = "N.A"
                WS3.Range("E2:E3").Value = "New task"
            Else
                WS3.Range("D2:D3").Value = "Y"
                WS3.Range("E2:E3").Value = ""
            End If
            WS3.Range("D2:E3").Font.Color = Color.Black
            If String.IsNullOrEmpty(listinfo(Position.OPL).GetValue) Then
                WS3.Range("D5:D6").Value = "N.A"
                WS3.Range("E5:E6").Value = "No OPL"
            Else
                WS3.Range("D5:D6").Value = "Y"
                WS3.Range("E5:E6").Value = ""
            End If
            WS3.Range("D5:E6").Font.Color = Color.Black

            If String.IsNullOrEmpty(listinfo(Position.DECISIONS).GetValue) Then
                If listinfo(Position.STATEMENT).GetValue = "100" Then
                    WS4.Range("D7").Value = "N.A"
                    WS4.Range("E7").Value = "Statement: 100%, Decision: 100%"
                Else
                    WS4.Range("D7").Value = "Y"
                    WS4.Range("E7").Value = "Statement: " & listinfo(Position.STATEMENT).GetValue & "%"
                End If
            Else
                WS4.Range("D7").Value = "Y"
                WS4.Range("E7").Value = "Statement: " & listinfo(Position.STATEMENT).GetValue & "%" & vbCrLf & "Decisions: " & listinfo(Position.DECISIONS).GetValue & "%"
            End If
            WS4.Range("D13").Value = "N.A"
            WS4.Range("E13").Value = ""
            WS4.Range("D2:E13").Font.Color = Color.Black

            If String.IsNullOrEmpty(listinfo(Position.OPL).GetValue) Then
                WS5.Range("D2").Value = "N.A"
                WS5.Range("E2").Value = "No OPL"
            Else
                WS5.Range("D2").Value = "Y"
                WS5.Range("E2").Value = "all open points are closed"

            End If
            WS5.Range("D2:E5").Font.Color = Color.Black

            WS6.Range("A13:I17").Value = ""

            Return errormsg
        End Function

        Public Overrides Function GetFullPath() As String
            Return GetDocumentFolderPath()
        End Function

        Overrides Function AdditionCondition() As Boolean
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
            Dim isValid As Boolean = True
            Dim temp As ExcelHandle = New ExcelHandle("")
            If temp.CannotUse() Then
                additional_errorMsg = additional_errorMsg & vbNewLine & "You are editing cell. Please check and release cell." & vbNewLine
                isValid = False
            Else
                additional_errorMsg = Nothing
                isValid = True
            End If
            temp.ExitObject()
            temp = Nothing
            GC.Collect()

            If isValid = True Then
                Dim isExistObject As Information.CheckPathExist = New Information.CheckPathExist(GetFullPath() & "\" & "Documents" & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue & "_Review_Checklist.xlsx")
                isValid = isExistObject.IsValid()
                additional_errorMsg = isExistObject.GetErrorMsg()
            End If

            Return isValid
        End Function


    End Class
End Namespace
