Imports System.Text.RegularExpressions

Namespace Button
    Class NeededContent
        Inherits ButtonBase

        Private Enum Position
            TOOL_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
            MODULE_PATH
            SANDBOX
            HASH
            COMMIT
            MY_NAME
            STATEMENT
            DECISIONS
            RESULT_PATH
            ELOC
        End Enum

        Private SysType As String = Nothing

        Public Sub New(tfd_in As Information.CheckNull,
                         project_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath,
                       sandbox_in As Information.CheckNull,
                       hash_in As Information.NoCheck,
                       commit_in As Information.NoCheck,
                       myname_in As Information.CheckNull,
                       statement_in As Information.NoCheck,
                       decision_in As Information.NoCheck,
                       resultpath_in As Information.NoCheck,
                       eloc_in As Information.NoCheck)
            MyBase.New({tfd_in,
                        project_in,
                       taskid_in,
                       modulename_in,
                       modulepath_in,
                       sandbox_in,
                       hash_in,
                       commit_in,
                       myname_in,
                       statement_in,
                       decision_in,
                       resultpath_in,
                       eloc_in})
        End Sub

        Public Overrides Function DoFunctionality() As String

            Dim content1 As String = Nothing

            Dim aPath() As String
            aPath = Split(listinfo(Position.MODULE_PATH).GetValue, "\")
            Dim fullmodulename As String = aPath(UBound(aPath))

            Dim statement As String
            Dim decisions As String
            If InStr(listinfo(Position.STATEMENT).GetValue, "100") <> 0 And String.IsNullOrEmpty(listinfo(Position.DECISIONS).GetValue) Then
                statement = "100"
                decisions = "100"
            Else
                statement = listinfo(Position.STATEMENT).GetValue
                decisions = listinfo(Position.DECISIONS).GetValue
            End If


            content1 = content1 &
"/***************-Git Command-*************/

Get Commit: git rev-parse HEAD
Get Hash: 	git rev-parse " & listinfo(Position.COMMIT).GetValue & ":" & Mid(listinfo(Position.MODULE_PATH).GetValue, 2, Len(listinfo(Position.MODULE_PATH).GetValue)).Replace("\", "/") & "

/**************-JIRA Infomation-***********/

/*---- For Start Task:-----*/

" & CStr(Now().Date) & ": " & listinfo(Position.MY_NAME).GetValue & ": Task is start with commit " & listinfo(Position.COMMIT).GetValue & "

{panel:title=Delivery Information}
||File_Name||File_Hash||Code_Coverage||Tested_ELOC||
|" & fullmodulename & "| " & listinfo(Position.HASH).GetValue & " | Function:	/, Statement: " & statement & "%, Decision: " & decisions & "% |" & listinfo(Position.ELOC).GetValue & "/" & listinfo(Position.ELOC).GetValue & " |
{panel}

/*---- For Delivery Task:-----*/

{panel:title=Delivery Information}
||File_Name||File_Hash||Code_Coverage||Tested_ELOC||
|" & fullmodulename & "| " & listinfo(Position.HASH).GetValue & " | Function: 	/, Statement: " & statement & "%, Decision: " & decisions & "% |" & listinfo(Position.ELOC).GetValue & "/" & listinfo(Position.ELOC).GetValue & " |

*Other info*
||Commit|" & listinfo(Position.COMMIT).GetValue & "|
||Requirement_baseline|NA|
||Delivery folder|" & listinfo(Position.RESULT_PATH).GetValue & "|
||Link to report|[Report|" & listinfo(Position.RESULT_PATH).GetValue & "\Reports\test_report.html" & "]|
{panel}

Summary: SW_UVE - Observation for " & fullmodulename & "
Summary: SW_UVE - OPL for " & fullmodulename & "
Labels: SW_UVE_OPL

Summary: Test Analysis Review for " & fullmodulename & "
Summary: TCDS&TS Review for " & fullmodulename & "
Summary: PDC for " & fullmodulename & "
Labels: SW_UVE_Review

/*************-Script Template-***************/

#define TASK_ID """ & listinfo(Position.TASK_ID).GetValue & """

// -> adding WRITE_LOG
WRITE_LOG(""Task ID:"" TASK_ID, cppth_ll_normal, false);

/************************************************************************************
    Verified requirements: requirement IDs / DOXYGEN
    Test goal: Verify that method perform:
                --> 
    In case: 
    Testing technique: Requirement based / Boundary values & Equivalence class
************************************************************************************/

/* Expected Call Sequence  */
/* Declare variables */
/* Input */
/* Function call */
/* Check output */

""--ci:--no_instr:all""
""--ci:--instr:stmt;decn#""" & fullmodulename & """:*""
--parse:--warning_suppress:9001
--parse:--warning_suppress:9815
--parse:--warning_suppress:11485
--parse:--warning_suppress:10696
--parse:--warning_suppress:9174

#include <limits> 	
std::numeric_limits<float>::max()
std::numeric_limits<float>::min()
std::numeric_limits<float>::infinity()

"

            Dim fso As Object
            fso = CreateObject("Scripting.FileSystemObject")
            Dim Fileout As Object
            Fileout = fso.CreateTextFile(listinfo(Position.TOOL_FOLDER).GetValue & "\Temp_File.cpp", True, True)
            Fileout.Write(content1)
            Fileout.Close()
            Call Shell("C:\Program Files (x86)\Notepad++\notepad++.exe " & listinfo(Position.TOOL_FOLDER).GetValue & "\Temp_File.cpp", 0, False)
            Return Nothing
        End Function

        Overrides Function AdditionCondition() As Boolean
            Dim LinkObject As Information.CheckPathExist
            Dim IsValid As Boolean = True
            'Check Notepad is exist or not
            LinkObject = New Information.CheckPathExist("C:\Program Files (x86)\Notepad++\notepad++.exe")
            IsValid = LinkObject.IsValid()
            additional_errorMsg = LinkObject.GetErrorMsg() & vbNewLine
            If IsValid = True Then
                'Check Tool folder is exist or not
                LinkObject = New Information.CheckPathExist(listinfo(Position.TOOL_FOLDER).GetValue)
                IsValid = LinkObject.IsValid()
                additional_errorMsg = LinkObject.GetErrorMsg() & vbNewLine
            End If
            Return IsValid
        End Function
    End Class

End Namespace
