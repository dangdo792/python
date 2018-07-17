Imports Microsoft.Office.Interop.Excel

Namespace Button
    ''' <summary>
    ''' Fill OPL button
    ''' </summary>
    Class FillOPL
        Inherits FillBase

        ''' <summary>
        ''' The position of first 7 field is a must, do not adapt it
        ''' </summary>
        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
            MODULE_OWNER
            MY_NAME
        End Enum

        Public Sub New(explorer_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       mo_in As Information.CheckNull,
                       myname_in As Information.CheckNull)

            MyBase.New({explorer_in,
                       wfd_in,
                       project_in,
                       taskid_in,
                       modulename_in,
                       mo_in,
                       myname_in})
        End Sub

        Public Overrides Function DoFunctionality() As String
            Dim OPLWB = New ExcelHandle(GetFullPath())

            Dim WB As Workbook = OPLWB.Get_WB
            Dim WS As Worksheet = WB.Worksheets("OPL")
            For i = 5 To 10
                WS.Range("A" & i).Value = i - 4
                WS.Range("B" & i).Value = listinfo(Position.TASK_ID).GetValue
                WS.Range("F" & i).Value = CStr(Now().Date)
                WS.Range("G" & i).Value = listinfo(Position.MODULE_OWNER).GetValue
                WS.Range("I" & i).Value = listinfo(Position.MY_NAME).GetValue
                WS.Range("K" & i).Value = "Open"
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Get full path OPL file
        ''' </summary>
        ''' <returns>String full OPL path</returns>
        Public Overrides Function GetFullPath() As String
            Return Me.GetDocumentFolderPath & "\" & "Documents" & "\" & UCase(listinfo(Position.MODULE_NAME).GetValue) & "_OPL.xls"
        End Function

        Overrides Function AdditionCondition() As Boolean
            ' Check if excel editable or not
            Dim temp As ExcelHandle = New ExcelHandle("")
            Dim isExistObject As Information.CheckPathExist = New Information.CheckPathExist(GetFullPath())
            Dim isValid As Boolean = True
            isValid = isExistObject.IsValid()
            additional_errorMsg = isExistObject.GetErrorMsg()
            If temp.CannotUse() Then
                additional_errorMsg = additional_errorMsg & vbNewLine & "You are editing cell. Please check and release cell." & vbNewLine
                isValid = False
            End If
            Return isValid
        End Function

    End Class

End Namespace
