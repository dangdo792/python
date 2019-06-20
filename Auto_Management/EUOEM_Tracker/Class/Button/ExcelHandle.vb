Imports Microsoft.Office.Interop.Excel

Class ExcelHandle

    Private Exel_Path As String
    Public objExcel As Application
    Public workbook As Workbook

    Public Sub New(ByVal Exel_Path_in As String)
        Exel_Path = Exel_Path_in

        ' Get excel object
        If GetObject("winmgmts:").ExecQuery("select * from win32_process where name='Excel.exe'").Count > 0 Then
            objExcel = GetObject(, "Excel.Application")
        Else
            objExcel = CreateObject("Excel.Application")
        End If
        objExcel.Visible = True
    End Sub

    Protected Overrides Sub Finalize()
        If Not objExcel Is Nothing Then
            If objExcel.Workbooks.Count = 0 Then
                objExcel.Quit()
            End If
        End If
    End Sub

    Public Sub ExitObject()
        If Not objExcel Is Nothing Then
            If objExcel.Workbooks.Count = 0 Then
                objExcel.Quit()
            End If
        End If
    End Sub

    Public Function Get_WB() As Workbook
        Me.workbook = Nothing
        Dim ShortName As String = Right(Exel_Path, Len(Exel_Path) - InStrRev(Exel_Path, "\"))
        For Each WB In objExcel.Workbooks
            If WB.Name = ShortName Then
                Me.workbook = WB
                Exit For
            End If
        Next

        If Me.workbook Is Nothing Then
            Me.workbook = objExcel.Workbooks.Open(Exel_Path)
        End If

        Return Me.workbook
    End Function

    Public Sub CloseWB()
        Me.workbook.Close(SaveChanges:=True)
    End Sub

    Public Function CannotUse() As Boolean
        Try
            Me.workbook = objExcel.Workbooks.Add()
            Me.CloseWB()
            Return False
        Catch ex As Exception
            Return True
        End Try
    End Function
End Class
