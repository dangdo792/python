Imports System.IO

Module tex
    Public Sub init(ByRef ctr As Control, ByRef bs As BindingSource)
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            For Each textboxes As MetroFramework.Controls.MetroTextBox In ctr.Controls.OfType(Of MetroFramework.Controls.MetroTextBox)()
                If textboxes.Name <> "RS_InfoTextBox" And
                   textboxes.Name <> "SD_InfoTextBox" And
                   textboxes.Name <> "TS_InfoTextBox" Then
                    Try
                        textboxes.DataBindings.Clear()
                        textboxes.DataBindings.Add("Text", bs, Replace(textboxes.Name, "TextBox", ""), False, DataSourceUpdateMode.OnPropertyChanged)
                    Catch ex As Exception
                    End Try

                End If
            Next
        End If
    End Sub
End Module
