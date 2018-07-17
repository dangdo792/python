Namespace Button
    ''' <summary>
    ''' Goto ILM path or computer path object
    ''' </summary>
    Public Class GotoFolderOrILM
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            LINK
        End Enum

        Private FireFoxPath As String = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        Private Link As String
        Private is_ILM_link As Boolean

        Public Sub New(explorerpath_in As Information.CheckNull, opllink_in As Information.CheckNull)
            MyBase.New({explorerpath_in, opllink_in})

            ' If the link has https://, it is the ILM link
            If InStr(listinfo(Position.LINK).GetValue, "https://") <> 0 Then
                Me.is_ILM_link = True
            Else
                Me.is_ILM_link = False
            End If

            ' Corrected the link
            Me.Link = GetCorrectPath()
        End Sub

        Public Overrides Function DoFunctionality() As String
            If is_ILM_link Then
                ' Call firefox when it is an ILM link_
                Call Shell(FireFoxPath & " -url" & " " & Me.Link, vbMaximizedFocus)
            Else
                ' Call explorer when it is ...
                Call Shell(listinfo(Position.EXPLORER).GetValue & " " & Me.Link, vbNormalFocus)
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Check if the link to folder is exist or not
        ''' </summary>
        ''' <returns>True if the link exist, False otherwise</returns>
        Public Overrides Function AdditionCondition() As Boolean
            Dim IsValid As Boolean = True

            ' If the link is not ILM, check the validity of the corrected link
            If Not is_ILM_link Then
                Dim LinkObject As Information.CheckPathExist
                'Check Explorer is exist or not
                LinkObject = New Information.CheckPathExist(listinfo(Position.EXPLORER).GetValue)
                IsValid = LinkObject.IsValid()
                additional_errorMsg = LinkObject.GetErrorMsg() & vbNewLine
                If IsValid Then
                    ' Check if the path is exist or not 
                    LinkObject = New Information.CheckPathExist(Me.Link)
                    IsValid = LinkObject.IsValid()
                    additional_errorMsg = LinkObject.GetErrorMsg()
                End If
            End If
            Return IsValid
        End Function

        ''' <summary>
        ''' Correct the link
        ''' </summary>
        ''' <returns>Use 2 helper function for ILM and FOLDER to correct the link for individual case</returns>
        Public Overridable Function GetCorrectPath() As String
            If Me.is_ILM_link Then
                Return GetCorrectPath_ILM()
            Else
                Return GetCorrectPath_Folder()
            End If
        End Function
        Public Overridable Function GetCorrectPath_ILM() As String
            Return listinfo(Position.LINK).GetValue
        End Function
        Public Overridable Function GetCorrectPath_Folder() As String
            Return listinfo(Position.LINK).GetValue
        End Function

    End Class

End Namespace
