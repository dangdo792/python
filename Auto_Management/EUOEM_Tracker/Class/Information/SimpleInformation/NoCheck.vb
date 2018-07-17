Namespace Information
    ''' <summary>
    ''' Class for nocheck information
    ''' Do not perform anything, use the dafault behavior as base class do
    ''' </summary>
    Public Class NoCheck
        Inherits InfoBase
        Public Sub New(infovalue As String)
            MyBase.New(infovalue)
        End Sub

    End Class

End Namespace
