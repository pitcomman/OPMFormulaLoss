
Public Class ModifyType

    Private Key As String

    Public Shared ReadOnly Add As ModifyType = New ModifyType("Add")
    Public Shared ReadOnly Update As ModifyType = New ModifyType("Update")
    Public Shared ReadOnly Delete As ModifyType = New ModifyType("Delete")
    Public Shared ReadOnly Rollback As ModifyType = New ModifyType("Rollback")

    Private Sub New(key As String)
        Me.Key = key
    End Sub

    Public Overrides Function ToString() As String
        Return Me.Key
    End Function
End Class
