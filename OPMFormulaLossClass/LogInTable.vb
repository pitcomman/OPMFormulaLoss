Imports System.Data.SqlClient

Public Class LogInTable
    Inherits Rist.OPMCmnClass.DBBase

    Public Sub New()

    End Sub

    Public Function GetLoginUserTbl(ByVal pstrOperator As String, ByVal pstrPassword As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Select * From OPMFormulaLoss.dbo.Operator Where OPID = '" & Trim(pstrOperator) & "' and Password = '" & Trim(pstrPassword) & "'"

            objSqlCmd.CommandType = CommandType.Text

            ' execute referance
            objDataTbl = MyBase.GetDataTable(objSqlCmd)

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objSqlCmd Is Nothing Then
                objSqlCmd.Dispose()
                objSqlCmd = Nothing
            End If
        End Try

        ' return value
        Return objDataTbl
    End Function

End Class
