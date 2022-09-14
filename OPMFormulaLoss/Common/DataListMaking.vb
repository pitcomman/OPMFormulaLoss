Imports Rist.OPMCmnClass
Imports System.Data.SqlClient

Public Class DataListMaking
    Inherits DBBase
    Public Function GetDataTableCnt(ByVal prmStrSql As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmStrSql
            objSqlCmd.CommandType = CommandType.Text
            objSqlCmd.CommandTimeout = 0
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

    Public Function GetDataTableSql(ByVal prmStrSql As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing
        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmStrSql
            objSqlCmd.CommandType = CommandType.StoredProcedure
            objSqlCmd.CommandTimeout = 0
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

    Public Function UpdateFormulaLossList(ByVal strOpName As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Update [OPMFormulaLoss].dbo.[FormulaLossList] set OperatorName = '" & strOpName & "' "
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
