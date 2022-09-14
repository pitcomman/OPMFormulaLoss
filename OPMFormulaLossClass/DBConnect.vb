'***********************************************************************
' Program Name	    : DBConect Class
' Program ID	    : DBConnect
' Function			: this Class have 
' Create Date		: 2006/06/15
' Create Person		: Athicha J.
' 
' Supplement	    :
' Version		    : 1.00
' ---------------------------------------------------------------------
' Condition	        : SqlServer2000,ADO.Net,.NetFramework
' Starting Way		:
'***********************************************************************
Imports System.Data.SqlClient
Imports Rist.OPMCmnClass


Public Class DBConnect
    Inherits DBBase

    Public Function GetOpMonthTbl(ByVal strOpMonth As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Select * from OPMFormulaLoss.dbo.OperationMonth where OpMonth = '" & strOpMonth & "' "
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

    Public Function GetOpMonthTbl2() As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Select * from OPMFormulaLoss.dbo.VewFormulaLossHistOPMonth "
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

    Public Function GetOpMonthTbl3(ByVal strOpMonth As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Select * from OPMFormulaLoss.dbo.VewFormulaLossHistOPMonth where OpMonth = '" & strOpMonth & "' "
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

    Public Function AllOpmonthDisplay(ByVal strOpMonth As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Select * from OPMFormulaLoss.dbo.OperationMonth where OpMonth Like '" & strOpMonth & "%'"
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

    Public Function OpMonthInsert(ByVal strOpMonth As String, ByVal strFromDate As DateTime, _
                                  ByVal strToDate As DateTime, _
                                  ByVal strOperatorName As String) As DataTable

        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing


        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Exec sprOpMonthInsert '" _
                        & Trim(strOpMonth) & "','" _
                        & Trim(strFromDate.ToString("MM/dd/yyyy")) & "','" _
                        & Trim(strToDate.ToString("MM/dd/yyyy")) & "','" _
                        & Trim(strOperatorName) & "'"

            'objSqlCmd.CommandText = "Exec sprOpMonthInsert '" _
            '& Trim(strOpMonth) & "','" _
            '& Trim(strFromDate) & "','" _
            '& Trim(strToDate) & "','" _
            '& Trim(strOperatorName) & "'"
            '& Trim(strMonthlyClosingFlag) & "','" _


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


    Public Function OpMonthEdit(ByVal strOpMonth As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Delete from OPMFormulaLoss.dbo.OperationMonth where OpMonth = '" & strOpMonth & "' "
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



    Public Function OpMonthCancel(ByVal strOpMonth As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "Delete From OPMFormulaLoss.dbo.OperationMonth Where OpMonth = '" & strOpMonth & "' "

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

    Public Function QueryAll(ByVal target As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "SELECT * FROM " + target
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
