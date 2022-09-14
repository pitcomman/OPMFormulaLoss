'***********************************************************************
' Program Name	    : Log In Class
' Program ID	    : LogInTable
' Function			: this Class have OPMMaterialLedger app Login function
' Create Date		: 2006/07/27
' Create Person		: iwami
' 
' Supplement	    :
' Version		    : 1.00
' ---------------------------------------------------------------------
' Condition　　　　	: SqlServer2000,ADO.Net,.NetFramework
' Starting Way		:
'***********************************************************************
Imports System.Data.SqlClient
Imports Rist.OPMCmnClass

Public Class Authentication
    Inherits DBBase

    Public Function AuthenticateOperator(ByVal parameters As System.Collections.Specialized.StringDictionary) As DataTable

        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = "sprOperatorAuthentication"
            objSqlCmd.CommandType = CommandType.StoredProcedure

            For Each key As String In parameters.Keys
                objSqlCmd.Parameters.Add(key, parameters.Item(key))
            Next

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


End Class
