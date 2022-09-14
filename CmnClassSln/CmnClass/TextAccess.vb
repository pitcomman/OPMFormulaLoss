'***********************************************************************
' Program Name	    : TextAccess
' Program ID	    : TextAccess
' Function			: 
' Create Date		: 2006/12/12
' Create Person		: H.Yamashita
' Supplement	    :
' Version		    : 1.0.0
' ---------------------------------------------------------------------
' Condition     	: SqlServer2000, ADO.Net, .NetFramework
' Starting Way		:
'***********************************************************************

Imports System.IO
Public Class TextAccess

    Private objReader As StreamReader


    Public Sub New(ByVal pstrFilePath As String)
        objReader = New StreamReader(pstrFilePath)
    End Sub


    ' close text
    Public Function CloseText()
        Try
            objReader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ' read row text
    Public Function GetText() As String
        Dim strRet As String
        Try
            strRet = objReader.ReadLine()
        Catch ex As Exception
            Throw ex
        End Try
        Return strRet
    End Function

End Class
