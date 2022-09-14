Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass


Module MdGrdFrm

    Public OpMonthTbl As DataTable
    Public VarOPMonth As String



#Region " OpMonthDetailTbl Table Design "

    Public Function CreateOpMonthDetailTbl() As DataTable
        Dim tmpTable As New DataTable
        Dim tmpColumn As DataColumn

        tmpColumn = New DataColumn
        tmpColumn.DataType = Type.GetType("System.String")
        tmpColumn.ColumnName = "OpMonth"
        tmpColumn.MaxLength = 6
        tmpTable.Columns.Add(tmpColumn)


        tmpColumn = New DataColumn
        tmpColumn.DataType = Type.GetType("System.DateTime")
        tmpColumn.ColumnName = "FromDate"
        tmpTable.Columns.Add(tmpColumn)

        tmpColumn = New DataColumn
        tmpColumn.DataType = Type.GetType("System.DateTime")
        tmpColumn.ColumnName = "ToDate"
        tmpTable.Columns.Add(tmpColumn)



        tmpColumn = New DataColumn
        tmpColumn.DataType = Type.GetType("System.String")
        tmpColumn.ColumnName = "MonthlyClosingFlag"
        'tmpColumn.MaxLength = 1
        tmpTable.Columns.Add(tmpColumn)



        tmpColumn = New DataColumn
        tmpColumn.DataType = Type.GetType("System.String")
        tmpColumn.ColumnName = "OperatorName"
        tmpColumn.MaxLength = 50
        tmpTable.Columns.Add(tmpColumn)




        Return tmpTable
    End Function
#End Region

#Region " Set Style for OpMonthTbl "
    Public Sub initialOpMonthTblStyle(ByVal grd As DataGrid, ByVal dtTable As DataTable)
        'Initial OpMonthData
        Dim OpMonthData As New DataView(dtTable)
        OpMonthData.AllowNew = False



        With grd
            .CaptionVisible = False
            .ColumnHeadersVisible = True
            .RowHeadersVisible = False
            .DataSource = OpMonthData

        End With

        ' You must clear out the TableStyles collection before 
        grd.TableStyles.Clear()

        Dim grdTableStyle1 As New DataGridTableStyle
        With grdTableStyle1
            .MappingName = OpMonthData.Table.TableName
        End With


        Dim grdColStyle1 As New DataGridLabelColumn
        With grdColStyle1
            .MappingName = "OpMonth"
            .HeaderText = "OpMonth"
            .Width = 70
        End With


        Dim grdColStyle2 As New DataGridLabelColumn
        With grdColStyle2
            .MappingName = "FromDate"
            .HeaderText = "FromDate"
            .Width = 130
        End With



        Dim grdColStyle3 As New DataGridLabelColumn
        With grdColStyle3
            .MappingName = "ToDate"
            .HeaderText = "ToDate"
            .Width = 130
        End With


        Dim grdColStyle4 As New DataGridLabelColumn
        With grdColStyle4
            .MappingName = "MonthlyClosingFlag"
            .HeaderText = "MonthlyClosingFlag"
            .Width = 130
        End With



        Dim grdColStyle5 As New DataGridLabelColumn
        With grdColStyle5
            .MappingName = "OperatorName"
            .HeaderText = "OperatorName"
            .Width = 130
        End With




        grdTableStyle1.GridColumnStyles.AddRange _
        (New DataGridColumnStyle() {grdColStyle1, grdColStyle2, grdColStyle3, grdColStyle4, _
                                    grdColStyle5})

        grd.TableStyles.Add(grdTableStyle1)
        grdTableStyle1.AllowSorting = False

    End Sub
#End Region




End Module
