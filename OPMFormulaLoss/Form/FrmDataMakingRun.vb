Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass

#Region "FormulaLossMaking "
Public Class FormulaLossMaking
    Inherits Rist.OPMCmnClass.DialogControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtOpName As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtOpName = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtOpName
        '
        Me.txtOpName.Location = New System.Drawing.Point(72, 160)
        Me.txtOpName.Name = "txtOpName"
        Me.txtOpName.Size = New System.Drawing.Size(112, 16)
        Me.txtOpName.TabIndex = 40
        Me.txtOpName.Visible = False
        '
        'MaterialLedgerMaking
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(406, 197)
        Me.Controls.Add(Me.txtOpName)
        Me.Name = "FormulaLossMaking"
        Me.Controls.SetChildIndex(Me.txtOpName, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub MaterialLedgerMaking_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Formula Loss Data Making"
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object

        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable = Nothing
        Dim intStatus As Integer = 0
        Dim strStatus As String = ""
        Dim strSql As String = ""
        Dim objDataMaking As DataListMaking
        Dim strOpName As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListMaking(strOpName)

            If strStatus = 1 Then
                MyBase.ShowMsg("E038")
                Me.Cursor = Cursors.Default
                Exit Function
            ElseIf strStatus = 2 Then
                MyBase.ShowMsg("E039")
                Me.Cursor = Cursors.Default
                Exit Function
            ElseIf strStatus = 3 Then
                MyBase.ShowMsg("E040")
                Me.Cursor = Cursors.Default
                Exit Function
            ElseIf strStatus = 4 Then
                MyBase.ShowMsg("E041")
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("Data Making Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try


    End Function

    Public Function funDataListMaking(ByVal strOpName As String) As Integer

        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim WkMaterialCode As String = ""
        Dim WkDescript As String = ""
        Dim WkFinalStkQty As Decimal = 0
        Dim WkRcvQty As Decimal = 0
        Dim WkIssQty As Decimal = 0
        strOpName = MyBase.[Operator]()

        Try
            ' make TableTest class instance and reference WAInvoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            'Data Making process
            strSql = "Exec sprFormulaLossListMake '" & strOpName & "'"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            'Update OperatorName
            'objDataTbl = objTbl.UpdateFormulaLossList(strOpName)

            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource

            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If

        End Try


    End Function


End Class
#End Region

#Region "FormulaLossListGrpAmountPrinting"
Public Class FormulaLossListGrpAmountPrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Formula Loss Amount Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E002")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListGrpAmount")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossAmount
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "SELECT  * From vewFormulaLossListGrp"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            strSql2 = "SELECT * FROM  vewFormulaLossListGrpAmount"
            objDataTbl2 = objTbl.GetDataTableCnt(strSql2)

            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    dtData2 = objDataTbl2.Clone
                    For Each drDataList2 In objDataTbl2.Rows
                        drData2 = dtData2.NewRow
                        drData2.ItemArray = drDataList2.ItemArray
                        dtData2.Rows.Add(drData2)
                    Next

                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "FormulaLossList"

                    dsMatLedger.Tables.Add(dtData2)
                    dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"
                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------

                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "FormulaLossListGrpAmountRe-Printing"
Public Class FormulaLossListGrpAmountRePrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Formula Loss Amount Re-Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E002")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListGrpAmount")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossAmount
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn



            ' Data check Proces
            strSql = "SELECT  * From vewFormulaLossListGrpHist Where OpMonth = '" & VarOPMonth & "' "
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            strSql2 = "SELECT * FROM  vewFormulaLossListGrpAmountHist Where OpMonth = '" & VarOPMonth & "'"
            objDataTbl2 = objTbl.GetDataTableCnt(strSql2)

            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    dtData2 = objDataTbl2.Clone
                    For Each drDataList2 In objDataTbl2.Rows
                        drData2 = dtData2.NewRow
                        drData2.ItemArray = drDataList2.ItemArray
                        dtData2.Rows.Add(drData2)
                    Next

                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "FormulaLossList"

                    dsMatLedger.Tables.Add(dtData2)
                    dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"
                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------

                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "FormulaLossListGrpUnitPrinting"
Public Class FormulaLossListGrpUnitPrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Formula Loss Unit Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListGrpUnit")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossUnit
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "SELECT * FROM  vewFormulaLossListGrp"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            strSql2 = "SELECT *  FROM  vewFormulaLossListGrpUnit "
            objDataTbl2 = objTbl.GetDataTableCnt(strSql2)

            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    dtData2 = objDataTbl2.Clone
                    For Each drDataList2 In objDataTbl2.Rows
                        drData2 = dtData2.NewRow
                        drData2.ItemArray = drDataList2.ItemArray
                        dtData2.Rows.Add(drData2)
                    Next

                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "FormulaLossList"

                    dsMatLedger.Tables.Add(dtData2)
                    dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"
                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------

                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "FormulaLossListGrpUnitRe-Printing"
Public Class FormulaLossListGrpUnitRePrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Formula Loss Unit Re-Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListGrpUnit")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossUnit
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "SELECT * FROM  vewFormulaLossListGrpHist Where OpMonth = '" & VarOPMonth & "'"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            strSql2 = "SELECT *  FROM  vewFormulaLossListGrpUnitHist Where OpMonth = '" & VarOPMonth & "'"
            objDataTbl2 = objTbl.GetDataTableCnt(strSql2)

            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    dtData2 = objDataTbl2.Clone
                    For Each drDataList2 In objDataTbl2.Rows
                        drData2 = dtData2.NewRow
                        drData2.ItemArray = drDataList2.ItemArray
                        dtData2.Rows.Add(drData2)
                    Next

                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "FormulaLossList"

                    dsMatLedger.Tables.Add(dtData2)
                    dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"
                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------

                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "SummaryOfUsePrinting"
Public Class SummaryOfUsePrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Summary Of Use Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListSumUse")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossSumUse
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "EXEC sprSummaryUseListAllPrint 'xxxxxx','0' "
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            'strSql2 = "SELECT * FROM  vewFormulaLossListGrpAmount"
            'objDataTbl2 = objTbl.GetDataTableCnt(strSql2)


            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    'dtData2 = objDataTbl2.Clone
                    'For Each drDataList2 In objDataTbl2.Rows
                    '    drData2 = dtData2.NewRow
                    '    drData2.ItemArray = drDataList2.ItemArray
                    '    dtData2.Rows.Add(drData2)
                    'Next


                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "SumUse0192C"

                    'dsMatLedger.Tables.Add(dtData2)
                    'dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"

                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------



                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "SummaryOfUseRe-Printing"
Public Class SummaryOfUseRePrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Summary Of Use Re-Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListSumUse")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossSumUse
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "EXEC sprSummaryUseListAllPrint '" & VarOPMonth & "','1' "
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            'strSql2 = "SELECT * FROM  vewFormulaLossListGrpAmount"
            'objDataTbl2 = objTbl.GetDataTableCnt(strSql2)


            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    'dtData2 = objDataTbl2.Clone
                    'For Each drDataList2 In objDataTbl2.Rows
                    '    drData2 = dtData2.NewRow
                    '    drData2.ItemArray = drDataList2.ItemArray
                    '    dtData2.Rows.Add(drData2)
                    'Next


                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "SumUse0192C"

                    'dsMatLedger.Tables.Add(dtData2)
                    'dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"

                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------



                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "SummaryOfRatePrinting"
Public Class SummaryOfRatePrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Summary Of Rate Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListSumRate")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossSumRate
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "EXEC sprSummaryRateListAllPrint 'xxxxxx','0'"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            strSql2 = "EXEC sprSummaryRateTotalAllPrint 'xxxxxx','0'"
            objDataTbl2 = objTbl.GetDataTableCnt(strSql2)


            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    dtData2 = objDataTbl2.Clone
                    For Each drDataList2 In objDataTbl2.Rows
                        drData2 = dtData2.NewRow
                        drData2.ItemArray = drDataList2.ItemArray
                        dtData2.Rows.Add(drData2)
                    Next


                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "SumUse0192C"

                    dsMatLedger.Tables.Add(dtData2)
                    dsMatLedger.Tables(1).TableName = "SumAverage"

                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------



                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "SummaryOfRateRe-Printing"
Public Class SummaryOfRateRePrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Summary Of Rate Re-Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListSumRate")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossSumRate
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "EXEC sprSummaryRateListAllPrint '" & VarOPMonth & "','1'"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            strSql2 = "EXEC sprSummaryRateTotalAllPrint '" & VarOPMonth & "','1'"
            objDataTbl2 = objTbl.GetDataTableCnt(strSql2)


            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    dtData2 = objDataTbl2.Clone
                    For Each drDataList2 In objDataTbl2.Rows
                        drData2 = dtData2.NewRow
                        drData2.ItemArray = drDataList2.ItemArray
                        dtData2.Rows.Add(drData2)
                    Next


                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "SumUse0192C"

                    dsMatLedger.Tables.Add(dtData2)
                    dsMatLedger.Tables(1).TableName = "SumAverage"

                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------



                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "SummaryOfAmountPrinting"
Public Class SummaryOfAmountPrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Summary Of Amount Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListSumUse")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossSumAmount
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "EXEC sprSummaryUseListAllPrint 'xxxxxx','0'"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            'strSql2 = "SELECT * FROM  vewFormulaLossListGrpAmount"
            'objDataTbl2 = objTbl.GetDataTableCnt(strSql2)


            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    'dtData2 = objDataTbl2.Clone
                    'For Each drDataList2 In objDataTbl2.Rows
                    '    drData2 = dtData2.NewRow
                    '    drData2.ItemArray = drDataList2.ItemArray
                    '    dtData2.Rows.Add(drData2)
                    'Next


                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "SumUse0192C"

                    'dsMatLedger.Tables.Add(dtData2)
                    'dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"

                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------



                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "SummaryOfAmountRe-Printing"
Public Class SummaryOfAmountRePrinting
    Inherits Rist.OPMCmnClass.DialogControl


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MaterialLedgerPrinting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Summary Of Amount Re-Printing"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object
        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        'Dim strMsg As String
        'Dim xx As String

        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataListPrinting()

            If strStatus = 0 Then
                MyBase.ShowMsg("E003")                      'Error : Data not found on IssuancePur
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            Me.Cursor = Cursors.Default
            Me.SetMessage("List Printing Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataListPrinting() As Integer
        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objDataTbl2 As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        Dim strSql2 As String
        Dim frmReport As New FrmPrintPreview
        Dim dsMatLedger As New DataSet("FormulaLossListSumUse")
        Dim drDataList As DataRow
        Dim drDataList2 As DataRow
        Dim dtData As DataTable
        Dim dtData2 As DataTable
        Dim drData As DataRow
        Dim drData2 As DataRow
        Dim Report As New FormulaLossSumAmount
        Try
            ' make TableTest class instance and reference Invoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            ' Data check Proces
            strSql = "EXEC sprSummaryUseListAllPrint '" & VarOPMonth & "','1'"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

            'strSql2 = "SELECT * FROM  vewFormulaLossListGrpAmount"
            'objDataTbl2 = objTbl.GetDataTableCnt(strSql2)


            If Not objDataTbl Is Nothing Then
                If objDataTbl.Rows.Count > 0 Then

                    dtData = objDataTbl.Clone
                    For Each drDataList In objDataTbl.Rows
                        drData = dtData.NewRow
                        drData.ItemArray = drDataList.ItemArray
                        dtData.Rows.Add(drData)
                    Next
                    '--------Start New-------------------------------------------------------
                    'dtData2 = objDataTbl2.Clone
                    'For Each drDataList2 In objDataTbl2.Rows
                    '    drData2 = dtData2.NewRow
                    '    drData2.ItemArray = drDataList2.ItemArray
                    '    dtData2.Rows.Add(drData2)
                    'Next


                    dsMatLedger.Tables.Add(dtData)
                    dsMatLedger.Tables(0).TableName = "SumUse0192C"

                    'dsMatLedger.Tables.Add(dtData2)
                    'dsMatLedger.Tables(1).TableName = "FormulaLossListGrp"

                    Report.SetDataSource(dsMatLedger)
                    '--------End New --------------------------------------------------------



                    frmReport.ReportSource = Report
                    frmReport.Preview()
                    ''Report.PrintToPrinter(1, True, 0, 0)

                Else
                    strRet = 0
                    Return strRet
                    Exit Function
                End If
            Else
                strRet = 0
                Return strRet
                Exit Function
            End If
            strRet = 1
            Return strRet

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function

End Class
#End Region

#Region "MonthlyClosingMaking"
Public Class MonthlyClosingMaking
    Inherits Rist.OPMCmnClass.DialogControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constractor
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        MyBase.New(pobjUserInfo)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub MonthlyClosingMaking_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Title = "Monthly Closing Making"
            Me.TitleFontSize = 18
            Me.SetMessage("Please Click [Start] to Run Process")

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function RunBatch() As Object

        Dim strUserName As String = SystemInformation.UserName
        Dim strTable As DataTable
        Dim intStatus As Integer
        Dim strStatus As String = ""
        Dim strSql As String
        Dim objDataMaking As DataListMaking
        Dim strOpName As String

        'Export to Excel Function (MaterialLedgerList Report)

        Me.SetMessage("Monthly Closing Processing...")
        Me.Refresh()


        Try
            objDataMaking = New DataListMaking
            objDataMaking.Conn = Me.objDBBase.Conn
            Me.Cursor = Cursors.WaitCursor

            strStatus = funDataMonthlyClosingMaking(strOpName)

            Me.Cursor = Cursors.Default
            Me.SetMessage("Data Making Successful")


        Catch ex As Exception
            'MsgBox("System Error please contact ..IS")
            MessageBox.Show("System Error! Please contact IS", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.SetMessage(ex.Message)
        Finally
        End Try
    End Function

    Public Function funDataMonthlyClosingMaking(ByVal strOpName As String) As Integer

        Dim strRet As Integer
        Dim objDataTbl As DataTable = Nothing
        Dim objTbl As DataListMaking
        Dim strSql As String
        strOpName = MyBase.[Operator]()

        Try

            ' make TableTest class instance and reference WAInvoice connection object
            objTbl = New DataListMaking
            objTbl.Conn = Me.objDBBase.Conn

            'Data Making process
            strSql = "Exec sprMonthlyClosingMake  '" & strOpName & "'"
            objDataTbl = objTbl.GetDataTableCnt(strSql)

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objTbl Is Nothing Then
                objTbl = Nothing
            End If
        End Try

        Return strRet
    End Function


End Class


#End Region