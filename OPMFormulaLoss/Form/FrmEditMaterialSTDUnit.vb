Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass

Public Class FrmEditMaterialSTDUnit
    Inherits Rist.OPMCmnClass.PageBase

    'Inherits Rist.OPMCmnClass.DialogControl

    Private ActForm As Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        'MyBase.New()

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
    Friend WithEvents KeyControl1 As Rist.OPMCmnClass.KeyControl
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.KeyControl1 = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.dgvDataDetail = New System.Windows.Forms.DataGridView()
        Me.lblMaterialCode = New System.Windows.Forms.Label()
        Me.txtMaterialCode = New System.Windows.Forms.TextBox()
        Me.lblMaterialName = New System.Windows.Forms.Label()
        Me.txtMaterialName = New System.Windows.Forms.TextBox()
        Me.lblSpec = New System.Windows.Forms.Label()
        Me.txtSpec = New System.Windows.Forms.TextBox()
        Me.lblUnit = New System.Windows.Forms.Label()
        Me.txtUnit = New System.Windows.Forms.TextBox()
        Me.lblProcessCode = New System.Windows.Forms.Label()
        Me.txtProcessCode = New System.Windows.Forms.TextBox()
        Me.lblMaterialRatio = New System.Windows.Forms.Label()
        Me.txtMaterialRatio = New System.Windows.Forms.TextBox()
        Me.lblStandardUnit = New System.Windows.Forms.Label()
        Me.txtStandardUnit = New System.Windows.Forms.TextBox()
        Me.gbMaterialUse = New System.Windows.Forms.GroupBox()
        Me.cb01C5C = New System.Windows.Forms.CheckBox()
        Me.cb0192CW11 = New System.Windows.Forms.CheckBox()
        Me.cb0192C = New System.Windows.Forms.CheckBox()
        Me.lblKeyF1 = New Rist.OPMCmnClass.FunctionKeyControl()
        Me.lblKeyF2 = New Rist.OPMCmnClass.FunctionKeyControl()
        Me.lblKeyF3 = New Rist.OPMCmnClass.FunctionKeyControl()
        Me.lblKeyF5 = New Rist.OPMCmnClass.FunctionKeyControl()
        CType(Me.dgvDataDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbMaterialUse.SuspendLayout()
        Me.SuspendLayout()
        '
        'KeyControl1
        '
        '
        'dgvDataDetail
        '
        Me.dgvDataDetail.AllowUserToAddRows = False
        Me.dgvDataDetail.AllowUserToDeleteRows = False
        Me.dgvDataDetail.AllowUserToOrderColumns = True
        Me.dgvDataDetail.AllowUserToResizeRows = False
        Me.dgvDataDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDataDetail.Location = New System.Drawing.Point(52, 211)
        Me.dgvDataDetail.Name = "dgvDataDetail"
        Me.dgvDataDetail.ReadOnly = True
        Me.dgvDataDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDataDetail.Size = New System.Drawing.Size(900, 400)
        Me.dgvDataDetail.TabIndex = 3
        '
        'lblMaterialCode
        '
        Me.lblMaterialCode.AutoSize = True
        Me.lblMaterialCode.Location = New System.Drawing.Point(76, 114)
        Me.lblMaterialCode.Name = "lblMaterialCode"
        Me.lblMaterialCode.Size = New System.Drawing.Size(84, 13)
        Me.lblMaterialCode.TabIndex = 4
        Me.lblMaterialCode.Text = "MaterialCode:"
        '
        'txtMaterialCode
        '
        Me.txtMaterialCode.Location = New System.Drawing.Point(166, 111)
        Me.txtMaterialCode.Name = "txtMaterialCode"
        Me.txtMaterialCode.Size = New System.Drawing.Size(100, 20)
        Me.txtMaterialCode.TabIndex = 5
        '
        'lblMaterialName
        '
        Me.lblMaterialName.AutoSize = True
        Me.lblMaterialName.Location = New System.Drawing.Point(76, 138)
        Me.lblMaterialName.Name = "lblMaterialName"
        Me.lblMaterialName.Size = New System.Drawing.Size(86, 13)
        Me.lblMaterialName.TabIndex = 6
        Me.lblMaterialName.Text = "MaterialName:"
        '
        'txtMaterialName
        '
        Me.txtMaterialName.Location = New System.Drawing.Point(166, 135)
        Me.txtMaterialName.Name = "txtMaterialName"
        Me.txtMaterialName.Size = New System.Drawing.Size(100, 20)
        Me.txtMaterialName.TabIndex = 7
        '
        'lblSpec
        '
        Me.lblSpec.AutoSize = True
        Me.lblSpec.Location = New System.Drawing.Point(122, 162)
        Me.lblSpec.Name = "lblSpec"
        Me.lblSpec.Size = New System.Drawing.Size(38, 13)
        Me.lblSpec.TabIndex = 8
        Me.lblSpec.Text = "Spec:"
        '
        'txtSpec
        '
        Me.txtSpec.Location = New System.Drawing.Point(166, 159)
        Me.txtSpec.Name = "txtSpec"
        Me.txtSpec.Size = New System.Drawing.Size(100, 20)
        Me.txtSpec.TabIndex = 9
        '
        'lblUnit
        '
        Me.lblUnit.AutoSize = True
        Me.lblUnit.Location = New System.Drawing.Point(127, 186)
        Me.lblUnit.Name = "lblUnit"
        Me.lblUnit.Size = New System.Drawing.Size(33, 13)
        Me.lblUnit.TabIndex = 10
        Me.lblUnit.Text = "Unit:"
        '
        'txtUnit
        '
        Me.txtUnit.Location = New System.Drawing.Point(166, 184)
        Me.txtUnit.Name = "txtUnit"
        Me.txtUnit.Size = New System.Drawing.Size(100, 20)
        Me.txtUnit.TabIndex = 11
        '
        'lblProcessCode
        '
        Me.lblProcessCode.AutoSize = True
        Me.lblProcessCode.Location = New System.Drawing.Point(355, 118)
        Me.lblProcessCode.Name = "lblProcessCode"
        Me.lblProcessCode.Size = New System.Drawing.Size(85, 13)
        Me.lblProcessCode.TabIndex = 12
        Me.lblProcessCode.Text = "ProcessCode:"
        '
        'txtProcessCode
        '
        Me.txtProcessCode.Location = New System.Drawing.Point(446, 115)
        Me.txtProcessCode.Name = "txtProcessCode"
        Me.txtProcessCode.Size = New System.Drawing.Size(100, 20)
        Me.txtProcessCode.TabIndex = 13
        '
        'lblMaterialRatio
        '
        Me.lblMaterialRatio.AutoSize = True
        Me.lblMaterialRatio.Location = New System.Drawing.Point(356, 144)
        Me.lblMaterialRatio.Name = "lblMaterialRatio"
        Me.lblMaterialRatio.Size = New System.Drawing.Size(84, 13)
        Me.lblMaterialRatio.TabIndex = 14
        Me.lblMaterialRatio.Text = "MaterialRatio:"
        '
        'txtMaterialRatio
        '
        Me.txtMaterialRatio.Location = New System.Drawing.Point(446, 141)
        Me.txtMaterialRatio.Name = "txtMaterialRatio"
        Me.txtMaterialRatio.Size = New System.Drawing.Size(100, 20)
        Me.txtMaterialRatio.TabIndex = 15
        '
        'lblStandardUnit
        '
        Me.lblStandardUnit.AutoSize = True
        Me.lblStandardUnit.Location = New System.Drawing.Point(358, 169)
        Me.lblStandardUnit.Name = "lblStandardUnit"
        Me.lblStandardUnit.Size = New System.Drawing.Size(82, 13)
        Me.lblStandardUnit.TabIndex = 16
        Me.lblStandardUnit.Text = "StandardUnit:"
        '
        'txtStandardUnit
        '
        Me.txtStandardUnit.Location = New System.Drawing.Point(446, 167)
        Me.txtStandardUnit.Name = "txtStandardUnit"
        Me.txtStandardUnit.Size = New System.Drawing.Size(100, 20)
        Me.txtStandardUnit.TabIndex = 17
        '
        'gbMaterialUse
        '
        Me.gbMaterialUse.Controls.Add(Me.cb01C5C)
        Me.gbMaterialUse.Controls.Add(Me.cb0192CW11)
        Me.gbMaterialUse.Controls.Add(Me.cb0192C)
        Me.gbMaterialUse.Location = New System.Drawing.Point(587, 114)
        Me.gbMaterialUse.Name = "gbMaterialUse"
        Me.gbMaterialUse.Size = New System.Drawing.Size(260, 70)
        Me.gbMaterialUse.TabIndex = 20
        Me.gbMaterialUse.TabStop = False
        Me.gbMaterialUse.Text = "MaterialUse"
        '
        'cb01C5C
        '
        Me.cb01C5C.AutoSize = True
        Me.cb01C5C.Location = New System.Drawing.Point(175, 29)
        Me.cb01C5C.Name = "cb01C5C"
        Me.cb01C5C.Size = New System.Drawing.Size(65, 17)
        Me.cb01C5C.TabIndex = 2
        Me.cb01C5C.Text = "01C5C"
        Me.cb01C5C.UseVisualStyleBackColor = True
        '
        'cb0192CW11
        '
        Me.cb0192CW11.AutoSize = True
        Me.cb0192CW11.Location = New System.Drawing.Point(82, 29)
        Me.cb0192CW11.Name = "cb0192CW11"
        Me.cb0192CW11.Size = New System.Drawing.Size(87, 17)
        Me.cb0192CW11.TabIndex = 1
        Me.cb0192CW11.Text = "0192CW11"
        Me.cb0192CW11.UseVisualStyleBackColor = True
        '
        'cb0192C
        '
        Me.cb0192C.AutoSize = True
        Me.cb0192C.Location = New System.Drawing.Point(13, 30)
        Me.cb0192C.Name = "cb0192C"
        Me.cb0192C.Size = New System.Drawing.Size(63, 17)
        Me.cb0192C.TabIndex = 0
        Me.cb0192C.Text = "0192C"
        Me.cb0192C.UseVisualStyleBackColor = True
        '
        'lblKeyF1
        '
        Me.lblKeyF1.Location = New System.Drawing.Point(53, 656)
        Me.lblKeyF1.Name = "lblKeyF1"
        Me.lblKeyF1.Size = New System.Drawing.Size(152, 48)
        Me.lblKeyF1.TabIndex = 21
        '
        'lblKeyF2
        '
        Me.lblKeyF2.Location = New System.Drawing.Point(211, 656)
        Me.lblKeyF2.Name = "lblKeyF2"
        Me.lblKeyF2.Size = New System.Drawing.Size(152, 48)
        Me.lblKeyF2.TabIndex = 22
        '
        'lblKeyF3
        '
        Me.lblKeyF3.Location = New System.Drawing.Point(369, 656)
        Me.lblKeyF3.Name = "lblKeyF3"
        Me.lblKeyF3.Size = New System.Drawing.Size(152, 48)
        Me.lblKeyF3.TabIndex = 23
        '
        'lblKeyF5
        '
        Me.lblKeyF5.Location = New System.Drawing.Point(527, 656)
        Me.lblKeyF5.Name = "lblKeyF5"
        Me.lblKeyF5.Size = New System.Drawing.Size(152, 48)
        Me.lblKeyF5.TabIndex = 24
        '
        'FrmEditMaterialSTDUnit
        '
        Me.AccessibleName = ""
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(1013, 711)
        Me.Controls.Add(Me.lblKeyF5)
        Me.Controls.Add(Me.lblKeyF3)
        Me.Controls.Add(Me.lblKeyF2)
        Me.Controls.Add(Me.lblKeyF1)
        Me.Controls.Add(Me.gbMaterialUse)
        Me.Controls.Add(Me.txtStandardUnit)
        Me.Controls.Add(Me.lblStandardUnit)
        Me.Controls.Add(Me.txtMaterialRatio)
        Me.Controls.Add(Me.lblMaterialRatio)
        Me.Controls.Add(Me.txtProcessCode)
        Me.Controls.Add(Me.lblProcessCode)
        Me.Controls.Add(Me.txtUnit)
        Me.Controls.Add(Me.lblUnit)
        Me.Controls.Add(Me.txtSpec)
        Me.Controls.Add(Me.lblSpec)
        Me.Controls.Add(Me.txtMaterialName)
        Me.Controls.Add(Me.lblMaterialName)
        Me.Controls.Add(Me.txtMaterialCode)
        Me.Controls.Add(Me.lblMaterialCode)
        Me.Controls.Add(Me.dgvDataDetail)
        Me.Name = "FrmEditMaterialSTDUnit"
        Me.Controls.SetChildIndex(Me.dgvDataDetail, 0)
        Me.Controls.SetChildIndex(Me.lblMaterialCode, 0)
        Me.Controls.SetChildIndex(Me.txtMaterialCode, 0)
        Me.Controls.SetChildIndex(Me.lblMaterialName, 0)
        Me.Controls.SetChildIndex(Me.txtMaterialName, 0)
        Me.Controls.SetChildIndex(Me.lblSpec, 0)
        Me.Controls.SetChildIndex(Me.txtSpec, 0)
        Me.Controls.SetChildIndex(Me.lblUnit, 0)
        Me.Controls.SetChildIndex(Me.txtUnit, 0)
        Me.Controls.SetChildIndex(Me.lblProcessCode, 0)
        Me.Controls.SetChildIndex(Me.txtProcessCode, 0)
        Me.Controls.SetChildIndex(Me.lblMaterialRatio, 0)
        Me.Controls.SetChildIndex(Me.txtMaterialRatio, 0)
        Me.Controls.SetChildIndex(Me.lblStandardUnit, 0)
        Me.Controls.SetChildIndex(Me.txtStandardUnit, 0)
        Me.Controls.SetChildIndex(Me.gbMaterialUse, 0)
        Me.Controls.SetChildIndex(Me.lblFkey12, 0)
        Me.Controls.SetChildIndex(Me.lblKeyF1, 0)
        Me.Controls.SetChildIndex(Me.lblKeyF2, 0)
        Me.Controls.SetChildIndex(Me.lblKeyF3, 0)
        Me.Controls.SetChildIndex(Me.lblKeyF5, 0)
        CType(Me.dgvDataDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbMaterialUse.ResumeLayout(False)
        Me.gbMaterialUse.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub FrmMainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            MyBase.Title = MyBase.GetItemName("1002")
            MyBase.TitleFontSize = 20
            Me.lblKeyF1.Caption = "F1:" + MyBase.GetButtonName("0012")
            Me.lblKeyF2.Caption = "F2:" + MyBase.GetButtonName("0013")
            Me.lblKeyF3.Caption = "F3:" + MyBase.GetButtonName("0014")
            Me.lblKeyF5.Caption = "F5:" + MyBase.GetButtonName("0015")
            Me.CloseCaption = "F12:" & MyBase.GetButtonName("0016")
            MyBase.IsErrMsg = True
            MyBase.Message = ""
            Me.ShowInTaskbar = False

            Call RefreshForm()
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message

        End Try
    End Sub


    Private Sub lblKeyF2_UCClick() Handles lblKeyF2.UCClick
        Try
            Me.KeyControl1.Push(Keys.F2)
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub KeyControl1_PushF1() Handles KeyControl1.PushF1
        Try
            Call UpdateTable()
        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg(ex.MsgCode)
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub KeyControl1_PushF2() Handles KeyControl1.PushF2
        Try
            Call UpdateTable()
        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg(ex.MsgCode)
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub


    'Tool for get data from dataGridview for prevent error with null value
    Function getDataFormDataGridView(dataGridView As DataGridView, row As String, column As String) As Object
        If Not IsDBNull(dataGridView.Rows(row).Cells(column).Value) Then
            Return dataGridView.Rows(row).Cells(column).Value
        Else
            Return ""
        End If
    End Function

    'Detect mouse click DataGridView for update form detail on current row
    Private Sub dgv_CurrentTable_MouseClick(sender As Object, e As MouseEventArgs) Handles dgvDataDetail.MouseClick
        If Not dgvDataDetail.CurrentRow Is Nothing Then

            Dim index As String

            index = dgvDataDetail.CurrentRow.Index
            txtMaterialCode.Text = getDataFormDataGridView(dgvDataDetail, index, "MaterialCode").Trim
            txtMaterialName.Text = getDataFormDataGridView(dgvDataDetail, index, "MaterialName").Trim
            txtSpec.Text = getDataFormDataGridView(dgvDataDetail, index, "Spec").Trim
            txtUnit.Text = getDataFormDataGridView(dgvDataDetail, index, "Unit").Trim
            txtProcessCode.Text = getDataFormDataGridView(dgvDataDetail, index, "ProcessCode").Trim
            txtMaterialRatio.Text = getDataFormDataGridView(dgvDataDetail, index, "MaterialRatio")
            txtStandardUnit.Text = getDataFormDataGridView(dgvDataDetail, index, "StandardUnit")

            If getDataFormDataGridView(dgvDataDetail, index, "MaterialUse_0192C") = True Then
                cb0192C.Checked = True
            Else
                cb0192C.Checked = False
            End If

            If getDataFormDataGridView(dgvDataDetail, index, "MaterialUse_0192CW11") = True Then
                cb0192CW11.Checked = True
            Else
                cb0192CW11.Checked = False
            End If

            If getDataFormDataGridView(dgvDataDetail, index, "MaterialUse_01C5C") = True Then
                cb01C5C.Checked = True
            Else
                cb01C5C.Checked = False
            End If

        Else
            Call ClearDataForm()
        End If
    End Sub


    Private Sub FrmMainMenu_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Try
            Me.KeyControl1.Push(e.KeyValue)
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub SetLabel()

        'Dim objCMT As CmnMasterTable

        Try
            Me.lblMaterialCode.Text = MyBase.GetItemName("0009") + ": "
            Me.lblMaterialName.Text = MyBase.GetItemName("0010") + ": "
            Me.lblSpec.Text = MyBase.GetItemName("0011") + ": "
            Me.lblUnit.Text = MyBase.GetItemName("0012") + ": "
            Me.lblProcessCode.Text = MyBase.GetItemName("0013") + ": "
            Me.lblMaterialRatio.Text = MyBase.GetItemName("0014") + ": "
            Me.lblStandardUnit.Text = MyBase.GetItemName("0015") + ": "

            Me.gbMaterialUse.Text = GetItemName("0016")
            Me.cb0192C.Text = MyBase.GetItemName("0017")
            Me.cb0192CW11.Text = MyBase.GetItemName("0018")
            Me.cb01C5C.Text = MyBase.GetItemName("0019")


        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Function PushF12() As Object
        Close()
    End Function

    Private Sub LoadMaterialStandardUnit()
        Dim dtMaterialSTDUnit As DataTable

        Try
            Dim dbConnect As DBConnect = New DBConnect
            dbConnect.Conn = MyBase.objDBBase.Conn
            dtMaterialSTDUnit = dbConnect.QueryAll("vewMaterialStandardUnit")
            GridviewDataBind(dgvDataDetail, dtMaterialSTDUnit)


        Catch ex As Exception

        End Try

    End Sub

    Private Sub ClearDataForm()
        txtMaterialCode.Clear()
        txtMaterialName.Clear()
        txtSpec.Clear()
        txtUnit.Clear()
        txtProcessCode.Clear()
        txtMaterialRatio.Clear()
        txtStandardUnit.Clear()

        cb0192C.Checked = False
        cb0192CW11.Checked = False
        cb01C5C.Checked = False
    End Sub

    Private Sub RefreshForm()
        Call SetLabel()
        Call ClearDataForm()
        Call LoadMaterialStandardUnit()
    End Sub




    'Formatting column row number
    Private Sub gv_table_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvDataDetail.CellFormatting
        dgvDataDetail.Rows(e.RowIndex).HeaderCell.Value = CStr(e.RowIndex + 1)
    End Sub


    'Bind data without some column, that we don't need
    Private Sub GridviewDataBind(gridview As DataGridView, dataTable As DataTable)
        gridview.DataSource = dataTable

        For Each column As DataGridViewColumn In gridview.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            If column.Name = "No" Then
                column.Visible = False
            End If
        Next
    End Sub

    'Get Context form from textbox
    Private Sub GetContextForm(parameters As System.Collections.Specialized.StringDictionary)

        If RTrim(txtMaterialCode.Text) <> "" Then
            parameters.Add("MaterialCode", txtMaterialCode.Text.Trim)
        End If

        If RTrim(txtMaterialName.Text) <> "" Then
            parameters.Add("MaterialName", txtMaterialName.Text.Trim)
        End If

        If RTrim(txtSpec.Text) <> "" Then
            parameters.Add("Spec", txtSpec.Text.Trim)
        End If

        If RTrim(txtUnit.Text) <> "" Then
            parameters.Add("Unit", txtUnit.Text.Trim)
        End If

        If RTrim(txtProcessCode.Text) <> "" Then
            parameters.Add("ProcessCode", txtProcessCode.Text.Trim)
        End If

        If RTrim(txtMaterialRatio.Text) <> "" Then
            parameters.Add("MaterialRatio", txtMaterialRatio.Text.Trim)
        End If

        If RTrim(txtStandardUnit.Text) <> "" Then
            parameters.Add("StandardUnit", txtStandardUnit.Text.Trim)
        End If

        If RTrim(cb0192C.Checked.ToString) <> "" Then
            parameters.Add("MaterialUse_0192C", cb0192C.Checked.ToString.Trim)
        End If

        If RTrim(cb0192CW11.Checked.ToString) <> "" Then
            parameters.Add("MaterialUse_0192CW11", cb0192CW11.Checked.ToString.Trim)
        End If

        If RTrim(cb01C5C.Checked.ToString) <> "" Then
            parameters.Add("MaterialUse_01C5C", cb01C5C.Checked.ToString.Trim)
        End If

    End Sub

    'Update row data table in database
    Private Sub UpdateTable()
        If Not dgvDataDetail.CurrentRow Is Nothing Then
            Dim dbConn As DBConnect = New DBConnect
            Dim resultTable As DataTable = Nothing
            Dim index As String
            Dim key As String = Nothing

            ' Generate Parameter for use on storeprocedure
            Dim parameters As New System.Collections.Specialized.StringDictionary

            index = dgvDataDetail.CurrentRow.Index
            key = dgvDataDetail.Rows(index).Cells(0).Value

            parameters.Add("No", key)
            GetContextForm(parameters)


            'Connect db and call function stoprocedure
            Try
                dbConn.Conn = MyBase.objDBBase.Conn
                resultTable = dbConn.ModifyMaterialStandardUnit(ModifyType.Update, parameters)

                If Not resultTable Is Nothing And resultTable.Rows.Count > 0 Then
                    Dim resultMessage As String = resultTable.Rows(0)("ResultMessage").ToString().Trim()
                    Dim resultCode As Integer = resultTable.Rows(0)("ResultCode").ToString().Trim()

                    If resultCode = 0 Then
                        MyBase.IsErrMsg = False
                    Else
                        MyBase.IsErrMsg = True
                    End If

                    MyBase.Message = resultMessage

                Else
                    MyBase.IsErrMsg = False
                    MyBase.Message = MyBase.GetPurposeVal("MESG", "E001")
                End If

            Catch ex As Exception
                MyBase.IsErrMsg = True
                MyBase.Message = ex.Message
                RefreshForm()
            End Try

            RefreshForm()

        End If
    End Sub

    Friend WithEvents dgvDataDetail As System.Windows.Forms.DataGridView
    Friend WithEvents lblMaterialCode As System.Windows.Forms.Label
    Friend WithEvents txtMaterialCode As System.Windows.Forms.TextBox
    Friend WithEvents lblMaterialName As System.Windows.Forms.Label
    Friend WithEvents txtMaterialName As System.Windows.Forms.TextBox
    Friend WithEvents lblSpec As System.Windows.Forms.Label
    Friend WithEvents txtSpec As System.Windows.Forms.TextBox
    Friend WithEvents lblUnit As System.Windows.Forms.Label
    Friend WithEvents txtUnit As System.Windows.Forms.TextBox
    Friend WithEvents lblProcessCode As System.Windows.Forms.Label
    Friend WithEvents txtProcessCode As System.Windows.Forms.TextBox
    Friend WithEvents lblMaterialRatio As System.Windows.Forms.Label
    Friend WithEvents txtMaterialRatio As System.Windows.Forms.TextBox
    Friend WithEvents lblStandardUnit As System.Windows.Forms.Label
    Friend WithEvents txtStandardUnit As System.Windows.Forms.TextBox
    Friend WithEvents gbMaterialUse As System.Windows.Forms.GroupBox
    Friend WithEvents cb01C5C As System.Windows.Forms.CheckBox
    Friend WithEvents cb0192CW11 As System.Windows.Forms.CheckBox
    Friend WithEvents cb0192C As System.Windows.Forms.CheckBox
    Friend WithEvents lblKeyF1 As Rist.OPMCmnClass.FunctionKeyControl
    Friend WithEvents lblKeyF2 As Rist.OPMCmnClass.FunctionKeyControl
    Friend WithEvents lblKeyF3 As Rist.OPMCmnClass.FunctionKeyControl
    Friend WithEvents lblKeyF5 As Rist.OPMCmnClass.FunctionKeyControl


End Class