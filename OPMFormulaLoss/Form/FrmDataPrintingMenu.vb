Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass
Public Class FrmDataPrintingMenu
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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents CmbOPMonth As System.Windows.Forms.ComboBox
    Friend WithEvents lblSumRateRP As System.Windows.Forms.Label
    Friend WithEvents lblFormulaLossUnitRP As System.Windows.Forms.Label
    Friend WithEvents lblFormulaLossAmountRP As System.Windows.Forms.Label
    Friend WithEvents lblSumUseRP As System.Windows.Forms.Label
    Friend WithEvents lblSumAmountRP As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.KeyControl1 = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.lblSumRateRP = New System.Windows.Forms.Label()
        Me.lblFormulaLossUnitRP = New System.Windows.Forms.Label()
        Me.lblFormulaLossAmountRP = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblSumAmountRP = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.CmbOPMonth = New System.Windows.Forms.ComboBox()
        Me.lblSumUseRP = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblSumRateRP
        '
        Me.lblSumRateRP.BackColor = System.Drawing.Color.LightGreen
        Me.lblSumRateRP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSumRateRP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSumRateRP.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSumRateRP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSumRateRP.Location = New System.Drawing.Point(48, 320)
        Me.lblSumRateRP.Name = "lblSumRateRP"
        Me.lblSumRateRP.Size = New System.Drawing.Size(368, 64)
        Me.lblSumRateRP.TabIndex = 6
        Me.lblSumRateRP.Text = "4.Summary Rate"
        Me.lblSumRateRP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFormulaLossUnitRP
        '
        Me.lblFormulaLossUnitRP.BackColor = System.Drawing.Color.LightGreen
        Me.lblFormulaLossUnitRP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFormulaLossUnitRP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblFormulaLossUnitRP.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormulaLossUnitRP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFormulaLossUnitRP.Location = New System.Drawing.Point(48, 164)
        Me.lblFormulaLossUnitRP.Name = "lblFormulaLossUnitRP"
        Me.lblFormulaLossUnitRP.Size = New System.Drawing.Size(368, 64)
        Me.lblFormulaLossUnitRP.TabIndex = 4
        Me.lblFormulaLossUnitRP.Text = "2.Formula Loss Unit"
        Me.lblFormulaLossUnitRP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFormulaLossAmountRP
        '
        Me.lblFormulaLossAmountRP.BackColor = System.Drawing.Color.LightGreen
        Me.lblFormulaLossAmountRP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFormulaLossAmountRP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblFormulaLossAmountRP.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormulaLossAmountRP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFormulaLossAmountRP.Location = New System.Drawing.Point(48, 88)
        Me.lblFormulaLossAmountRP.Name = "lblFormulaLossAmountRP"
        Me.lblFormulaLossAmountRP.Size = New System.Drawing.Size(368, 64)
        Me.lblFormulaLossAmountRP.TabIndex = 5
        Me.lblFormulaLossAmountRP.Text = "1.Formula Loss Amount"
        Me.lblFormulaLossAmountRP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.LightCyan
        Me.GroupBox2.Controls.Add(Me.lblSumAmountRP)
        Me.GroupBox2.Controls.Add(Me.lblType)
        Me.GroupBox2.Controls.Add(Me.CmbOPMonth)
        Me.GroupBox2.Controls.Add(Me.lblSumUseRP)
        Me.GroupBox2.Controls.Add(Me.lblSumRateRP)
        Me.GroupBox2.Controls.Add(Me.lblFormulaLossUnitRP)
        Me.GroupBox2.Controls.Add(Me.lblFormulaLossAmountRP)
        Me.GroupBox2.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.SeaGreen
        Me.GroupBox2.Location = New System.Drawing.Point(216, 112)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(464, 485)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "List Re-Printing"
        '
        'lblSumAmountRP
        '
        Me.lblSumAmountRP.BackColor = System.Drawing.Color.LightGreen
        Me.lblSumAmountRP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSumAmountRP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSumAmountRP.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSumAmountRP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSumAmountRP.Location = New System.Drawing.Point(48, 398)
        Me.lblSumAmountRP.Name = "lblSumAmountRP"
        Me.lblSumAmountRP.Size = New System.Drawing.Size(368, 64)
        Me.lblSumAmountRP.TabIndex = 88
        Me.lblSumAmountRP.Text = "5.Summary Amount"
        Me.lblSumAmountRP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblType
        '
        Me.lblType.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.lblType.ForeColor = System.Drawing.Color.Black
        Me.lblType.Location = New System.Drawing.Point(40, 42)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(80, 24)
        Me.lblType.TabIndex = 87
        Me.lblType.Text = "OPMonth"
        '
        'CmbOPMonth
        '
        Me.CmbOPMonth.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbOPMonth.Location = New System.Drawing.Point(120, 40)
        Me.CmbOPMonth.Name = "CmbOPMonth"
        Me.CmbOPMonth.Size = New System.Drawing.Size(96, 26)
        Me.CmbOPMonth.TabIndex = 8
        '
        'lblSumUseRP
        '
        Me.lblSumUseRP.BackColor = System.Drawing.Color.LightGreen
        Me.lblSumUseRP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSumUseRP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSumUseRP.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSumUseRP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSumUseRP.Location = New System.Drawing.Point(48, 243)
        Me.lblSumUseRP.Name = "lblSumUseRP"
        Me.lblSumUseRP.Size = New System.Drawing.Size(368, 64)
        Me.lblSumUseRP.TabIndex = 7
        Me.lblSumUseRP.Text = "3.Summary Use"
        Me.lblSumUseRP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmDataPrintingMenu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(1013, 711)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "FrmDataPrintingMenu"
        Me.Controls.SetChildIndex(Me.lblFkey12, 0)
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FrmMainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Private Sub FrmMainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            MyBase.Title = MyBase.GetItemName("1006")
            MyBase.TitleFontSize = 20
            Me.CloseCaption = "F12:" & MyBase.GetItemName("0009")
            MyBase.IsErrMsg = True
            MyBase.Message = ""
            Me.ShowInTaskbar = False
            Call SetLabel()
            Call funGetCmbOPMonth()
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub FrmMainMenu_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            If Not ActForm Is Nothing Then
                ActForm.Activate()
            End If

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub



    Private Sub FrmMainMenu_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Try
            Me.KeyControl1.Push(e.KeyValue)
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Function SetLabel()
        Dim objCMT As CmnMasterTable
        Try
            objCMT = New CmnMasterTable(MyBase.SystemName)
            objCMT.Conn = MyBase.objDBBase.Conn

            Me.lblFormulaLossAmountRP.Text = objCMT.GetName("BOTM", "0005")
            Me.lblFormulaLossUnitRP.Text = objCMT.GetName("BOTM", "0006")
            Me.lblSumUseRP.Text = objCMT.GetName("BOTM", "0007")
            Me.lblSumRateRP.Text = objCMT.GetName("BOTM", "0010")
            Me.lblSumAmountRP.Text = objCMT.GetName("BOTM", "0011")

        Catch ex As Exception
            Throw ex
        Finally
            If Not objCMT Is Nothing Then
                objCMT = Nothing
            End If
        End Try
    End Function

    Public Overrides Function PushF12() As Object
        Close()
    End Function

    Private Sub funGetCmbOPMonth()
        Dim dtTable As DataTable = Nothing
        Dim objOPMonth As DBConnect
        Try

            objOPMonth = New DBConnect
            objOPMonth.Conn = MyBase.objDBBase.Conn
            dtTable = objOPMonth.GetOpMonthTbl2()

            For Each drRow As DataRow In dtTable.Rows
                Me.CmbOPMonth.Items.Add(drRow("OPMonth"))
            Next
            dtTable.Dispose()

        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg(ex.MsgCode)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub CmbOPMonth_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CmbOPMonth.KeyUp
        Try
            If e.KeyValue = Keys.Enter Then
                If funOpMonthChecking() = False Then
                    Me.CmbOPMonth.Focus()
                    MyBase.IsErrMsg = True
                    MyBase.ShowMsg("E056")
                    Exit Sub
                Else
                    MyBase.Message = ""
                    VarOPMonth = Me.CmbOPMonth.Text
                End If
            End If

        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg(ex.MsgCode)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub CmbOPMonth_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbOPMonth.SelectedIndexChanged
        Try

            If funOpMonthChecking() = False Then
                Me.CmbOPMonth.Focus()
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E056")
                Exit Sub
            Else
                MyBase.Message = ""
                VarOPMonth = Me.CmbOPMonth.Text
            End If

        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg(ex.MsgCode)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    Function funOpMonthChecking() As Boolean

        Dim objOpMonth As DBConnect
        Dim dtData As DataTable = Nothing

        Try
            '    ' make new DBConnect instance
            objOpMonth = New DBConnect
            objOpMonth.Conn = MyBase.objDBBase.Conn
            dtData = objOpMonth.GetOpMonthTbl3(CmbOPMonth.Text)
            If Not dtData Is Nothing And dtData.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As CustomErrException
            Throw ex

        Catch ex As Exception
            Throw ex

        End Try

    End Function

    Private Sub lblFormulaLossAmountRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFormulaLossAmountRP.Click

        Try
            ActForm = New FormulaLossListGrpAmountRePrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try

    End Sub
    Private Sub lblFormulaLossUnitRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFormulaLossUnitRP.Click
        Try
            ActForm = New FormulaLossListGrpUnitRePrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub
    Private Sub lblSumUseRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSumUseRP.Click
        Try
            ActForm = New SummaryOfUseRePrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub
    Private Sub lblSumRateRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSumRateRP.Click
        Try
            ActForm = New SummaryOfRateRePrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub
    Private Sub lblSumAmountRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSumAmountRP.Click
        Try
            ActForm = New SummaryOfAmountRePrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub


End Class