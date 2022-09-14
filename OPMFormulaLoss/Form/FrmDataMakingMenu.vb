Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass

Public Class FrmDataMakingMenu
    Inherits Rist.OPMCmnClass.PageBase

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
    Friend WithEvents lblFormulaLossMaking As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblFormulaLossAmount As System.Windows.Forms.Label
    Friend WithEvents lblSumUse As System.Windows.Forms.Label
    Friend WithEvents lblFormulaLossUnit As System.Windows.Forms.Label
    Friend WithEvents lblSumAmount As System.Windows.Forms.Label
    Friend WithEvents lblSumRate As System.Windows.Forms.Label

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.KeyControl1 = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.lblFormulaLossMaking = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblSumAmount = New System.Windows.Forms.Label()
        Me.lblSumRate = New System.Windows.Forms.Label()
        Me.lblSumUse = New System.Windows.Forms.Label()
        Me.lblFormulaLossAmount = New System.Windows.Forms.Label()
        Me.lblFormulaLossUnit = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFormulaLossMaking
        '
        Me.lblFormulaLossMaking.BackColor = System.Drawing.Color.DarkSalmon
        Me.lblFormulaLossMaking.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFormulaLossMaking.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblFormulaLossMaking.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormulaLossMaking.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFormulaLossMaking.Location = New System.Drawing.Point(40, 40)
        Me.lblFormulaLossMaking.Name = "lblFormulaLossMaking"
        Me.lblFormulaLossMaking.Size = New System.Drawing.Size(320, 40)
        Me.lblFormulaLossMaking.TabIndex = 5
        Me.lblFormulaLossMaking.Text = "1.Making Formula Loss"
        Me.lblFormulaLossMaking.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.LightCyan
        Me.GroupBox2.Controls.Add(Me.lblFormulaLossMaking)
        Me.GroupBox2.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.Red
        Me.GroupBox2.Location = New System.Drawing.Point(72, 112)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(408, 350)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Data Making "
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.LightCyan
        Me.GroupBox1.Controls.Add(Me.lblSumAmount)
        Me.GroupBox1.Controls.Add(Me.lblSumRate)
        Me.GroupBox1.Controls.Add(Me.lblSumUse)
        Me.GroupBox1.Controls.Add(Me.lblFormulaLossAmount)
        Me.GroupBox1.Controls.Add(Me.lblFormulaLossUnit)
        Me.GroupBox1.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox1.Location = New System.Drawing.Point(544, 112)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(408, 350)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "List Printing"
        '
        'lblSumAmount
        '
        Me.lblSumAmount.BackColor = System.Drawing.Color.LightBlue
        Me.lblSumAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSumAmount.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSumAmount.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSumAmount.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSumAmount.Location = New System.Drawing.Point(40, 271)
        Me.lblSumAmount.Name = "lblSumAmount"
        Me.lblSumAmount.Size = New System.Drawing.Size(320, 40)
        Me.lblSumAmount.TabIndex = 13
        Me.lblSumAmount.Text = "5.Summary Amount"
        Me.lblSumAmount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSumRate
        '
        Me.lblSumRate.BackColor = System.Drawing.Color.LightBlue
        Me.lblSumRate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSumRate.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSumRate.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSumRate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSumRate.Location = New System.Drawing.Point(40, 215)
        Me.lblSumRate.Name = "lblSumRate"
        Me.lblSumRate.Size = New System.Drawing.Size(320, 40)
        Me.lblSumRate.TabIndex = 12
        Me.lblSumRate.Text = "4.Summary Rate"
        Me.lblSumRate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSumUse
        '
        Me.lblSumUse.BackColor = System.Drawing.Color.LightBlue
        Me.lblSumUse.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSumUse.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSumUse.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSumUse.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSumUse.Location = New System.Drawing.Point(40, 161)
        Me.lblSumUse.Name = "lblSumUse"
        Me.lblSumUse.Size = New System.Drawing.Size(320, 40)
        Me.lblSumUse.TabIndex = 11
        Me.lblSumUse.Text = "3.Summary Use"
        Me.lblSumUse.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFormulaLossAmount
        '
        Me.lblFormulaLossAmount.BackColor = System.Drawing.Color.LightBlue
        Me.lblFormulaLossAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFormulaLossAmount.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblFormulaLossAmount.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormulaLossAmount.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFormulaLossAmount.Location = New System.Drawing.Point(40, 48)
        Me.lblFormulaLossAmount.Name = "lblFormulaLossAmount"
        Me.lblFormulaLossAmount.Size = New System.Drawing.Size(320, 40)
        Me.lblFormulaLossAmount.TabIndex = 5
        Me.lblFormulaLossAmount.Text = "1.Formula Loss Amount"
        Me.lblFormulaLossAmount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFormulaLossUnit
        '
        Me.lblFormulaLossUnit.BackColor = System.Drawing.Color.LightBlue
        Me.lblFormulaLossUnit.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFormulaLossUnit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblFormulaLossUnit.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormulaLossUnit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFormulaLossUnit.Location = New System.Drawing.Point(40, 104)
        Me.lblFormulaLossUnit.Name = "lblFormulaLossUnit"
        Me.lblFormulaLossUnit.Size = New System.Drawing.Size(320, 40)
        Me.lblFormulaLossUnit.TabIndex = 10
        Me.lblFormulaLossUnit.Text = "2.Formula Loss Unit"
        Me.lblFormulaLossUnit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmDataMakingMenu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(999, 697)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "FrmDataMakingMenu"
        Me.Controls.SetChildIndex(Me.lblFkey12, 0)
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FrmDataMakingMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Private Sub FrmMainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            MyBase.Title = MyBase.GetItemName("1003")
            MyBase.TitleFontSize = 20
            Me.CloseCaption = "F12:" & MyBase.GetItemName("0009")
            MyBase.IsErrMsg = True
            MyBase.Message = ""
            Me.ShowInTaskbar = False
            Call SetLabel()
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
            Me.lblFormulaLossMaking.Text = objCMT.GetName("BOTM", "0004")
            Me.lblFormulaLossAmount.Text = objCMT.GetName("BOTM", "0005")
            Me.lblFormulaLossUnit.Text = objCMT.GetName("BOTM", "0006")
            Me.lblSumUse.Text = objCMT.GetName("BOTM", "0007")
            Me.lblSumRate.Text = objCMT.GetName("BOTM", "0010")
            Me.lblSumAmount.Text = objCMT.GetName("BOTM", "0011")

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


    Private Sub lblFormulaLossMaking_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFormulaLossMaking.Click
        Try
            ActForm = New FormulaLossMaking(MyBase.UserInfo)
            ActForm.Show()
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub lblFormulaLossAmount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFormulaLossAmount.Click

        Try
            ActForm = New FormulaLossListGrpAmountPrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try

    End Sub

    Private Sub lblFormulaLossunit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFormulaLossUnit.Click

        Try
            ActForm = New FormulaLossListGrpUnitPrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try

    End Sub

    Private Sub lblSumUse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSumUse.Click

        Try
            ActForm = New SummaryOfUsePrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try

    End Sub

    Private Sub lblSumRate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSumRate.Click

        Try
            ActForm = New SummaryOfRatePrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try

    End Sub
    Private Sub lblSumAmount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSumAmount.Click

        Try
            ActForm = New SummaryOfAmountPrinting(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try

    End Sub


End Class