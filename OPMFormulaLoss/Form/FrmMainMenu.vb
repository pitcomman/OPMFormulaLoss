Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass

Public Class FrmMainMenu
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
    Friend WithEvents lblMonthlyDataM As System.Windows.Forms.Label
    Friend WithEvents lblMonthlyClosing As System.Windows.Forms.Label
    Friend WithEvents lblMonthlyReportRP As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.KeyControl1 = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.lblMonthlyClosing = New System.Windows.Forms.Label()
        Me.lblMonthlyReportRP = New System.Windows.Forms.Label()
        Me.lblMonthlyDataM = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblEditMaterialSTDUnit = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblMonthlyClosing
        '
        Me.lblMonthlyClosing.BackColor = System.Drawing.Color.LightBlue
        Me.lblMonthlyClosing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMonthlyClosing.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblMonthlyClosing.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonthlyClosing.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblMonthlyClosing.Location = New System.Drawing.Point(19, 98)
        Me.lblMonthlyClosing.Name = "lblMonthlyClosing"
        Me.lblMonthlyClosing.Size = New System.Drawing.Size(420, 60)
        Me.lblMonthlyClosing.TabIndex = 6
        Me.lblMonthlyClosing.Text = "2. Monthly Closing Menu"
        Me.lblMonthlyClosing.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblMonthlyReportRP
        '
        Me.lblMonthlyReportRP.BackColor = System.Drawing.Color.LightBlue
        Me.lblMonthlyReportRP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMonthlyReportRP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblMonthlyReportRP.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonthlyReportRP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblMonthlyReportRP.Location = New System.Drawing.Point(19, 172)
        Me.lblMonthlyReportRP.Name = "lblMonthlyReportRP"
        Me.lblMonthlyReportRP.Size = New System.Drawing.Size(420, 60)
        Me.lblMonthlyReportRP.TabIndex = 4
        Me.lblMonthlyReportRP.Text = "3. List Re-Printing Menu"
        Me.lblMonthlyReportRP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblMonthlyDataM
        '
        Me.lblMonthlyDataM.BackColor = System.Drawing.Color.LightBlue
        Me.lblMonthlyDataM.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMonthlyDataM.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblMonthlyDataM.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonthlyDataM.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblMonthlyDataM.Location = New System.Drawing.Point(19, 29)
        Me.lblMonthlyDataM.Name = "lblMonthlyDataM"
        Me.lblMonthlyDataM.Size = New System.Drawing.Size(420, 60)
        Me.lblMonthlyDataM.TabIndex = 5
        Me.lblMonthlyDataM.Text = "1. Data Making and Printing Menu"
        Me.lblMonthlyDataM.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.AccessibleName = "D"
        Me.GroupBox2.BackColor = System.Drawing.Color.LightCyan
        Me.GroupBox2.Controls.Add(Me.lblMonthlyClosing)
        Me.GroupBox2.Controls.Add(Me.lblMonthlyReportRP)
        Me.GroupBox2.Controls.Add(Me.lblMonthlyDataM)
        Me.GroupBox2.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(37, 109)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(450, 250)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Data Making"
        '
        'GroupBox1
        '
        Me.GroupBox1.AccessibleName = "D"
        Me.GroupBox1.BackColor = System.Drawing.Color.LightCyan
        Me.GroupBox1.Controls.Add(Me.lblEditMaterialSTDUnit)
        Me.GroupBox1.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(513, 109)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(450, 250)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Maintenance"
        '
        'lblEditMaterialSTDUnit
        '
        Me.lblEditMaterialSTDUnit.BackColor = System.Drawing.Color.LightBlue
        Me.lblEditMaterialSTDUnit.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEditMaterialSTDUnit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblEditMaterialSTDUnit.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEditMaterialSTDUnit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEditMaterialSTDUnit.Location = New System.Drawing.Point(15, 29)
        Me.lblEditMaterialSTDUnit.Name = "lblEditMaterialSTDUnit"
        Me.lblEditMaterialSTDUnit.Size = New System.Drawing.Size(420, 60)
        Me.lblEditMaterialSTDUnit.TabIndex = 5
        Me.lblEditMaterialSTDUnit.Text = "1. Edit Material Standard Unit"
        Me.lblEditMaterialSTDUnit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmMainMenu
        '
        Me.AccessibleName = ""
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(1013, 711)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "FrmMainMenu"
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.lblFkey12, 0)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FrmMainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            MyBase.Title = MyBase.GetItemName("1002")
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
            Me.lblMonthlyDataM.Text = objCMT.GetName("BOTM", "0001")
            Me.lblMonthlyClosing.Text = objCMT.GetName("BOTM", "0002")
            Me.lblMonthlyReportRP.Text = objCMT.GetName("BOTM", "0003")

        Catch ex As Exception
            Throw ex
        Finally

            If Not objCMT Is Nothing Then
                objCMT = Nothing

            End If
        End Try

    End Function

    Public Overrides Function PushF12() As Object
        End
    End Function


    Private Sub lblMonthlyDataM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMonthlyDataM.Click
        Try
            ActForm = New FrmDataMakingMenu(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub lblMonthlyClosing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMonthlyClosing.Click
        Try
            ActForm = New FrmMonthlyClosingMenu(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub lblMonthlyReportRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMonthlyReportRP.Click
        Try
            ActForm = New FrmDataPrintingMenu(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblEditMaterialSTDUnit As System.Windows.Forms.Label

    Private Sub lblEditMaterialSTDUnit_Click(sender As Object, e As EventArgs) Handles lblEditMaterialSTDUnit.Click
        Try
            ActForm = New FrmEditMaterialSTDUnit(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub
End Class