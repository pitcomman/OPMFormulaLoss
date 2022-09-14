Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass
Imports System.Data.SqlClient
Imports System.IO

Public Class FrmMonthlyClosingMenu
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
    Friend WithEvents lblMonthlyClosing As System.Windows.Forms.Label
    Friend WithEvents lblOpeMonth As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.KeyControl1 = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblMonthlyClosing = New System.Windows.Forms.Label()
        Me.lblOpeMonth = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.LightCyan
        Me.GroupBox2.Controls.Add(Me.lblMonthlyClosing)
        Me.GroupBox2.Controls.Add(Me.lblOpeMonth)
        Me.GroupBox2.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(257, 149)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(472, 312)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Monthly Closing"
        '
        'lblMonthlyClosing
        '
        Me.lblMonthlyClosing.BackColor = System.Drawing.Color.LightBlue
        Me.lblMonthlyClosing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMonthlyClosing.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblMonthlyClosing.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonthlyClosing.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblMonthlyClosing.Location = New System.Drawing.Point(20, 132)
        Me.lblMonthlyClosing.Name = "lblMonthlyClosing"
        Me.lblMonthlyClosing.Size = New System.Drawing.Size(392, 56)
        Me.lblMonthlyClosing.TabIndex = 8
        Me.lblMonthlyClosing.Text = "2.Monthly Closing"
        Me.lblMonthlyClosing.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOpeMonth
        '
        Me.lblOpeMonth.BackColor = System.Drawing.Color.LightBlue
        Me.lblOpeMonth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpeMonth.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblOpeMonth.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpeMonth.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOpeMonth.Location = New System.Drawing.Point(20, 46)
        Me.lblOpeMonth.Name = "lblOpeMonth"
        Me.lblOpeMonth.Size = New System.Drawing.Size(392, 64)
        Me.lblOpeMonth.TabIndex = 9
        Me.lblOpeMonth.Text = "1.Operation Month Maintenance"
        Me.lblOpeMonth.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmMonthlyClosingMenu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(998, 696)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "FrmMonthlyClosingMenu"
        Me.Controls.SetChildIndex(Me.lblFkey12, 0)
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FrmMainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            MyBase.Title = MyBase.GetItemName("1004")
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


    Private Sub FrmMainMenu_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            If Not ActForm Is Nothing Then
                ActForm.Activate()
            End If
            If MyBase.AuthorityLV = "00" Then
                lblOpeMonth.Enabled = True
                lblMonthlyClosing.Enabled = True

            Else
                lblOpeMonth.Enabled = False
                lblMonthlyClosing.Enabled = False
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

            Me.lblOpeMonth.Text = objCMT.GetName("BOTM", "0008")
            Me.lblMonthlyClosing.Text = objCMT.GetName("BOTM", "0009")

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

    Private Sub lblOpeMonth_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblOpeMonth.Click
        Try
            ActForm = New FrmOpMonthMaintenance(MyBase.UserInfo)
            ActForm.Show()

        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub lblMonthlyClosing_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblMonthlyClosing.Click
        Try
            ActForm = New MonthlyClosingMaking(MyBase.UserInfo)
            ActForm.Show()
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub
End Class