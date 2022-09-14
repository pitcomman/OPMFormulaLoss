Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass
Imports System.Data.SqlClient
Imports System.IO
Public Class FrmOpMonthMaintenance
    Inherits Rist.OPMCmnClass.PageBase


    Dim ActForm As Form
    'Private objMatTbl As DataTable

#Region " Windows Form Designer generated code "

    Public Sub New()
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
            ' dispose event(connection close when Application ending)
            'If MyBase.objDBBase.Conn.State = ConnectionState.Open Then
            '    MyBase.objDBBase.CloseConnection()
            'End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblKeyF1 As Rist.OPMCmnClass.FunctionKeyControl
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents KeyControl As Rist.OPMCmnClass.KeyControl
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents txtFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtMonth As System.Windows.Forms.TextBox
    Friend WithEvents TxtYear As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GrdOpMonthDetail As System.Windows.Forms.DataGrid
    Friend WithEvents txtOperatorName As System.Windows.Forms.Label
    Friend WithEvents txtToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblKeyF3 As Rist.OPMCmnClass.FunctionKeyControl
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents lblMonth As System.Windows.Forms.Label
    Friend WithEvents lblToDate As System.Windows.Forms.Label
    Friend WithEvents lblFromDate As System.Windows.Forms.Label
    Friend WithEvents txtMCFlag As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.lblYear = New System.Windows.Forms.Label
        Me.lblMonth = New System.Windows.Forms.Label
        Me.TxtMonth = New System.Windows.Forms.TextBox
        Me.lblKeyF1 = New Rist.OPMCmnClass.FunctionKeyControl
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtYear = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtOperatorName = New System.Windows.Forms.Label
        Me.lblToDate = New System.Windows.Forms.Label
        Me.txtToDate = New System.Windows.Forms.DateTimePicker
        Me.txtFromDate = New System.Windows.Forms.DateTimePicker
        Me.lblFromDate = New System.Windows.Forms.Label
        Me.GrdOpMonthDetail = New System.Windows.Forms.DataGrid
        Me.KeyControl = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.LblKeyF3 = New Rist.OPMCmnClass.FunctionKeyControl
        Me.txtMCFlag = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.GrdOpMonthDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblVersion
        '
        Me.lblVersion.Name = "lblVersion"
        '
        'lblFkey12
        '
        Me.lblFkey12.Name = "lblFkey12"
        '
        'lblYear
        '
        Me.lblYear.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear.ForeColor = System.Drawing.Color.Black
        Me.lblYear.Location = New System.Drawing.Point(96, 24)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(64, 24)
        Me.lblYear.TabIndex = 4
        Me.lblYear.Text = "Year :"
        '
        'lblMonth
        '
        Me.lblMonth.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonth.ForeColor = System.Drawing.Color.Black
        Me.lblMonth.Location = New System.Drawing.Point(328, 24)
        Me.lblMonth.Name = "lblMonth"
        Me.lblMonth.Size = New System.Drawing.Size(72, 24)
        Me.lblMonth.TabIndex = 5
        Me.lblMonth.Text = "Month :"
        '
        'TxtMonth
        '
        Me.TxtMonth.AutoSize = False
        Me.TxtMonth.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtMonth.Location = New System.Drawing.Point(400, 16)
        Me.TxtMonth.Name = "TxtMonth"
        Me.TxtMonth.Size = New System.Drawing.Size(32, 32)
        Me.TxtMonth.TabIndex = 7
        Me.TxtMonth.Text = ""
        '
        'lblKeyF1
        '
        Me.lblKeyF1.Location = New System.Drawing.Point(24, 656)
        Me.lblKeyF1.Name = "lblKeyF1"
        Me.lblKeyF1.Size = New System.Drawing.Size(152, 48)
        Me.lblKeyF1.TabIndex = 10
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TxtYear)
        Me.GroupBox1.Controls.Add(Me.lblYear)
        Me.GroupBox1.Controls.Add(Me.lblMonth)
        Me.GroupBox1.Controls.Add(Me.TxtMonth)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(544, 56)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial Black", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Blue
        Me.Label5.Location = New System.Drawing.Point(224, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 32)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "( 20xx )"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial Black", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(440, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 32)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "( 01-12 )"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TxtYear
        '
        Me.TxtYear.AutoSize = False
        Me.TxtYear.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtYear.Location = New System.Drawing.Point(152, 16)
        Me.TxtYear.Name = "TxtYear"
        Me.TxtYear.Size = New System.Drawing.Size(64, 32)
        Me.TxtYear.TabIndex = 6
        Me.TxtYear.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtOperatorName)
        Me.GroupBox2.Controls.Add(Me.lblToDate)
        Me.GroupBox2.Controls.Add(Me.txtToDate)
        Me.GroupBox2.Controls.Add(Me.txtFromDate)
        Me.GroupBox2.Controls.Add(Me.lblFromDate)
        Me.GroupBox2.Controls.Add(Me.GrdOpMonthDetail)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 192)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1000, 432)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        '
        'txtOperatorName
        '
        Me.txtOperatorName.Location = New System.Drawing.Point(576, 392)
        Me.txtOperatorName.Name = "txtOperatorName"
        Me.txtOperatorName.Size = New System.Drawing.Size(104, 8)
        Me.txtOperatorName.TabIndex = 94
        Me.txtOperatorName.Visible = False
        '
        'lblToDate
        '
        Me.lblToDate.BackColor = System.Drawing.Color.LightCyan
        Me.lblToDate.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToDate.Location = New System.Drawing.Point(528, 32)
        Me.lblToDate.Name = "lblToDate"
        Me.lblToDate.Size = New System.Drawing.Size(72, 24)
        Me.lblToDate.TabIndex = 93
        Me.lblToDate.Text = "To Date:"
        '
        'txtToDate
        '
        Me.txtToDate.CustomFormat = ""
        Me.txtToDate.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtToDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.txtToDate.Location = New System.Drawing.Point(600, 24)
        Me.txtToDate.Name = "txtToDate"
        Me.txtToDate.Size = New System.Drawing.Size(128, 29)
        Me.txtToDate.TabIndex = 92
        Me.txtToDate.Value = New Date(2006, 10, 24, 17, 11, 7, 631)
        '
        'txtFromDate
        '
        Me.txtFromDate.CustomFormat = ""
        Me.txtFromDate.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.txtFromDate.Location = New System.Drawing.Point(304, 24)
        Me.txtFromDate.Name = "txtFromDate"
        Me.txtFromDate.Size = New System.Drawing.Size(128, 29)
        Me.txtFromDate.TabIndex = 73
        Me.txtFromDate.Value = New Date(2006, 10, 24, 17, 11, 7, 631)
        '
        'lblFromDate
        '
        Me.lblFromDate.BackColor = System.Drawing.Color.LightCyan
        Me.lblFromDate.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromDate.Location = New System.Drawing.Point(208, 32)
        Me.lblFromDate.Name = "lblFromDate"
        Me.lblFromDate.Size = New System.Drawing.Size(96, 24)
        Me.lblFromDate.TabIndex = 4
        Me.lblFromDate.Text = "From Date:"
        '
        'GrdOpMonthDetail
        '
        Me.GrdOpMonthDetail.DataMember = ""
        Me.GrdOpMonthDetail.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.GrdOpMonthDetail.Location = New System.Drawing.Point(16, 64)
        Me.GrdOpMonthDetail.Name = "GrdOpMonthDetail"
        Me.GrdOpMonthDetail.ReadOnly = True
        Me.GrdOpMonthDetail.Size = New System.Drawing.Size(960, 296)
        Me.GrdOpMonthDetail.TabIndex = 89
        '
        'KeyControl
        '
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.GroupBox1)
        Me.GroupBox4.Controls.Add(Me.txtMCFlag)
        Me.GroupBox4.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox4.Location = New System.Drawing.Point(8, 96)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(1000, 88)
        Me.GroupBox4.TabIndex = 13
        Me.GroupBox4.TabStop = False
        '
        'LblKeyF3
        '
        Me.LblKeyF3.Location = New System.Drawing.Point(216, 656)
        Me.LblKeyF3.Name = "LblKeyF3"
        Me.LblKeyF3.Size = New System.Drawing.Size(152, 48)
        Me.LblKeyF3.TabIndex = 14
        '
        'txtMCFlag
        '
        Me.txtMCFlag.AutoSize = False
        Me.txtMCFlag.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMCFlag.Location = New System.Drawing.Point(704, 24)
        Me.txtMCFlag.Name = "txtMCFlag"
        Me.txtMCFlag.Size = New System.Drawing.Size(64, 32)
        Me.txtMCFlag.TabIndex = 10
        Me.txtMCFlag.Text = ""
        Me.txtMCFlag.Visible = False
        '
        'OpMonthMaintenance
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(1012, 710)
        Me.Controls.Add(Me.LblKeyF3)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.lblKeyF1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "OpMonthMaintenance"
        Me.Controls.SetChildIndex(Me.lblFkey12, 0)
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.Controls.SetChildIndex(Me.lblKeyF1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox4, 0)
        Me.Controls.SetChildIndex(Me.LblKeyF3, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.GrdOpMonthDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region " Envents Function "

    Private Sub OpMonthMaintenance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim dd As Date = DateAdd(DateInterval.Day, 1, Now())

        Try
            MyBase.Title = MyBase.GetItemName("1005")
            MyBase.TitleFontSize = 20
            Me.lblKeyF1.Caption = "F1:" & MyBase.GetItemName("0005")
            Me.CloseCaption = "F12:" & MyBase.GetItemName("0009")
            Me.LblKeyF3.Caption = "F3:" & MyBase.GetItemName("0007")
            MyBase.Message = ""
            Me.ShowInTaskbar = False
            Call InitialComponent()
            Me.txtFromDate.Value = Now()
            Me.txtToDate.Value = dd
            Me.TxtYear.Focus()
            Call SetLabel()


            OpMonthTbl = CreateOpMonthDetailTbl()
            Call initialOpMonthTblStyle(GrdOpMonthDetail, OpMonthTbl)


        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg(ex.MsgCode)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub
    Private Sub OpMonthMaintenance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Try
            Me.KeyControl.Push(e.KeyValue)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub OpMonthMaintenance_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            If Not ActForm Is Nothing Then
                ActForm.Activate()
            End If
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub
#End Region

#Region " Overrides Function "

    ' on push F12 event
    Public Overrides Function PushF12() As Object
        Close()
    End Function

    ' on click F1 label event
    Private Sub lblKeyF1_UCClick() Handles lblKeyF1.UCClick
        Try
            Me.KeyControl.Push(Keys.F1)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    'on click F3 label event
    Private Sub LblKeyF3_UCClick() Handles LblKeyF3.UCClick
        Try
            Me.KeyControl.Push(Keys.F3)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub


    'Delete Funcrion
    Private Sub KeyControl_PushF3() Handles KeyControl.PushF3
        Dim strOpMonth As String

        strOpMonth = (Trim(Me.TxtYear.Text) & Trim(Me.TxtMonth.Text))

        Try
            'Me.txtYear.Text = UCase(Me.txtSheetNo.Text)
            If Trim(Me.TxtYear.Text) = "" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E042")
                Exit Sub

            Else
                If Trim(Me.TxtYear.Text) > "2020" Or Trim(Me.TxtYear.Text) < "2006" Then
                    MyBase.IsErrMsg = True
                    MyBase.ShowMsg("E043")
                    Me.TxtYear.Focus()
                    Exit Sub

                End If
            End If

            If Trim(Me.TxtMonth.Text) = "" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E044")
                Exit Sub

            Else
                If funOpMonthEdit(strOpMonth) = True Then

                    funOpMonthCancel(strOpMonth)

                Else

                    MyBase.Message = ""

                End If
            End If
            OpMonthTbl.Rows.Clear()
            Call OpMonthTableSubfrom2()

        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg("E022")
            MyBase.ShowMsg(ex.MsgCode)
            Exit Sub
        Catch ex As Exception
            MyBase.Message = ex.Message
            MyBase.IsErrMsg = True
            MyBase.ShowMsg("E022")
            Exit Sub

        End Try

        MyBase.IsErrMsg = False
        MyBase.ShowMsg("I003")
        Call InitialComponent()
    End Sub

    ' on push F1 event
    Private Sub KeyControl_PushF1() Handles KeyControl.PushF1

        Dim strOpmonth As String

        Try
            If Trim(Me.TxtYear.Text) = "" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E042")
                Me.TxtYear.Focus()
                Exit Sub
            End If


            If Trim(Me.TxtYear.Text) > "2020" Or Trim(Me.TxtYear.Text) < "2006" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E043")
                Me.TxtYear.Focus()
                Exit Sub

            End If

            If Trim(Me.TxtMonth.Text) = "" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E044")
                Me.TxtMonth.Focus()
                Exit Sub

            End If
            If Trim(Me.TxtMonth.Text) > "12" Or Trim(Me.TxtMonth.Text) < "01" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E045")
                Me.TxtMonth.Focus()
                Exit Sub

            End If
            If Trim(Me.txtMCFlag.Text) = "1" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E058")
                Me.TxtMonth.Focus()
                Exit Sub
            End If

            Me.txtOperatorName.Text = MyBase.[Operator]()
            'OpMonth = YearMonth
            strOpmonth = Trim(Me.TxtYear.Text) & Trim(Me.TxtMonth.Text)

            funOpMonthEditCheck(strOpmonth)               '<----Delete old entry

            funOpMonthInsert(strOpmonth, Me.txtFromDate.Value, Me.txtToDate.Value, Me.txtOperatorName.Text)

        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg("E022")
            MyBase.ShowMsg(ex.MsgCode)
            Exit Sub
        Catch ex As Exception
            MyBase.Message = ex.Message
            MyBase.IsErrMsg = True
            MyBase.ShowMsg("E022")
            Exit Sub
        Finally

        End Try

        'Call SubForm
        OpMonthTbl.Rows.Clear()
        Call OpMonthTableSubfrom2()
        Call InitialComponent()
        MyBase.IsErrMsg = False
        MyBase.ShowMsg("I005")

        Me.TxtYear.Focus()

    End Sub
    'Set LabelName from Purpose Table
    Private Function SetLabel()
        Dim objCMT As CmnMasterTable
        Try
            objCMT = New CmnMasterTable(MyBase.SystemName)
            objCMT.Conn = MyBase.objDBBase.Conn
            Me.lblYear.Text = objCMT.GetName("ITEM", "1007")
            Me.lblMonth.Text = objCMT.GetName("ITEM", "1008")
            Me.lblFromDate.Text = objCMT.GetName("ITEM", "1009")
            Me.lblToDate.Text = objCMT.GetName("ITEM", "1010")

        Catch ex As Exception
            Throw ex
        Finally
            If Not objCMT Is Nothing Then
                objCMT = Nothing
            End If
        End Try
    End Function

#End Region

#Region " Processing Function "

    'Clear Entry Field
    Private Sub InitialComponent()

        Me.TxtMonth.Text = ""
        Me.txtMCFlag.Text = ""
        MyBase.Message = ""
        Me.TxtYear.Focus()
    End Sub


    'Set DataGrid
    Private Sub OpMonthTableSubfrom()
        Dim DataGriddrRow As DataRow
        Dim dtTbl As DataTable

        Dim strOpMonth As String
        strOpMonth = Trim(Me.TxtYear.Text) & Trim(Me.TxtMonth.Text)


        Try
            dtTbl = funOpMonthChecking(strOpMonth)
            If dtTbl.Rows.Count <> 0 Then
                For Each drRow As DataRow In dtTbl.Rows
                    DataGriddrRow = OpMonthTbl.NewRow
                    DataGriddrRow("OpMonth") = drRow("OpMonth")
                    DataGriddrRow("FromDate") = drRow("FromDate")
                    DataGriddrRow("ToDate") = drRow("ToDate")
                    DataGriddrRow("MonthlyClosingFlag") = drRow("MonthlyClosingFlag")
                    DataGriddrRow("OperatorName") = drRow("OperatorName")


                    OpMonthTbl.Rows.Add(DataGriddrRow)
                    OpMonthTbl.AcceptChanges()
                Next
            Else
                Exit Sub
            End If
            Me.GrdOpMonthDetail.Refresh()


        Catch ex As CustomErrException
            Throw ex

        Catch ex As Exception
            Throw ex

        End Try

    End Sub



    'Error Checking when entry data

    Private Sub TxtYear_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtYear.KeyUp

        TxtYear.MaxLength = 4   '<---Set the maxlength input characters in year textbox


        If e.KeyValue = Keys.Enter Then


            If Trim(Me.TxtYear.Text) = "" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E042")
                Me.TxtYear.Focus()
                Exit Sub

            Else
                If Trim(Me.TxtYear.Text) > "2020" Or Trim(Me.TxtYear.Text) < "2006" Then
                    MyBase.IsErrMsg = True
                    MyBase.ShowMsg("E043")
                    Me.TxtYear.Focus()
                    Exit Sub

                Else


                    OpMonthTbl.Rows.Clear()
                    Call OpMonthTableSubfrom2()
                    MyBase.Message = ""
                    Me.TxtMonth.Focus()


                End If
            End If
        End If


    End Sub



    'check input character
    Private Sub TxtMonth_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMonth.KeyUp



        TxtMonth.MaxLength = 2      '<---Set the maxlength input characters in month textbox
        Dim dtTbl As DataTable
        Dim strOpMonth As String
        strOpMonth = Trim(Me.TxtYear.Text) & Trim(Me.TxtMonth.Text)

        If e.KeyValue = Keys.Enter Then

            If Trim(Me.TxtMonth.Text) = "" Then
                MyBase.IsErrMsg = True
                MyBase.ShowMsg("E044")
                Me.TxtMonth.Focus()
                Exit Sub
            Else
                If Trim(Me.TxtMonth.Text) > "12" Or Trim(Me.TxtMonth.Text) < "01" Then
                    MyBase.IsErrMsg = True
                    MyBase.ShowMsg("E045")
                    Me.TxtMonth.Focus()
                    Exit Sub
                Else

                    Try

                        dtTbl = funOpMonthChecking(strOpMonth)
                        If dtTbl.Rows.Count > 0 Then
                            Me.txtFromDate.Text = IIf(IsDBNull(dtTbl.Rows(0)("FromDate")), "", dtTbl.Rows(0)("FromDate"))
                            Me.txtToDate.Text = IIf(IsDBNull(dtTbl.Rows(0)("ToDate")), "", dtTbl.Rows(0)("ToDate"))
                            Me.txtMCFlag.Text = IIf(IsDBNull(dtTbl.Rows(0)("MonthlyClosingFlag")), "", dtTbl.Rows(0)("MonthlyClosingFlag"))

                            'for f1 Edit  =  delete + insert
                        End If
                    Catch ex As CustomErrException
                        Throw ex

                    Catch ex As Exception
                        Throw ex

                    End Try


                End If

                MyBase.Message = ""
                Me.txtFromDate.Focus()

            End If
        End If

    End Sub
    'on select value,check data
    Private Sub txtFromDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFromDate.ValueChanged

        '' for RASP & RAJP
        ''If (Format(Me.txtFromDate.Value, "MM/dd/YYYY")) > (Format(Me.txtToDate.Value, "MM/dd/YYYY")) Then
        '' for RAET
        ''If (Format(Me.txtFromDate.Value, "dd/MM/YYYY")) > (Format(Me.txtToDate.Value, "dd/MM/YYYY")) Then
        ''  If (Trim(Me.txtFromDate.Value) > Trim(Me.txtToDate.Value)) Then
        If (Me.txtFromDate.Value > Me.txtToDate.Value) Then

            MyBase.IsErrMsg = True
            MyBase.ShowMsg("E046")
            Me.txtToDate.Focus()
            Exit Sub
        Else
            MyBase.Message = ""
            Me.txtToDate.Focus()
        End If

    End Sub
    'Check if ToDate is greater than FromDate
    Private Sub txtToDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtToDate.ValueChanged

        ''If (Format(Me.txtFromDate.Value, "dd/MM/YYYY")) > (Format(Me.txtToDate.Value, "dd/MM/YYYY")) Then
        ''If (Trim(Me.txtFromDate.Value) > Trim(Me.txtToDate.Value)) Then
        If (Me.txtFromDate.Value > Me.txtToDate.Value) Then
            MyBase.IsErrMsg = True
            MyBase.ShowMsg("E046")
            Me.txtToDate.Focus()
            Exit Sub
        Else
            MyBase.Message = ""
        End If

    End Sub

    'check if data exist
    Function funOpMonthChecking(ByVal strOpMonth As String) As DataTable

        Dim objOpMonth As DBConnect
        Dim dtData As DataTable = Nothing

        Try
            '    ' make new DBConnect instance
            objOpMonth = New DBConnect
            objOpMonth.Conn = MyBase.objDBBase.Conn
            dtData = objOpMonth.GetOpMonthTbl(strOpMonth)
            Return dtData

        Catch ex As CustomErrException
            Throw ex

        Catch ex As Exception
            Throw ex

        End Try

    End Function


    Function funOpMonthEditCheck(ByVal strOpMonth As String) As Boolean
        Dim objOpMonthEditCheck As DBConnect
        Dim dtData As DataTable = Nothing

        strOpMonth = Trim(Me.TxtYear.Text) & Trim(Me.TxtMonth.Text)

        Try
            '    ' make new DBConnect instance
            objOpMonthEditCheck = New DBConnect
            objOpMonthEditCheck.Conn = MyBase.objDBBase.Conn
            dtData = objOpMonthEditCheck.OpMonthEdit(strOpMonth)
            If Not dtData Is Nothing And dtData.Rows.Count > 0 Then
                'call delete
                Call funOpMonthEdit(strOpMonth)
                Return True

            End If
        Catch ex As CustomErrException
            Throw ex

        Catch ex As Exception
            Throw ex

        End Try

    End Function

    'Check data exist and retrieve data

    'Insert Data Entry from Table OperationMonth

    Function funOpMonthInsert(ByVal strOpmonth As String, ByVal strFromDate As DateTime, _
                        ByVal strToDate As DateTime, ByVal strOperatorName As String) As Boolean

        Dim objOpMonth As DBConnect
        Dim dtData As DataTable = Nothing

        'Dim Qty As Double
        Try
            '    ' make new DBConnect instance
            objOpMonth = New DBConnect
            objOpMonth.Conn = MyBase.objDBBase.Conn
            dtData = objOpMonth.OpMonthInsert(strOpmonth, strFromDate, strToDate, strOperatorName)

        Catch ex As CustomErrException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    'delete data if exist
    Function funOpMonthEdit(ByVal strOpMonth As String) As Boolean
        Dim objOpMonthEdit As DBConnect
        Dim dtData As DataTable = Nothing

        strOpMonth = Trim(Me.TxtYear.Text) & Trim(Me.TxtMonth.Text)

        Try
            '    ' make new DBConnect instance
            objOpMonthEdit = New DBConnect
            objOpMonthEdit.Conn = MyBase.objDBBase.Conn
            dtData = objOpMonthEdit.OpMonthEdit(strOpMonth)


        Catch ex As CustomErrException
            Throw ex

        Catch ex As Exception
            Throw ex

        End Try

    End Function


    'Function Filter by year 

    Function funAllOpMonthDisplay(ByVal strOpMonth As String) As DataTable

        Dim objAllOpMonth As DBConnect
        Dim dtData As DataTable = Nothing

        Try
            '    ' make new DBConnect instance
            objAllOpMonth = New DBConnect
            objAllOpMonth.Conn = MyBase.objDBBase.Conn
            dtData = objAllOpMonth.AllOpmonthDisplay(strOpMonth)
            Return dtData

        Catch ex As CustomErrException
            Throw ex

        Catch ex As Exception
            Throw ex

        End Try

    End Function

    'Display all data In SubForm,Filter by year
    Private Sub OpMonthTableSubfrom2()
        Dim DataGriddrRow As DataRow
        Dim dtTbl As DataTable

        Dim strOpMonth As String
        strOpMonth = Trim(Me.TxtYear.Text)


        Try
            dtTbl = funAllOpMonthDisplay(strOpMonth)
            If dtTbl.Rows.Count <> 0 Then
                For Each drRow As DataRow In dtTbl.Rows
                    DataGriddrRow = OpMonthTbl.NewRow
                    DataGriddrRow("OpMonth") = drRow("OpMonth")
                    DataGriddrRow("FromDate") = drRow("FromDate")
                    DataGriddrRow("ToDate") = drRow("ToDate")
                    DataGriddrRow("MonthlyClosingFlag") = drRow("MonthlyClosingFlag")
                    DataGriddrRow("OperatorName") = drRow("OperatorName")


                    OpMonthTbl.Rows.Add(DataGriddrRow)
                    OpMonthTbl.AcceptChanges()
                Next
            Else
                Exit Sub
            End If
            Me.GrdOpMonthDetail.Refresh()


        Catch ex As CustomErrException
            Throw ex

        Catch ex As Exception
            Throw ex

        End Try

    End Sub

    'Cancel Data Entry
    Function funOpMonthCancel(ByVal strOpMonth As String) As Boolean

        Dim objOpMonthCancel As DBConnect
        Dim dtData As DataTable = Nothing

        strOpMonth = Trim(Me.TxtYear.Text) & Trim(Me.TxtMonth.Text)


        Try
            '    ' make new DBConnect instance
            objOpMonthCancel = New DBConnect
            objOpMonthCancel.Conn = MyBase.objDBBase.Conn
            dtData = objOpMonthCancel.OpMonthCancel(strOpMonth)

        Catch ex As CustomErrException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try

    End Function

#End Region






End Class