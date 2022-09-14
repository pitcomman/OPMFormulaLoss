'***********************************************************************
' Program Name	    : Dialog Control Class
' Program ID	    : DialogControl
' Function			: this Class have Dialog Control Function
' Create Date		: 2006/06/19
' Create Person		: Manop
' 
' Supplement	    :
' Version		    : 1.00
' ---------------------------------------------------------------------
' Condition　　　　	: SqlServer2000,ADO.Net,.NetFramework
' Starting Way		:
'***********************************************************************
Imports System.Data.SqlClient

Public Class DialogControl
    Inherits System.Windows.Forms.Form

    ' connection object
    Protected objDBBase As DBBase
    ' System Name
    Private strSystemName As String
    ' Factory Name
    Private strFactoryName As String
    ' Operator
    Private strOperator As String
    ' Workstation Id
    Private strWorkstation As String
    ' AuthorityLV
    Private strAuthLV As String
    ' Field Name(for Message or Item Name or etc.)
    Private strFieldName As String
    ' Login information
    Private objUserInfo As Hashtable

#Region " Windows Form Designer generated code "

    ' default constractor
    Public Sub New()
        MyBase.New()

        ' make DBBase class instance
        objDBBase = New DBBase

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' custom constracter(for except Login Form)
    Public Sub New(ByVal pobjUserInfo As Hashtable)
        ' call base constractor
        MyBase.New()
        Try
            ' make DBBase class instance
            objDBBase = New DBBase

            InitializeComponent()

            ' set User information
            Me.objUserInfo = pobjUserInfo

            ' set base information
            Me.SetBaseInfo()

            ' open connection
            If objDBBase Is Nothing Or objDBBase.Conn.State = ConnectionState.Closed Then
                objDBBase.OpenConnection()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' custom constracter(for Login Form)
    Public Sub New(ByVal pstrSystemName As String)
        ' call base constractor
        MyBase.New()

        Try
            ' make DBBase class instance
            objDBBase = New DBBase(pstrSystemName)

            ' make DBBase class instance and open connection
            If objDBBase Is Nothing Or objDBBase.Conn.State = ConnectionState.Closed Then
                objDBBase.OpenConnection()
            End If

            'This call is required by the Windows Form Designer.
            InitializeComponent()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
            'If objDBBase.Conn.State = ConnectionState.Open Then
            '    objDBBase.CloseConnection()
            'End If
        End If

        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblMsg As System.Windows.Forms.Label
    Friend WithEvents lblFinishTime As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblStartTime As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblMsg = New System.Windows.Forms.Label
        Me.lblFinishTime = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblStartTime = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Blue
        Me.lblTitle.Location = New System.Drawing.Point(7, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(392, 32)
        Me.lblTitle.TabIndex = 37
        Me.lblTitle.Text = "lblTitle"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblMsg)
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.GroupBox1.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox1.Location = New System.Drawing.Point(7, 49)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(392, 64)
        Me.GroupBox1.TabIndex = 36
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Message"
        '
        'lblMsg
        '
        Me.lblMsg.Font = New System.Drawing.Font("Tahoma", 12.0!)
        Me.lblMsg.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblMsg.Location = New System.Drawing.Point(8, 24)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(376, 24)
        Me.lblMsg.TabIndex = 0
        Me.lblMsg.Text = "lblMsg"
        Me.lblMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFinishTime
        '
        Me.lblFinishTime.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.lblFinishTime.Location = New System.Drawing.Point(287, 129)
        Me.lblFinishTime.Name = "lblFinishTime"
        Me.lblFinishTime.Size = New System.Drawing.Size(72, 10)
        Me.lblFinishTime.TabIndex = 35
        Me.lblFinishTime.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.Label2.Location = New System.Drawing.Point(199, 121)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 24)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Finish Time :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.Label1.Location = New System.Drawing.Point(39, 121)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 24)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Start Time :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblStartTime
        '
        Me.lblStartTime.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.lblStartTime.Location = New System.Drawing.Point(127, 129)
        Me.lblStartTime.Name = "lblStartTime"
        Me.lblStartTime.Size = New System.Drawing.Size(69, 10)
        Me.lblStartTime.TabIndex = 32
        Me.lblStartTime.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCancel
        '
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCancel.Font = New System.Drawing.Font("Tahoma", 14.0!)
        Me.btnCancel.Location = New System.Drawing.Point(319, 153)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 34)
        Me.btnCancel.TabIndex = 31
        Me.btnCancel.Text = "Cancel"
        '
        'btnStart
        '
        Me.btnStart.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnStart.Font = New System.Drawing.Font("Tahoma", 14.0!)
        Me.btnStart.Location = New System.Drawing.Point(231, 153)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(80, 34)
        Me.btnStart.TabIndex = 30
        Me.btnStart.Text = "Start"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.Label3.Location = New System.Drawing.Point(111, 129)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "_______________"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.Label4.Location = New System.Drawing.Point(271, 129)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "_______________"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DialogControl
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(406, 197)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblFinishTime)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblStartTime)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "DialogControl"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DialogControl"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Property Function "
    Public WriteOnly Property Title()
        Set(ByVal Value)
            Me.lblTitle.Text = Value
        End Set
    End Property

    Public WriteOnly Property TitleFontSize() As Integer
        Set(ByVal Value As Integer)
            Me.lblTitle.Font = New System.Drawing.Font("Tahoma", Value)
        End Set
    End Property

    Public WriteOnly Property StartTime()
        Set(ByVal Value)
            lblStartTime.Text = Format(Value, "HH:mm:ss")
        End Set
    End Property

    Public WriteOnly Property FinishTime()
        Set(ByVal Value)
            lblFinishTime.Text = Format(Value, "HH:mm:ss")
        End Set
    End Property

    ' get Factory Name
    Public ReadOnly Property FactoryName() As String
        Get
            Return strFactoryName
        End Get
    End Property

    ' set and get Operator
    Public Property [Operator]() As String
        Get
            Return strOperator
        End Get
        Set(ByVal Value As String)
            Me.objUserInfo("Operator") = Value
            strOperator = Value
        End Set
    End Property

    ' get Workstation Id
    Public ReadOnly Property Workstation() As String
        Get
            Return strWorkstation
        End Get
    End Property

    ' set and get AuthorityLV
    Public Property AuthorityLV() As String
        Get
            Return Me.strAuthLV
        End Get
        Set(ByVal Value As String)
            Me.objUserInfo("AuthorityLV") = Value
            Me.strAuthLV = Value
        End Set
    End Property

    ' get SystemName
    Public ReadOnly Property SystemName() As String
        Get
            Return Me.strSystemName
        End Get
    End Property

    ' get FiledName
    Public ReadOnly Property FieldName() As String
        Get
            Return Me.strFieldName
        End Get
    End Property

    ' get User information
    Public ReadOnly Property UserInfo() As Hashtable
        Get
            Return Me.objUserInfo
        End Get
    End Property

#End Region

#Region " Private Function "
    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        '===== Set Start Time
        StartTime = Now
        Me.btnStart.Enabled = False
        Me.btnCancel.Enabled = False
        '===== Call Function RunBatch
        Me.RunBatch()
        '===== Set Stop Time
        FinishTime = Now
        Me.btnCancel.Enabled = True
        Me.btnCancel.Text = "Close"
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    ' on load event
    Private Sub DialogControl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

        Catch ex As Exception
            Me.lblMsg.Text = ex.Message
        End Try

    End Sub

    ' set system base information
    Private Sub SetBaseInfo()
        Try
            ' set information
            Me.strSystemName = Me.objUserInfo("SystemName").ToString()
            Me.strFactoryName = Me.objUserInfo("FactoryName").ToString()
            Me.strOperator = Me.objUserInfo("Operator").ToString()
            Me.strAuthLV = Me.objUserInfo("AuthorityLV").ToString()
            Me.strWorkstation = Me.objUserInfo("Workstation").ToString()
            Me.strFieldName = Me.objUserInfo("FieldName").ToString()

            Me.objDBBase.Conn = CType(Me.objUserInfo("Conn"), SqlConnection)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Public or Protected Function "
    Public Function EnabledBtnStart(ByVal bFlag As Boolean)
        Me.btnStart.Enabled = bFlag
    End Function

    Public Function SetMessage(ByVal strMessage As String)
        Me.lblMsg.Text = strMessage
    End Function

    Public Overridable Function RunBatch()
        MsgBox("Click [Start] on Class")
    End Function

    ' show message
    Protected Sub ShowMsg(ByVal pstrMsgNo As String)
        Try
            Me.lblMsg.Text = Me.GetPurposeVal(CmnUtil.GROUP_MSG, pstrMsgNo)
        Catch ex As Exception
            Me.lblMsg.Text = ex.Message
        End Try
    End Sub

    ' get item name
    Protected Function GetItemName(ByVal pstrItemNo As String) As String
        Dim strRtn As String

        Try
            strRtn = Me.GetPurposeVal(CmnUtil.GROUP_ITM, pstrItemNo)
        Catch ex As Exception
            Me.lblMsg.Text = ex.Message
        End Try
        ' retrun value
        Return strRtn
    End Function

    ' get button name
    Protected Function GetButtonName(ByVal pstrBtnNo As String) As String
        Dim strRtn As String

        Try
            strRtn = Me.GetPurposeVal(CmnUtil.GROUP_BTN, pstrBtnNo)
        Catch ex As Exception
            Me.lblMsg.Text = ex.Message
        End Try
        ' retrun value
        Return strRtn
    End Function

    ' get Purpose value
    Protected Function GetPurposeVal(ByVal pstrGroupCode As String, ByVal pstrCode As String) As String
        Dim objCmnMaster As CmnMasterTable
        Dim strRtn As String

        Try
            objCmnMaster = New CmnMasterTable(Me.SystemName)
            objCmnMaster.Conn = Me.objDBBase.Conn

            strRtn = objCmnMaster.GetName(pstrGroupCode, pstrCode)

        Catch ex As Exception
            Me.lblMsg.Text = ex.Message
        End Try
        ' retrun value
        Return strRtn
    End Function

#End Region

End Class
