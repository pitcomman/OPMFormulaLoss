'***********************************************************************
' Program Name	    : Form Page Base Class
' Program ID	    : PageBase
' Function			: this Class have base function of windows form
' Create Date		: 2006/06/05
' Create Person		: Wattana
' 
' Supplement	    : 
' Version		    : 1.00
' ---------------------------------------------------------------------
' Condition     	: SqlServer2000, ADO.Net, .NetFramework
' Starting Way		: 
'***********************************************************************
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Imports System.Xml
Imports System.Text

Public Class PageBase
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

    ' default constracter
    Public Sub New()
        MyBase.New()
        ' make DBBase class instance
        objDBBase = New DBBase

        InitializeComponent()
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
            ' show error message
            Me.IsErrMsg = True
            Me.lblMsg.Text = ex.Message
        End Try
    End Sub

    ' custom constracter(for Login Form)
    Public Sub New(ByVal pstrSystemName As String)
        ' call base constractor
        MyBase.New()
        Try
            ' make DBBase class instance
            objDBBase = New DBBase(pstrSystemName)

            ' make User information class instance
            objUserInfo = New Hashtable

            ' get SystemName
            Me.strSystemName = pstrSystemName

            InitializeComponent()

            ' open connection
            If objDBBase Is Nothing Or objDBBase.Conn.State = ConnectionState.Closed Then
                objDBBase.OpenConnection()
            End If

            ' initialize control
            Me.InitControl()

            ' set base information
            Me.SetBaseInfo()

        Catch ex As Exception
            ' show error message
            Me.IsErrMsg = True
            Me.lblMsg.Text = ex.Message
        End Try
    End Sub

    ' common control
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblMsg As System.Windows.Forms.Label
    Friend WithEvents pnlHead As System.Windows.Forms.Panel
    Friend WithEvents lblCurDateTime As System.Windows.Forms.Label
    Friend WithEvents lblUserId As System.Windows.Forms.Label
    Friend WithEvents lblWSId As System.Windows.Forms.Label
    Friend WithEvents lblWId As System.Windows.Forms.Label
    Friend WithEvents lblUId As System.Windows.Forms.Label
    Friend WithEvents lblSystemCompany As System.Windows.Forms.Label
    Friend WithEvents lblVer As System.Windows.Forms.Label
    Friend WithEvents KeyControl1 As KeyControl
    Friend WithEvents PrintScreenControl1 As PrintScreenControl
    Private components As System.ComponentModel.IContainer
    Protected Friend WithEvents lblFkey12 As FunctionKeyControl
    Protected WithEvents lblVersion As System.Windows.Forms.Label

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.lblTitle = New System.Windows.Forms.Label
        Me.lblMsg = New System.Windows.Forms.Label
        Me.pnlHead = New System.Windows.Forms.Panel
        Me.lblSystemCompany = New System.Windows.Forms.Label
        Me.lblCurDateTime = New System.Windows.Forms.Label
        Me.lblUserId = New System.Windows.Forms.Label
        Me.lblWSId = New System.Windows.Forms.Label
        Me.lblWId = New System.Windows.Forms.Label
        Me.lblUId = New System.Windows.Forms.Label
        Me.lblVer = New System.Windows.Forms.Label
        Me.lblVersion = New System.Windows.Forms.Label
        Me.KeyControl1 = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.PrintScreenControl1 = New Rist.OPMCmnClass.PrintScreenControl(Me.components)
        Me.lblFkey12 = New Rist.OPMCmnClass.FunctionKeyControl
        Me.pnlHead.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(208, 24)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(552, 48)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Title"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMsg
        '
        Me.lblMsg.Font = New System.Drawing.Font("Arial", 15.75!)
        Me.lblMsg.ForeColor = System.Drawing.Color.Blue
        Me.lblMsg.Location = New System.Drawing.Point(0, 632)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(1008, 23)
        Me.lblMsg.TabIndex = 0
        Me.lblMsg.Text = "Message"
        Me.lblMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlHead
        '
        Me.pnlHead.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlHead.Controls.Add(Me.lblSystemCompany)
        Me.pnlHead.Controls.Add(Me.lblCurDateTime)
        Me.pnlHead.Controls.Add(Me.lblTitle)
        Me.pnlHead.Controls.Add(Me.lblUserId)
        Me.pnlHead.Controls.Add(Me.lblWSId)
        Me.pnlHead.Controls.Add(Me.lblWId)
        Me.pnlHead.Controls.Add(Me.lblUId)
        Me.pnlHead.Controls.Add(Me.lblVer)
        Me.pnlHead.Controls.Add(Me.lblVersion)
        Me.pnlHead.ForeColor = System.Drawing.SystemColors.ControlText
        Me.pnlHead.Location = New System.Drawing.Point(0, 0)
        Me.pnlHead.Name = "pnlHead"
        Me.pnlHead.Size = New System.Drawing.Size(1016, 88)
        Me.pnlHead.TabIndex = 2
        '
        'lblSystemCompany
        '
        Me.lblSystemCompany.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSystemCompany.ForeColor = System.Drawing.Color.MediumSeaGreen
        Me.lblSystemCompany.Location = New System.Drawing.Point(8, 8)
        Me.lblSystemCompany.Name = "lblSystemCompany"
        Me.lblSystemCompany.Size = New System.Drawing.Size(200, 20)
        Me.lblSystemCompany.TabIndex = 0
        Me.lblSystemCompany.Text = "SystemName"
        '
        'lblCurDateTime
        '
        Me.lblCurDateTime.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurDateTime.ForeColor = System.Drawing.Color.MediumSeaGreen
        Me.lblCurDateTime.Location = New System.Drawing.Point(780, 8)
        Me.lblCurDateTime.Name = "lblCurDateTime"
        Me.lblCurDateTime.Size = New System.Drawing.Size(208, 24)
        Me.lblCurDateTime.TabIndex = 0
        Me.lblCurDateTime.Text = "DateTime"
        '
        'lblUserId
        '
        Me.lblUserId.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserId.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblUserId.Location = New System.Drawing.Point(872, 48)
        Me.lblUserId.Name = "lblUserId"
        Me.lblUserId.Size = New System.Drawing.Size(128, 16)
        Me.lblUserId.TabIndex = 0
        Me.lblUserId.Text = "UserId"
        '
        'lblWSId
        '
        Me.lblWSId.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWSId.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblWSId.Location = New System.Drawing.Point(872, 64)
        Me.lblWSId.Name = "lblWSId"
        Me.lblWSId.Size = New System.Drawing.Size(128, 16)
        Me.lblWSId.TabIndex = 0
        Me.lblWSId.Text = "WSId"
        '
        'lblWId
        '
        Me.lblWId.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWId.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblWId.Location = New System.Drawing.Point(780, 64)
        Me.lblWId.Name = "lblWId"
        Me.lblWId.Size = New System.Drawing.Size(88, 16)
        Me.lblWId.TabIndex = 0
        Me.lblWId.Text = "WSId :"
        Me.lblWId.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblUId
        '
        Me.lblUId.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUId.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblUId.Location = New System.Drawing.Point(778, 48)
        Me.lblUId.Name = "lblUId"
        Me.lblUId.Size = New System.Drawing.Size(88, 16)
        Me.lblUId.TabIndex = 0
        Me.lblUId.Text = "UserId :"
        Me.lblUId.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVer
        '
        Me.lblVer.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVer.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblVer.Location = New System.Drawing.Point(780, 32)
        Me.lblVer.Name = "lblVer"
        Me.lblVer.Size = New System.Drawing.Size(88, 16)
        Me.lblVer.TabIndex = 0
        Me.lblVer.Text = "Version :"
        Me.lblVer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVersion
        '
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblVersion.Location = New System.Drawing.Point(872, 32)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(128, 16)
        Me.lblVersion.TabIndex = 0
        Me.lblVersion.Text = "Version"
        '
        'lblFkey12
        '
        Me.lblFkey12.Location = New System.Drawing.Point(808, 656)
        Me.lblFkey12.Name = "lblFkey12"
        Me.lblFkey12.Size = New System.Drawing.Size(184, 48)
        Me.lblFkey12.TabIndex = 0
        Me.lblFkey12.TabStop = False
        '
        'PageBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.LightCyan
        Me.ClientSize = New System.Drawing.Size(978, 676)
        Me.Controls.Add(Me.pnlHead)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.lblFkey12)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.KeyPreview = True
        Me.Name = "PageBase"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlHead.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Property Area"

    ' set Title
    Public WriteOnly Property Title() As String
        Set(ByVal Value As String)
            lblTitle.Text = Value
        End Set
    End Property

    ' set Message
    Public WriteOnly Property Message() As String
        Set(ByVal Value As String)
            lblMsg.Text = Value
        End Set
    End Property

    Public WriteOnly Property CloseCaption()
        Set(ByVal Value)
            lblFkey12.Caption = Value
        End Set
    End Property

    Public WriteOnly Property CloseFontSize() As Integer
        Set(ByVal Value As Integer)
            Me.lblFkey12.FontSize = Value
        End Set
    End Property

    Public WriteOnly Property TitleFontSize() As Integer
        Set(ByVal Value As Integer)
            Me.lblTitle.Font = New System.Drawing.Font("Tahoma", Value, Drawing.FontStyle.Bold)
        End Set
    End Property

    ' set Message Color
    Public WriteOnly Property IsErrMsg() As Boolean
        Set(ByVal Value As Boolean)
            If Value Then
                lblMsg.ForeColor = System.Drawing.Color.Red
            Else
                lblMsg.ForeColor = System.Drawing.Color.Blue
            End If
        End Set
    End Property

    ' set back ground color
    Public WriteOnly Property BGColor() As System.Drawing.Color
        Set(ByVal Value As System.Drawing.Color)
            Me.BackColor = Value
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

    ' on page open
    Protected Sub PageBase_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

        Catch ex As Exception
            ' show error message
            Me.IsErrMsg = True
            Me.lblMsg.Text = ex.Message
        End Try
    End Sub

    ' on class dispose
    'Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    '    If disposing Then
    '        If objDBBase.Conn.State = ConnectionState.Open Then
    '            objDBBase.CloseConnection()
    '        End If
    '    End If
    '    MyBase.Dispose(disposing)
    'End Sub

    ' initialize common control
    Private Sub InitControl()
        Dim objXml As XmlControl

        Try
            ' make new XmlControl class instance
            objXml = New XmlControl

            ' get and set Information
            objUserInfo.Add("FactoryName", objXml.GetFieldValue(Me.strSystemName, CmnUtil.FACTORY_NAME))
            objUserInfo.Add("Version", objXml.GetFieldValue(Me.strSystemName, CmnUtil.APP_VERSION))

            objUserInfo.Add("Operator", System.Environment.UserName)
            objUserInfo.Add("Workstation", System.Environment.MachineName)
            objUserInfo.Add("AuthorityLV", CmnUtil.VALUE_ZERO)

            Select Case objXml.GetFieldValue(Me.strSystemName, CmnUtil.FACTORY_NAME)
                Case CmnUtil.FACT_RIST
                    Me.strFieldName = CmnUtil.FIELD_ENG
                Case CmnUtil.FACT_REPI
                    Me.strFieldName = CmnUtil.FIELD_ENG
                Case CmnUtil.FACT_RAJP
                    Me.strFieldName = CmnUtil.FIELD_JPN
                Case Else
                    Throw New Exception("initial.xml is injustice")
            End Select

            ' set User information
            objUserInfo.Add("SystemName", Me.strSystemName)
            objUserInfo.Add("FieldName", Me.strFieldName)
            objUserInfo.Add("Conn", Me.objDBBase.Conn)

            ' set Database name of Purpose table
            CmnUtil.Purpose_DataBase_Name = objXml.GetFieldValue(Me.strSystemName, CmnUtil.PURPOSE_NAME).Trim

        Catch ex As Exception
            Throw ex
        Finally
            If Not objXml Is Nothing Then
                objXml = Nothing
            End If
        End Try
    End Sub

    ' set system base information
    Private Sub SetBaseInfo()
        Try
            ' set information
            Me.strSystemName = Me.objUserInfo("SystemName").ToString()
            Me.strFactoryName = Me.objUserInfo("FactoryName").ToString()
            Me.lblVersion.Text = Me.objUserInfo("Version").ToString()
            Me.strOperator = Me.objUserInfo("Operator").ToString()
            Me.strAuthLV = Me.objUserInfo("AuthorityLV").ToString()
            Me.strWorkstation = Me.objUserInfo("Workstation").ToString()
            Me.strFieldName = Me.objUserInfo("FieldName").ToString()

            Me.objDBBase.Conn = CType(Me.objUserInfo("Conn"), SqlConnection)

            Me.lblSystemCompany.Text = Me.strSystemName
            Me.lblCurDateTime.Text = Now.ToString("yyyy/MM/dd hh:mm:ss")
            Me.lblUserId.Text = Me.strOperator
            Me.lblWSId.Text = Me.strWorkstation

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

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
            Throw ex
        End Try
        ' retrun value
        Return strRtn
    End Function

    ' get Purpose value
    Protected Function GetPurposeVal(ByVal pstrFieldName As String, ByVal pstrGroupCode As String, ByVal pstrCode As String) As String
        Dim objCmnMaster As CmnMasterTable
        Dim strRtn As String

        Try
            objCmnMaster = New CmnMasterTable
            objCmnMaster.Conn = Me.objDBBase.Conn

            strRtn = objCmnMaster.GetName(pstrFieldName, pstrGroupCode, pstrCode)

        Catch ex As Exception
            Throw ex
        End Try
        ' retrun value
        Return strRtn
    End Function


    ' Get DataTable
    Public Function GetDataTableSql(ByVal prmStrSql As String) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmStrSql
            objSqlCmd.CommandType = CommandType.Text

            ' execute referance
            objDataTbl = Me.objDBBase.GetDataTable(objSqlCmd)

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objSqlCmd Is Nothing Then
                objSqlCmd.Dispose()
                objSqlCmd = Nothing
            End If
        End Try

        ' return value
        Return objDataTbl
    End Function

    ' Get DataTable (in transaction)
    Public Function GetDataTableSql(ByVal prmStrSql As String, ByVal pobjTran As SqlClient.SqlTransaction) As DataTable
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmStrSql
            objSqlCmd.CommandType = CommandType.Text
            objSqlCmd.Transaction = pobjTran

            ' execute referance
            objDataTbl = Me.objDBBase.GetDataTable(objSqlCmd)

        Catch ex As Exception
            Throw ex
        Finally
        End Try

        ' return value
        Return objDataTbl
    End Function

    ' get DataSet
    Public Function GetDataSetSql(ByVal prmStrSql As String) As DataSet

        Dim objDataSet As DataSet = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmStrSql
            objSqlCmd.CommandType = CommandType.Text

            ' execute referance
            objDataSet = Me.objDBBase.GetDataSet(objSqlCmd)

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objSqlCmd Is Nothing Then
                objSqlCmd.Dispose()
                objSqlCmd = Nothing
            End If
        End Try

        ' return value
        Return objDataSet

    End Function

    ' get DataSet (in transaction)
    Public Function GetDataSetSql(ByVal prmStrSql As String, ByVal pobjTran As SqlClient.SqlTransaction) As DataSet

        Dim objDataSet As DataSet = Nothing
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmStrSql
            objSqlCmd.CommandType = CommandType.Text
            objSqlCmd.Transaction = pobjTran

            ' execute referance
            objDataSet = Me.objDBBase.GetDataSet(objSqlCmd)

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objSqlCmd Is Nothing Then
                objSqlCmd.Dispose()
                objSqlCmd = Nothing
            End If
        End Try

        ' return value
        Return objDataSet

    End Function

    ' Run Sql command Text
    Public Function ExecSQLText(ByVal prmName As String) As Integer

        Dim intRet As Integer
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmName
            objSqlCmd.CommandType = CommandType.Text

            ' execute proc
            intRet = Me.objDBBase.ExecProc(objSqlCmd)

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objSqlCmd Is Nothing Then
                objSqlCmd.Dispose()
                objSqlCmd = Nothing
            End If
        End Try
        ' return value
        Return intRet
    End Function

    ' Run Sql command Text (in transaction)
    Public Function ExecSQLText(ByVal prmName As String, ByVal pobjTran As SqlTransaction) As Integer

        Dim intRet As Integer
        Dim objSqlCmd As SqlCommand = Nothing

        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.CommandText = prmName
            objSqlCmd.CommandType = CommandType.Text
            objSqlCmd.Transaction = pobjTran

            ' execute proc
            intRet = Me.objDBBase.ExecProc(objSqlCmd)

        Catch ex As Exception
            Throw ex
        Finally
            ' opening resource
            If Not objSqlCmd Is Nothing Then
                objSqlCmd.Dispose()
                objSqlCmd = Nothing
            End If
        End Try
        ' return value
        Return intRet
    End Function

    ' Get SchemaTbl(Meta information)
    Public Function GetSchemaTbl(ByVal prmStrSql As String) As DataTable
        Dim DRSCM As SqlDataReader
        Dim objDataTbl As DataTable = Nothing
        Dim objSqlCmd As SqlCommand = Nothing
        Try
            ' create command object
            objSqlCmd = New SqlCommand
            objSqlCmd.Connection = Me.objDBBase.Conn
            objSqlCmd.CommandText = prmStrSql
            objSqlCmd.CommandType = CommandType.Text
            DRSCM = objSqlCmd.ExecuteReader
            objDataTbl = DRSCM.GetSchemaTable
        Catch ex As Exception
            'System.Diagnostics.EventLog.WriteEntry("PIControlService", _
            '  "Failed {Function: getSchema [" & prmStrSql & "]}" & _
            '  Err.Description, EventLogEntryType.Error)
        Finally
            If Not DRSCM Is Nothing Then
                DRSCM.Close()
                DRSCM = Nothing
            End If
        End Try
        ' return value
        Return objDataTbl
    End Function


    Public Overridable Function PushF12()
        'To Do List
    End Function

    Private Sub lblFkey12_UCClick() Handles lblFkey12.UCClick
        PushF12()
    End Sub
End Class

