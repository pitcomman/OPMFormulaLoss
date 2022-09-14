Imports Rist.OPMCmnClass
Imports OPMFormulaLossClass

Public Class FrmLogIn
    Inherits Rist.OPMCmnClass.PageBase

    Dim ActForm As Form
#Region " Windows Form Designer generated code "

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
            ' dispose event(connection close when Application ending)
            If MyBase.objDBBase.Conn.State = ConnectionState.Open Then
                MyBase.objDBBase.CloseConnection()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtOperator As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents KeyControl1 As Rist.OPMCmnClass.KeyControl
    Friend WithEvents lblKeyF1 As Rist.OPMCmnClass.FunctionKeyControl
    Friend WithEvents lblOperator As System.Windows.Forms.Label
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.lblOperator = New System.Windows.Forms.Label
        Me.txtOperator = New System.Windows.Forms.TextBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.lblPassword = New System.Windows.Forms.Label
        Me.KeyControl1 = New Rist.OPMCmnClass.KeyControl(Me.components)
        Me.lblKeyF1 = New Rist.OPMCmnClass.FunctionKeyControl
        Me.SuspendLayout()
        '
        'lblOperator
        '
        Me.lblOperator.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOperator.Location = New System.Drawing.Point(328, 240)
        Me.lblOperator.Name = "lblOperator"
        Me.lblOperator.Size = New System.Drawing.Size(120, 32)
        Me.lblOperator.TabIndex = 4
        Me.lblOperator.Text = "Operator"
        Me.lblOperator.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOperator
        '
        Me.txtOperator.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOperator.Location = New System.Drawing.Point(448, 240)
        Me.txtOperator.Name = "txtOperator"
        Me.txtOperator.Size = New System.Drawing.Size(152, 29)
        Me.txtOperator.TabIndex = 5
        Me.txtOperator.Text = ""
        '
        'txtPassword
        '
        Me.txtPassword.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(448, 296)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(152, 29)
        Me.txtPassword.TabIndex = 7
        Me.txtPassword.Text = ""
        '
        'lblPassword
        '
        Me.lblPassword.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.Location = New System.Drawing.Point(320, 296)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(128, 32)
        Me.lblPassword.TabIndex = 6
        Me.lblPassword.Text = "Password"
        Me.lblPassword.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'KeyControl1
        '
        '
        'lblKeyF1
        '
        Me.lblKeyF1.Location = New System.Drawing.Point(40, 656)
        Me.lblKeyF1.Name = "lblKeyF1"
        Me.lblKeyF1.Size = New System.Drawing.Size(152, 48)
        Me.lblKeyF1.TabIndex = 9
        '
        'FrmLogIn
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(1012, 710)
        Me.Controls.Add(Me.lblKeyF1)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtOperator)
        Me.Controls.Add(Me.lblOperator)
        Me.Name = "FrmLogIn"
        Me.Controls.SetChildIndex(Me.lblOperator, 0)
        Me.Controls.SetChildIndex(Me.txtOperator, 0)
        Me.Controls.SetChildIndex(Me.lblPassword, 0)
        Me.Controls.SetChildIndex(Me.txtPassword, 0)
        Me.Controls.SetChildIndex(Me.lblKeyF1, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Envents Function "

    Public Sub New()
        'MyBase.New(MaterialEntry.IndClass.PLUtil.SYSTEM_NAME)
        MyBase.New("OPMFormulaLoss")

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    ' on load event
    Private Sub FrmLogIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            MyBase.Title = MyBase.GetItemName("1001")
            MyBase.TitleFontSize = 20
            Me.lblKeyF1.Caption = "F1:" & MyBase.GetItemName("0003")
            Me.CloseCaption = "F12:" & MyBase.GetItemName("0009")
            MyBase.Message = ""
            Me.ShowInTaskbar = False
            Call SetLabel()
            Me.txtOperator.Focus()
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    ' keyup event
    Private Sub FrmLogIn_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Try
            Me.KeyControl1.Push(e.KeyValue)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    ' Activate event
    Private Sub FrmLogIn_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
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
        End
    End Function

    ' on click F1 label event
    Private Sub lblKeyF1_UCClick() Handles lblKeyF1.UCClick
        Try
            Me.KeyControl1.Push(Keys.F1)
        Catch ex As Exception
            MyBase.Message = ex.Message
        End Try
    End Sub

    ' on push F1 event
    Private Sub KeyControl1_PushF1() Handles KeyControl1.PushF1
        Try
            Call Authentication()
        Catch ex As CustomErrException
            MyBase.IsErrMsg = True
            MyBase.ShowMsg(ex.MsgCode)
        Catch ex As Exception
            MyBase.IsErrMsg = True
            MyBase.Message = ex.Message
        End Try
    End Sub

    Private Sub FrmLogIn_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Return Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

#End Region

#Region " Private Function "

    ' login
    Private Sub ConfirmLogIn()

        Dim objLogIn As LogInTable
        Dim dtData As DataTable = Nothing
        Try
            ' make new LogInTable instance
            objLogIn = New LogInTable
            objLogIn.Conn = MyBase.objDBBase.Conn
            dtData = objLogIn.GetLoginUserTbl(Me.txtOperator.Text, Me.txtPassword.Text)
            If Not dtData Is Nothing And dtData.Rows.Count > 0 Then
                MyBase.[Operator] = dtData.Rows(0)("OPName").ToString().Trim()
                MyBase.AuthorityLV = dtData.Rows(0)("AuthorityLV").ToString().Trim()
                ActForm = New FrmMainMenu(MyBase.UserInfo)


                ActForm.Show()
            Else
                Throw New CustomErrException("E001")
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub Authentication()
        Dim objAuth As Authentication
        Dim dtData As DataTable = Nothing
        Dim parameters As New System.Collections.Specialized.StringDictionary

        Dim ResultMessage As String
        Dim ResultCode As Integer


        Try
            ' make new LogInTable instance
            objAuth = New Authentication
            objAuth.Conn = MyBase.objDBBase.Conn


            parameters.Add("@OperatorID", Me.txtOperator.Text.ToString)
            parameters.Add("@Password", Me.txtPassword.Text.ToString)


            dtData = objAuth.AuthenticateOperator(parameters)
            If Not dtData Is Nothing And dtData.Rows.Count > 0 Then
                MyBase.[Operator] = dtData.Rows(0)("OPName").ToString().Trim()
                MyBase.AuthorityLV = dtData.Rows(0)("AuthorityLV").ToString().Trim()
                ResultMessage = dtData.Rows(0)("ResultMessage").ToString().Trim()
                ResultCode = dtData.Rows(0)("ResultCode").ToString().Trim()

                If ResultCode = 0 Then
                    ActForm = New FrmMainMenu(MyBase.UserInfo)
                    ActForm.Show()

                Else
                    Throw New Exception(ResultMessage)
                End If


            Else
                Throw New CustomErrException("E006")
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' set itemname into label
    Private Sub SetLabel()
        Try
            Me.lblOperator.Text = MyBase.GetItemName("0001") & " :"
            Me.lblPassword.Text = MyBase.GetItemName("0002") & " :"
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

#End Region


End Class
