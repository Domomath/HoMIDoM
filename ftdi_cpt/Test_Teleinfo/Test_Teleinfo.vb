Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class Test_Teleinfo

#Region "Windows Form Designer generated code"
    Inherits Form

    ' Clean up any resources being used.
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If components IsNot Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private Sub InitializeComponent()
        Me.Button_connecter = New System.Windows.Forms.Button()
        Me.Button_quitter = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label_messages_interface = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Buttoncpt1on = New System.Windows.Forms.RadioButton()
        Me.Buttoncpt2on = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Edit_port_com = New System.Windows.Forms.TextBox()
        Me.Label_messages = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_connecter
        '
        Me.Button_connecter.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Button_connecter.Location = New System.Drawing.Point(461, 134)
        Me.Button_connecter.Name = "Button_connecter"
        Me.Button_connecter.Size = New System.Drawing.Size(148, 33)
        Me.Button_connecter.TabIndex = 3
        Me.Button_connecter.Text = "Débuter lecture compteur"
        Me.Button_connecter.UseVisualStyleBackColor = False
        '
        'Button_quitter
        '
        Me.Button_quitter.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Button_quitter.Location = New System.Drawing.Point(461, 448)
        Me.Button_quitter.Name = "Button_quitter"
        Me.Button_quitter.Size = New System.Drawing.Size(148, 25)
        Me.Button_quitter.TabIndex = 1
        Me.Button_quitter.Text = "Quitter"
        Me.Button_quitter.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label_messages_interface)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.Edit_port_com)
        Me.Panel2.Location = New System.Drawing.Point(20, 24)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(589, 89)
        Me.Panel2.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(136, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 20)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Compteur :"
        '
        'Label_messages_interface
        '
        Me.Label_messages_interface.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label_messages_interface.Location = New System.Drawing.Point(8, 8)
        Me.Label_messages_interface.Name = "Label_messages_interface"
        Me.Label_messages_interface.Size = New System.Drawing.Size(577, 20)
        Me.Label_messages_interface.TabIndex = 1
        Me.Label_messages_interface.Text = "Label_messages_interface"
        Me.Label_messages_interface.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(16, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Port"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Panel3.Controls.Add(Me.Buttoncpt1on)
        Me.Panel3.Controls.Add(Me.Buttoncpt2on)
        Me.Panel3.Controls.Add(Me.RadioButton1)
        Me.Panel3.Location = New System.Drawing.Point(232, 40)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(337, 33)
        Me.Panel3.TabIndex = 0
        '
        'Buttoncpt1on
        '
        Me.Buttoncpt1on.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.Buttoncpt1on.Location = New System.Drawing.Point(9, 8)
        Me.Buttoncpt1on.Name = "Buttoncpt1on"
        Me.Buttoncpt1on.Size = New System.Drawing.Size(96, 16)
        Me.Buttoncpt1on.TabIndex = 0
        Me.Buttoncpt1on.Text = "Compteur 1"
        '
        'Buttoncpt2on
        '
        Me.Buttoncpt2on.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.Buttoncpt2on.Location = New System.Drawing.Point(112, 8)
        Me.Buttoncpt2on.Name = "Buttoncpt2on"
        Me.Buttoncpt2on.Size = New System.Drawing.Size(97, 16)
        Me.Buttoncpt2on.TabIndex = 1
        Me.Buttoncpt2on.Text = "Compteur 2"
        '
        'RadioButton1
        '
        Me.RadioButton1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.RadioButton1.Location = New System.Drawing.Point(209, 8)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(113, 16)
        Me.RadioButton1.TabIndex = 2
        Me.RadioButton1.Text = "aucun compteur"
        '
        'Edit_port_com
        '
        Me.Edit_port_com.Cursor = System.Windows.Forms.Cursors.No
        Me.Edit_port_com.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Edit_port_com.Location = New System.Drawing.Point(56, 45)
        Me.Edit_port_com.Name = "Edit_port_com"
        Me.Edit_port_com.ReadOnly = True
        Me.Edit_port_com.Size = New System.Drawing.Size(57, 21)
        Me.Edit_port_com.TabIndex = 1
        '
        'Label_messages
        '
        Me.Label_messages.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Label_messages.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold)
        Me.Label_messages.Location = New System.Drawing.Point(28, 136)
        Me.Label_messages.Name = "Label_messages"
        Me.Label_messages.Size = New System.Drawing.Size(365, 24)
        Me.Label_messages.TabIndex = 1
        Me.Label_messages.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.Label_messages.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.TextBox6)
        Me.Panel1.Controls.Add(Me.Label_messages)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Button_quitter)
        Me.Panel1.Controls.Add(Me.Button_connecter)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(630, 500)
        Me.Panel1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(461, 174)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(148, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(11, 174)
        Me.TextBox6.Multiline = True
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox6.Size = New System.Drawing.Size(435, 299)
        Me.TextBox6.TabIndex = 4
        Me.TextBox6.TabStop = False
        Me.TextBox6.WordWrap = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(20, 134)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(426, 20)
        Me.TextBox1.TabIndex = 6
        '
        'Test_Teleinfo
        '
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(631, 503)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.Location = New System.Drawing.Point(327, 159)
        Me.Name = "Test_Teleinfo"
        Me.Text = "TELEINFO tests"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Private components As System.ComponentModel.IContainer
    Public WithEvents Button_connecter As System.Windows.Forms.Button
    Public WithEvents Button_quitter As System.Windows.Forms.Button
    Public WithEvents Panel2 As System.Windows.Forms.Panel
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label_messages_interface As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Panel3 As System.Windows.Forms.Panel
    Public WithEvents Buttoncpt1on As System.Windows.Forms.RadioButton
    Public WithEvents Buttoncpt2on As System.Windows.Forms.RadioButton
    Public WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Public WithEvents Edit_port_com As System.Windows.Forms.TextBox
    Public WithEvents Label_messages As System.Windows.Forms.Label
    Public WithEvents Panel1 As System.Windows.Forms.Panel

#End Region

    Friend WithEvents comPort1 As IO.Ports.SerialPort
    Public Shared My_Names(20) As String
    Public Shared nom_interface As String = String.Empty
    Public Shared module_detecte As Boolean = False
    Public Shared num_comport As Integer = 0
    Public Const nom_interface1 As String = "Interface USB 1 TIC"
    Public Const nom_interface2 As String = "Interface XBEE -> Compteur"
    Public Const nom_interface3 As String = "Interface USB -> Compteur"
    Private WithEvents ftdi As ftdi_cpt.Ftdi_cpt

    Private Delegate Sub _Affiche_ASCII(ByVal donnee As String)

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub inforeceive(ByVal donnee As String) Handles ftdi.InfoReceived
        TextBox1.Text = donnee
    End Sub


    Public Sub quitter()

        ftdi.cpt(0)
        Me.Close()
    End Sub

    Public Sub Buttoncpt1onClick(ByVal Sender As System.Object, ByVal _e1 As System.EventArgs) Handles Buttoncpt1on.Click
        ftdi.cpt(1)
    End Sub

    Public Sub Buttoncpt2onClick(ByVal Sender As System.Object, ByVal _e1 As System.EventArgs) Handles Buttoncpt2on.Click
        ftdi.cpt(2)
    End Sub

    Public Sub Button_quitterClick(ByVal Sender As System.Object, ByVal _e1 As System.EventArgs) Handles Button_quitter.Click
        quitter()
    End Sub

    ' procedure executée au lancement du programme
    Public Sub Compteur()  'Panel1.Enter

        Try
            ftdi = New ftdi_cpt.Ftdi_cpt
            ftdi.Start("Interface USB 1 TIC")

            num_comport = ftdi.PortCOM
            If num_comport > 0 Then
                comPort1 = New IO.Ports.SerialPort
                comPort1.PortName = "COM" & (num_comport).ToString()
                Edit_port_com.Text = comPort1.PortName
            Else
                Edit_port_com.Text = "----"
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub info(ByVal data As String) Handles ftdi.InfoReceived
        Label_messages.Text = data
    End Sub
    Public Sub RadioButton1Click(ByVal Sender As System.Object, ByVal _e1 As System.EventArgs) Handles RadioButton1.Click
        ftdi.cpt(0)
    End Sub

    Public Sub Button_connecterClick(ByVal Sender As System.Object, ByVal _e1 As System.EventArgs) Handles Button_connecter.Click

        Try
            '@ Unsupported property or method(C): 'Connected'
            If (comPort1.IsOpen = False) Then
                '@ Unsupported property or method(C): 'Connected'

                comPort1.BaudRate = 1200
                comPort1.DataBits = 7
                comPort1.Parity = IO.Ports.Parity.Even
                comPort1.StopBits = IO.Ports.StopBits.One
                comPort1.DtrEnable = True
                comPort1.RtsEnable = True

                comPort1.Open()
                Button_connecter.Text = "arret lecture compteur"
                'Panel3.Enabled = False
            Else
                '@ Unsupported property or method(C): 'Connected'

                comPort1.DiscardInBuffer()
                comPort1.Dispose()
                'Threading.Thread.SpinWait(5000)
                'comPort1.Close()
                Button_connecter.Text = "Débuter lecture compteur"
                Panel3.Enabled = True
            End If
        Catch Ex As Exception
            MsgBox(" Bouton Connect Exception : " & Ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Compteur()
    End Sub

    Private Sub TForm1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Compteur()
    End Sub

    Private Sub DataReceived(ByVal sender As Object, ByVal e As IO.Ports.SerialDataReceivedEventArgs) Handles comPort1.DataReceived

        Try

            Dim BufferIn(8192) As Byte
            Dim count As Integer = 0

            Dim trame As String = ""

            ' Recherche du port du compteur Teleinfo

            ' Si le port est connecté - Reception des caracteres contenus dans le Buffer
            If comPort1.IsOpen Then
                count = comPort1.BytesToRead
                Try
                    If count > 0 Then
                        comPort1.Read(BufferIn, 0, count)

                        If Me.InvokeRequired Then
                            Me.Invoke(New _Affiche_ASCII(AddressOf Affiche_ASCII), System.Text.Encoding.ASCII.GetString(BufferIn))
                        End If
                    Else
                        MsgBox(" DataReceived  Pas de donnée recue")
                    End If
                Catch ex As Exception
                    MsgBox(" DataReceived  Exception : " & ex.Message)
                End Try
            End If

        Catch Ex As Exception
            MsgBox(" Datareceived  Exception : " & Ex.Message)
        End Try
    End Sub

    Private Sub Affiche_ASCII(ByVal donnee As String)
        Try

            TextBox6.Text = TextBox6.Text & donnee 'On affiche les données à l'écran

        Catch Ex As Exception
            MsgBox(" Datareceived  Exception : " & Ex.Message)
        End Try
    End Sub

    Public WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox


End Class ' 


