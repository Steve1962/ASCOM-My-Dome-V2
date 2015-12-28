<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.labelDriverId = New System.Windows.Forms.Label()
        Me.buttonConnect = New System.Windows.Forms.Button()
        Me.buttonChoose = New System.Windows.Forms.Button()
        Me.Get_Az_Button = New System.Windows.Forms.Button()
        Me.Az_Label = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SLewtoAz_Button = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Shutter_Button = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'labelDriverId
        '
        Me.labelDriverId.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.labelDriverId.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.ASCOM.MyDome.My.MySettings.Default, "DriverId", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.labelDriverId.Location = New System.Drawing.Point(12, 37)
        Me.labelDriverId.Name = "labelDriverId"
        Me.labelDriverId.Size = New System.Drawing.Size(291, 21)
        Me.labelDriverId.TabIndex = 5
        Me.labelDriverId.Text = Global.ASCOM.MyDome.My.MySettings.Default.DriverId
        Me.labelDriverId.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'buttonConnect
        '
        Me.buttonConnect.Location = New System.Drawing.Point(316, 36)
        Me.buttonConnect.Name = "buttonConnect"
        Me.buttonConnect.Size = New System.Drawing.Size(72, 23)
        Me.buttonConnect.TabIndex = 4
        Me.buttonConnect.Text = "Connect"
        Me.buttonConnect.UseVisualStyleBackColor = True
        '
        'buttonChoose
        '
        Me.buttonChoose.Location = New System.Drawing.Point(316, 7)
        Me.buttonChoose.Name = "buttonChoose"
        Me.buttonChoose.Size = New System.Drawing.Size(72, 23)
        Me.buttonChoose.TabIndex = 3
        Me.buttonChoose.Text = "Choose"
        Me.buttonChoose.UseVisualStyleBackColor = True
        '
        'Get_Az_Button
        '
        Me.Get_Az_Button.Location = New System.Drawing.Point(12, 84)
        Me.Get_Az_Button.Name = "Get_Az_Button"
        Me.Get_Az_Button.Size = New System.Drawing.Size(75, 23)
        Me.Get_Az_Button.TabIndex = 6
        Me.Get_Az_Button.Text = "Get Status"
        Me.Get_Az_Button.UseVisualStyleBackColor = True
        '
        'Az_Label
        '
        Me.Az_Label.AutoSize = True
        Me.Az_Label.Location = New System.Drawing.Point(316, 110)
        Me.Az_Label.Name = "Az_Label"
        Me.Az_Label.Size = New System.Drawing.Size(48, 13)
        Me.Az_Label.TabIndex = 7
        Me.Az_Label.Text = "Az Label"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(316, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Label1"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(13, 114)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(74, 20)
        Me.TextBox1.TabIndex = 9
        '
        'SLewtoAz_Button
        '
        Me.SLewtoAz_Button.Location = New System.Drawing.Point(93, 112)
        Me.SLewtoAz_Button.Name = "SLewtoAz_Button"
        Me.SLewtoAz_Button.Size = New System.Drawing.Size(75, 23)
        Me.SLewtoAz_Button.TabIndex = 10
        Me.SLewtoAz_Button.Text = "SlewtoAz"
        Me.SLewtoAz_Button.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(261, 151)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Slewing"
        '
        'Shutter_Button
        '
        Me.Shutter_Button.Location = New System.Drawing.Point(13, 151)
        Me.Shutter_Button.Name = "Shutter_Button"
        Me.Shutter_Button.Size = New System.Drawing.Size(75, 23)
        Me.Shutter_Button.TabIndex = 12
        Me.Shutter_Button.Text = "Shutter"
        Me.Shutter_Button.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(316, 132)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "shutter stat label"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(259, 110)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Azimuth"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(316, 151)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(84, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "shutter stat label"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(259, 132)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(41, 13)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Shutter"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(13, 208)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(39, 13)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Label7"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(265, 178)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(39, 13)
        Me.Label8.TabIndex = 19
        Me.Label8.Text = "Label8"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(400, 233)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Shutter_Button)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.SLewtoAz_Button)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Az_Label)
        Me.Controls.Add(Me.Get_Az_Button)
        Me.Controls.Add(Me.labelDriverId)
        Me.Controls.Add(Me.buttonConnect)
        Me.Controls.Add(Me.buttonChoose)
        Me.Name = "Form1"
        Me.Text = "ASCOM MyDome Test App"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents labelDriverId As System.Windows.Forms.Label
    Private WithEvents buttonConnect As System.Windows.Forms.Button
    Private WithEvents buttonChoose As System.Windows.Forms.Button
    Friend WithEvents Get_Az_Button As System.Windows.Forms.Button
    Friend WithEvents Az_Label As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SerialPort1 As System.IO.Ports.SerialPort
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents SLewtoAz_Button As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Shutter_Button As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label

End Class
