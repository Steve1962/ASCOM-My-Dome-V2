Public Class Form1

    Private driver As ASCOM.DriverAccess.Dome

    ''' <summary>
    ''' This event is where the driver is choosen. The device ID will be saved in the settings.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
    Private Sub buttonChoose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonChoose.Click
        My.Settings.DriverId = ASCOM.DriverAccess.Dome.Choose(My.Settings.DriverId)
        SetUIState()
    End Sub

    ''' <summary>
    ''' Connects to the device to be tested.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
    ''' 
    Private Sub buttonConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonConnect.Click

        If (IsConnected) Then
            driver.Connected = False
            Label1.Text = driver.Connected
            Timer1.Enabled = False
            Timer2.Enabled = False
        Else

            driver = New ASCOM.DriverAccess.Dome(My.Settings.DriverId)
                driver.Connected = True
                Label1.Text = driver.Connected
            Timer1.Enabled = True
            Timer1.Interval = 250

            Timer2.Enabled = True
            Timer2.Interval = 10000
                SetUIState()

        End If

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If IsConnected Then
            driver.Connected = False
        End If
        ' the settings are saved automatically when this application is closed.
    End Sub

    ''' <summary>
    ''' Sets the state of the UI depending on the device state
    ''' </summary>
    Private Sub SetUIState()
        buttonConnect.Enabled = Not String.IsNullOrEmpty(My.Settings.DriverId)
        buttonChoose.Enabled = Not IsConnected
        buttonConnect.Text = IIf(IsConnected, "Disconnect", "Connect")
    End Sub

    ''' <summary>
    ''' Gets a value indicating whether this instance is connected.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is connected; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Private ReadOnly Property IsConnected() As Boolean
        Get
            If Me.driver Is Nothing Then Return False
            Return driver.Connected
        End Get
    End Property

    ' TODO: Add additional UI and controls to test more of the driver being tested.

    Private Sub Get_Az_Button_Click(sender As Object, e As EventArgs) Handles Get_Az_Button.Click

        Az_Label.Text = driver.Azimuth

        Dim stat As Integer = driver.ShutterStatus
        Label3.Text = stat

        Label5.Text = driver.Slewing

        Select Case (stat)

            Case 0 Or 2

                Shutter_Button.Text = "CLOSE"

            Case 1 Or 3

                Shutter_Button.Text = "OPEN"

            Case 4

                Shutter_Button.Text = "ERROR"

        End Select

        Label8.Text = Now

    End Sub

    Private Sub SLewtoAz_Button_Click(sender As Object, e As EventArgs) Handles SLewtoAz_Button.Click

        'Timer1.Enabled = False
        Dim Az As Double = TextBox1.Text

        If (Az >= 0 And Az < 360) Then
            driver.SlewToAzimuth(Az)
            Label7.Text = ("Last Command - slew to " & Az & " at " & Now)


        Else

            Label7.Text = "Azimuth out of range - no command issued"

        End If

        'System.Threading.Thread.Sleep(250)
        ' Timer1.Enabled = True

    End Sub


    
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If IsConnected Then

            Az_Label.Text = driver.Azimuth

            Dim stat As Integer = driver.ShutterStatus

            Label3.Text = stat

            Label5.Text = driver.Slewing

            Select Case (stat)

                Case 0 Or 2

                    Shutter_Button.Text = "CLOSE"

                Case 1 Or 3

                    Shutter_Button.Text = "OPEN"

                Case 4

                    Shutter_Button.Text = "ERROR"

            End Select

            Label8.Text = Now
        End If


    End Sub

    Private Sub Shutter_Button_Click(sender As Object, e As EventArgs) Handles Shutter_Button.Click

        Select Case (Shutter_Button.Text)

            Case "OPEN"

                driver.OpenShutter()

            Case "CLOSE"

                driver.CloseShutter()

            Case "ERROR"

                MsgBox("Dome Shutter Error - it may be locked")

        End Select

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick


        Randomize()

        Dim rnd As New Random()

        Dim Az As Integer = rnd.Next(0, 360)

        driver.SlewToAzimuth(Az)

        Label7.Text = ("Last Command - slew to " & Az & " at " & Now)


    End Sub
End Class
