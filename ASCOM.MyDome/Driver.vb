'tabs=4
' --------------------------------------------------------------------------------
' This is a driver for an Arduino based dome rotation and shutter operation system.

'The codes used to communicate with the Arduino are in a string formatted as
'"997,command,998".

'The commands are 0-359 - SlewtoAzimuth, 800 - close shutter, 801 - open shutter
'808 - return dome status, 878 - find home and 999 Emergency Stop

'The Arduino responds to all messages with a standard 8 byte string termiated with a "#":
'The string is decoded as follows; ShutterStatus,Slewing, Status of Az Info(not used in the driver), "A", and finally a three digit domeAzimuth.
'The dome Azimuth returned by the Arduino is set to the target azimuth at the commencement of a slew.
'
' ASCOM Dome driver for My.Dome
'
' Description:	
'
' Implements:	ASCOM Dome interface version: 1.0
' Author:		Steve Loveridge
'
' Edit Log:
'
' Date			Who	Vers	Description
' -----------	---	-----	-------------------------------------------------------
' 21 12 2015	1.0.0	Initial edit, from Dome template
' ---------------------------------------------------------------------------------
'
'
' Your driver's ID is ASCOM.My.Dome
'
' The Guid attribute sets the CLSID for ASCOM.DeviceName.Dome
' The ClassInterface/None addribute prevents an empty interface called
' _Dome from being created and used as the [default] interface
'

' This definition is used to select code that's only applicable for one device type
#Const Device = "Dome"

Imports ASCOM
Imports ASCOM.Astrometry
Imports ASCOM.Astrometry.AstroUtils
Imports ASCOM.DeviceInterface
Imports ASCOM.Utilities

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text

<Guid("2d1d95b2-70cd-4efb-a541-d0eb0a155d90")> _
<ClassInterface(ClassInterfaceType.None)> _
Public Class Dome

    ' The Guid attribute sets the CLSID for ASCOM.My.Dome
    ' The ClassInterface/None addribute prevents an empty interface called
    ' _My from being created and used as the [default] interface


    Implements IDomeV2

    '
    ' Driver ID and descriptive string that shows in the Chooser
    '
    Public Const driverID As String = "ASCOM.My.Dome"
    Public Const driverDescription As String = "My Dome"
    Private domeSerialPort As ASCOM.Utilities.Serial


    Friend Shared comPortProfileName As String = "COM Port" 'Constants used for Profile persistence
    Friend Shared traceStateProfileName As String = "Trace Level"
    Friend Shared comPortDefault As String = My.Settings.COMPortName
    Friend Shared traceStateDefault As String = "False"

    
    Private connectedState As Boolean ' Private variable to hold the connected state
    Private utilities As Util ' Private variable to hold an ASCOM Utilities object
    Private astroUtilities As AstroUtils ' Private variable to hold an AstroUtils object to provide the Range method
    Private TL As TraceLogger ' Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)

    'EXTRA VARIABLES - added by author

    Dim localAzimuth As Double ' variables to hold driver.properties to minimise spamming of Arduino
    Dim localShutterStatus As Integer
    Dim localSlewing As String

    Dim serialCommsinProgress As Boolean = False ' a lock to avoid spaming the hardware

    Private serialQ As New Queue(Of String) ' queue object to handle Serial Port Comms

    ' Constructor - Must be public for COM registration!
    '
    Public Sub New()


        TL = New TraceLogger("", "My")
        TL.Enabled = My.Settings.TraceEnabled
        TL.LogMessage("Dome", "Starting initialisation")

        connectedState = False ' Initialise connected to false
        utilities = New Util() ' Initialise util object
        astroUtilities = New AstroUtils 'Initialise new astro utiliites object

        TL.LogMessage("Dome", "Completed initialisation")
    End Sub

    '
    ' PUBLIC COM INTERFACE IDomeV2 IMPLEMENTATION
    '

#Region "Common properties and methods"
    ''' <summary>
    ''' Displays the Setup Dialog form.
    ''' If the user clicks the OK button to dismiss the form, then
    ''' the new settings are saved, otherwise the old values are reloaded.
    ''' THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
    ''' </summary>
    Public Sub SetupDialog() Implements IDomeV2.SetupDialog
        ' consider only showing the setup dialog if not connected
        ' or call a different dialog if connected
        If IsConnected Then
            System.Windows.Forms.MessageBox.Show("Already connected, just press OK")
        End If

        Using F As SetupDialogForm = New SetupDialogForm()
            Dim result As System.Windows.Forms.DialogResult = F.ShowDialog()
            If result = DialogResult.OK Then

                My.Settings.Save()

            Else

                My.Settings.Reload()

            End If
        End Using
    End Sub

    Public ReadOnly Property SupportedActions() As ArrayList Implements IDomeV2.SupportedActions
        Get
            TL.LogMessage("SupportedActions Get", "Returning empty arraylist")
            Return New ArrayList()
        End Get
    End Property

    Public Function Action(ByVal ActionName As String, ByVal ActionParameters As String) As String Implements IDomeV2.Action
        Throw New ActionNotImplementedException("Action " & ActionName & " is not supported by this driver")
    End Function

    Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False) Implements IDomeV2.CommandBlind
        CheckConnected("CommandBlind")
        ' Call CommandString and return as soon as it finishes
        Me.CommandString(Command, Raw)
        ' or
        Throw New MethodNotImplementedException("CommandBlind")
    End Sub

    Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean _
        Implements IDomeV2.CommandBool
        CheckConnected("CommandBool")
        Dim ret As String = CommandString(Command, Raw)
        ' TODO decode the return string and return true or false
        ' or
        Throw New MethodNotImplementedException("CommandBool")
    End Function

    Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String _
        Implements IDomeV2.CommandString
        CheckConnected("CommandString")
        ' it's a good idea to put all the low level communication with the device here,
        ' then all communication calls this function
        ' you need something to ensure that only one command is in progress at a time
        Throw New MethodNotImplementedException("CommandString")
    End Function

    Public Property Connected() As Boolean Implements IDomeV2.Connected

        Get
            TL.LogMessage("Connected Get", IsConnected.ToString())
            Return IsConnected
        End Get

        Set(value As Boolean)
            TL.LogMessage("Connected Set", value.ToString())
            If value = IsConnected Then
                Return
            End If

            If value Then

                TL.LogMessage("Connected Set", "Connecting to port " + My.Settings.COMPortName)
                ' connect to the device
                Dim comPort As String = My.Settings.COMPortName
                domeSerialPort = New ASCOM.Utilities.Serial()
                domeSerialPort.PortName = comPort
                domeSerialPort.Speed = 57600
                domeSerialPort.ReceiveTimeout = 8

                Try
                    domeSerialPort.Connected = True
                    connectedState = True
                Catch

                    MsgBox("Cannot connect to COM Port")
                    connectedState = False
                    domeSerialPort.Connected = False

                End Try

            Else
                connectedState = False
                TL.LogMessage("Connected Set", "Disconnecting from port " + My.Settings.COMPortName)
                ' disconnect from the device
                domeSerialPort.Connected = False
                domeSerialPort.Dispose()
                domeSerialPort = Nothing

            End If
        End Set
    End Property

    Public ReadOnly Property Description As String Implements IDomeV2.Description
        Get
            ' this pattern seems to be needed to allow a public property to return a private field
            Dim d As String = driverDescription
            TL.LogMessage("Description Get", d)
            Return d
        End Get
    End Property

    Public ReadOnly Property DriverInfo As String Implements IDomeV2.DriverInfo
        Get
            Dim m_version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            ' TODO customise this driver description
            Dim s_driverInfo As String = "Information about the driver itself. Version: " + m_version.Major.ToString() + "." + m_version.Minor.ToString()
            TL.LogMessage("DriverInfo Get", s_driverInfo)
            Return s_driverInfo
        End Get
    End Property

    Public ReadOnly Property DriverVersion() As String Implements IDomeV2.DriverVersion
        Get
            ' Get our own assembly and report its version number
            TL.LogMessage("DriverVersion Get", Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2))
            Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2)
        End Get
    End Property

    Public ReadOnly Property InterfaceVersion() As Short Implements IDomeV2.InterfaceVersion
        Get
            TL.LogMessage("InterfaceVersion Get", "2")
            Return 2
        End Get
    End Property

    Public ReadOnly Property Name As String Implements IDomeV2.Name
        Get
            Dim s_name As String = "MyDome@PerryBarn"
            TL.LogMessage("Name Get", s_name)
            Return s_name
        End Get
    End Property

    Public Sub Dispose() Implements IDomeV2.Dispose
        ' Clean up the tracelogger and util objects
        TL.Enabled = False
        TL.Dispose()
        TL = Nothing
        utilities.Dispose()
        utilities = Nothing
        astroUtilities.Dispose()
        astroUtilities = Nothing
    End Sub

#End Region

#Region "IDome Implementation"

    Private domeShutterState As Boolean = False ' Variable to hold the open/closed status of the shutter, true = Open

    Public Sub AbortSlew() Implements IDomeV2.AbortSlew
        '''''' This is a mandatory parameter but we have no action to take in this simple driver
        SendArduinoCommand(999)


        TL.LogMessage("AbortSlew", "Completed")
    End Sub

    Public ReadOnly Property Altitude() As Double Implements IDomeV2.Altitude
        Get
            TL.LogMessage("Altitude Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("Altitude", False)
        End Get
    End Property

    Public ReadOnly Property AtHome() As Boolean Implements IDomeV2.AtHome
        Get

            SendArduinoCommand(808)

            If localAzimuth = 0 Then

                Return True

                TL.LogMessage("AtHome", True)

            Else

                Return False

                TL.LogMessage("AtHome", False)

            End If


        End Get
    End Property

    Public ReadOnly Property AtPark() As Boolean Implements IDomeV2.AtPark
        Get
            TL.LogMessage("AtPark", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("AtPark", False)
        End Get
    End Property

    Public ReadOnly Property Azimuth() As Double Implements IDomeV2.Azimuth
        Get

            SendArduinoCommand(808)

            Return localAzimuth
            TL.LogMessage("Azimuth Returned: ", localAzimuth)

        End Get
    End Property

    Public ReadOnly Property CanFindHome() As Boolean Implements IDomeV2.CanFindHome
        Get
            TL.LogMessage("CanFindHome Get", True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanPark() As Boolean Implements IDomeV2.CanPark
        Get
            TL.LogMessage("CanPark Get", False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetAltitude() As Boolean Implements IDomeV2.CanSetAltitude
        Get
            TL.LogMessage("CanSetAltitude Get", False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetAzimuth() As Boolean Implements IDomeV2.CanSetAzimuth
        Get
            TL.LogMessage("CanSetAzimuth Get", True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSetPark() As Boolean Implements IDomeV2.CanSetPark
        Get
            TL.LogMessage("CanSetPark Get", False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetShutter() As Boolean Implements IDomeV2.CanSetShutter
        Get
            TL.LogMessage("CanSetShutter Get", True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSlave() As Boolean Implements IDomeV2.CanSlave
        Get
            TL.LogMessage("CanSlave Get", False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSyncAzimuth() As Boolean Implements IDomeV2.CanSyncAzimuth
        Get
            TL.LogMessage("CanSyncAzimuth Get", False.ToString())
            Return False
        End Get
    End Property

    Public Sub CloseShutter() Implements IDomeV2.CloseShutter

        SendArduinoCommand(800)

        TL.LogMessage("CloseShutter", "Shutter has been closed")
        domeShutterState = False
    End Sub

    Public Sub FindHome() Implements IDomeV2.FindHome

        SendArduinoCommand(878)
       
        TL.LogMessage("FindHome", "Home Found")

    End Sub

    Public Sub OpenShutter() Implements IDomeV2.OpenShutter

        SendArduinoCommand(801)


        TL.LogMessage("OpenShutter", "Shutter has been opened")
        domeShutterState = True
    End Sub

    Public Sub Park() Implements IDomeV2.Park
        TL.LogMessage("Park", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("Park")
    End Sub

    Public Sub SetPark() Implements IDomeV2.SetPark
        TL.LogMessage("SetPark", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SetPark")
    End Sub

    Public ReadOnly Property ShutterStatus() As ShutterState Implements IDomeV2.ShutterStatus
        Get

            SendArduinoCommand(808)

            Return localShutterStatus

            TL.LogMessage("ShutterStatus Returned: ", localShutterStatus)
            
        End Get
    End Property

    Public Property Slaved() As Boolean Implements IDomeV2.Slaved
        Get
            TL.LogMessage("Slaved Get", False.ToString())
            Return False
        End Get
        Set(value As Boolean)
            TL.LogMessage("Slaved Set", "not implemented")
            Throw New ASCOM.PropertyNotImplementedException("Slaved", True)
        End Set
    End Property

    Public Sub SlewToAltitude(Altitude As Double) Implements IDomeV2.SlewToAltitude
        TL.LogMessage("SlewToAltitude", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToAltitude")
    End Sub

    Public Sub SlewToAzimuth(Azimuth As Double) Implements IDomeV2.SlewToAzimuth

        If Azimuth >= 0 And Azimuth < 360 Then

            SendArduinoCommand(Azimuth)

            TL.LogMessage("SlewToAzimuth", Azimuth)

        Else

            Throw New ASCOM.InvalidValueException(("Azimuth - " & Azimuth))

            TL.LogMessage("SlewToAzimuth - out of range - ", Azimuth)

        End If


    End Sub

    Public ReadOnly Property Slewing() As Boolean Implements IDomeV2.Slewing
        Get

           
            SendArduinoCommand(808)

            Dim slew As Boolean

            If localSlewing = "1" Then slew = True


            If localSlewing = "0" Then slew = False

            Return slew

            TL.LogMessage("Slewing Returned: ", slew)

        End Get
    End Property

    Public Sub SyncToAzimuth(Azimuth As Double) Implements IDomeV2.SyncToAzimuth
        TL.LogMessage("SyncToAzimuth", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SyncToAzimuth")
    End Sub

#End Region

#Region "Private properties and methods"
    ' here are some useful properties and methods that can be used as required
    ' to help with

#Region "ASCOM Registration"

    Private Shared Sub RegUnregASCOM(ByVal bRegister As Boolean)

        Using P As New Profile() With {.DeviceType = "Dome"}
            If bRegister Then
                P.Register(driverID, driverDescription)
            Else
                P.Unregister(driverID)
            End If
        End Using

    End Sub

    <ComRegisterFunction()> _
    Public Shared Sub RegisterASCOM(ByVal T As Type)

        RegUnregASCOM(True)

    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub UnregisterASCOM(ByVal T As Type)

        RegUnregASCOM(False)

    End Sub

#End Region

    ''' <summary>
    ''' Returns true if there is a valid connection to the driver hardware
    ''' </summary>
    Private ReadOnly Property IsConnected As Boolean
        Get
            ' check that the driver hardware connection exists and is connected to the hardware

            If Not domeSerialPort Is Nothing Then
                If domeSerialPort.Connected Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If

        End Get
    End Property

    ''' <summary>
    ''' Use this function to throw an exception if we aren't connected to the hardware
    ''' </summary>
    ''' <param name="message"></param>
    Private Sub CheckConnected(ByVal message As String)
        If Not IsConnected Then
            Throw New NotConnectedException(message)
        End If
    End Sub


    ''construct and send string to Arduino

    Private Sub SendArduinoCommand(ByVal command As String)

        'add last command received to queue

        serialQ.Enqueue(command)

        Dim s As String = Nothing

        'do until there's nothing in the queue

        Do Until serialQ.Count = 0

            'only proceed if the Serial Port isn't busy

            If serialCommsinProgress = False Then

                serialCommsinProgress = True

                ' load the first item in the queue

                Dim c As String = serialQ.Peek()

                'send that command string to the arduino

                domeSerialPort.Transmit("997," & c & ",998" & vbLf)


                Dim retries As Short = 0

                Try

                    s = domeSerialPort.ReceiveTerminated("#")

                Catch
                    retries += 1
                    If retries <= 2 Then
                        System.Threading.Thread.Sleep(500)
                        domeSerialPort.Transmit("997," & c & ",998" & vbLf)
                        s = domeSerialPort.ReceiveTerminated("#")
                       

                    Else

                        MsgBox("Serial error after 3 retries")

                    End If


                End Try
                'reset the flag to indicte that the serial port is free
                serialCommsinProgress = False
                'Remove the first item from the queue
                serialQ.Dequeue()

            End If
        Loop

        ' if s is valid extract the members and save as local variables

        If Not s = Nothing Then

            'if the response is too long - trim to 8 bytes
            If s.Length > 8 Then
                s = s.Substring(s.Length - 8, 8)
            End If

            Dim az As String
            az = s.Substring(4, 3)

            localAzimuth = Convert.ToDouble(az)

            Dim sh As String
            sh = s.Substring(0, 1)
            localShutterStatus = Convert.ToInt16(sh)

         localSlewing = s.Substring(1, 1)




        End If

    End Sub

    

#End Region

End Class
