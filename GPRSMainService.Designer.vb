Imports System.ServiceProcess
Imports System.ComponentModel
Imports System.Configuration.Install

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GPRSMainService
    Inherits System.ServiceProcess.ServiceBase

    'UserService overrides dispose to clean up the component list.
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

    ' The main entry point for the process
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New GPRSMainService}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        '
        'DCSMainService
        '


        Me.ServiceName = "GPRS Main Service"

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class

<RunInstaller(True)> Public Class ProjectInstaller

    Inherits Installer

    Private si As ServiceInstaller = New ServiceInstaller()

    Private m_Description As String

    Public Property Description() As String

        Get

            Return m_Description

        End Get

        Set(ByVal Value As String)

            m_Description = Value

        End Set

    End Property

    Public Sub New()

        Me.Description = "This service has been developed to  collect Data from AWS  by Partonegar Co.."

        Dim spi As ServiceProcessInstaller = New ServiceProcessInstaller()

        spi.Account = ServiceAccount.LocalSystem

        si = New ServiceInstaller()
        ' si.
        si.ServiceName = "GPRS Main Service" 'My.Resources.ServiceName

        si.DisplayName = "GPRS Main Service" 'My.Resources.DisplayName

        si.StartType = ServiceStartMode.Automatic
        Installers.AddRange(New Installer() {spi, si})

    End Sub

    Public Overrides Sub Install(ByVal v_IDstateserver As System.Collections.IDictionary)

        Dim rgkSystem As Microsoft.Win32.RegistryKey

        Dim rgkCurrentControlSet As Microsoft.Win32.RegistryKey

        Dim rgkServices As Microsoft.Win32.RegistryKey

        Dim rgkService As Microsoft.Win32.RegistryKey

        Dim rgkConfig As Microsoft.Win32.RegistryKey

        Try

            'Let the project installer do its job

            MyBase.Install(v_IDstateserver)

            'Open the HKEY_LOCAL_MACHINE\SYSTEM key

            rgkSystem = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("System")

            'Open CurrentControlSet

            rgkCurrentControlSet = rgkSystem.OpenSubKey("CurrentControlSet")

            'Go to the services key

            rgkServices = rgkCurrentControlSet.OpenSubKey("Services")

            'Open the key for your service, and allow writing

            rgkService = rgkServices.OpenSubKey(si.ServiceName, True)

            'Add your service's description as a REG_SZ value named "Description"

            rgkService.SetValue("Description", m_Description)

            rgkService.SetValue("Type", 272, Microsoft.Win32.RegistryValueKind.DWord)

            '(Optional) Add some custom information your service will use...

            rgkConfig = rgkService.CreateSubKey("Parameters")

            rgkConfig.Close()

            rgkService.Close()

            rgkCurrentControlSet.Close()

            rgkSystem.Close()

        Catch ex As Exception

            My.Application.Log.WriteEntry("An exception was thrown during service installation:" & vbCrLf & ex.ToString(), TraceEventType.Error)

            Console.WriteLine("An exception was thrown during service installation:" & vbCrLf & ex.ToString())

        End Try

    End Sub

    Public Overrides Sub Uninstall(ByVal v_IDstateserver As System.Collections.IDictionary)

        Dim rgkSystem As Microsoft.Win32.RegistryKey

        Dim rgkCurrentControlSet As Microsoft.Win32.RegistryKey

        Dim rgkServices As Microsoft.Win32.RegistryKey

        Dim rgkService As Microsoft.Win32.RegistryKey

        Try

            'Drill down to the service key and open it with write permission

            rgkSystem = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("System")

            rgkCurrentControlSet = rgkSystem.OpenSubKey("CurrentControlSet")

            rgkServices = rgkCurrentControlSet.OpenSubKey("Services")

            rgkService = rgkServices.OpenSubKey(si.ServiceName, True)

            'Delete any keys you created during installation (or that your service created)

            rgkService.DeleteSubKeyTree("Parameters")

            rgkService.Close()

            rgkCurrentControlSet.Close()

            rgkSystem.Close()

        Catch ex As Exception

            My.Application.Log.WriteEntry("Exception encountered while uninstalling service:" & vbCrLf & ex.ToString(), TraceEventType.Error)

            Console.WriteLine("Exception encountered while uninstalling service:" & vbCrLf & ex.ToString())

        Finally

            'Let the project installer do its job

            MyBase.Uninstall(v_IDstateserver)

        End Try

    End Sub

End Class