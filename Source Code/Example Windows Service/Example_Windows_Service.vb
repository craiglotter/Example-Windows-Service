Imports System.Serviceprocess
Imports Microsoft.Win32
Imports System.IO
Imports System.Management


Public Class Example_Windows_Service
    Inherits System.Serviceprocess.ServiceBase

    Dim submitfolder As String = ""

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call
    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.Serviceprocess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.Serviccurrentprocessess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.Serviceprocess.ServiceBase() {New Example_Windows_Service}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    Friend WithEvents Main_Timer As System.Timers.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Main_Timer = New System.Timers.Timer
        CType(Me.Main_Timer, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Main_Timer
        '
        Me.Main_Timer.Enabled = True
        Me.Main_Timer.Interval = 60000
        '
        'Example_Windows_Service
        '
        Me.ServiceName = "Example_Windows_Service"
        CType(Me.Main_Timer, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

#End Region

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            Activity_Logger("Service Successfully Started")
            Main_Timer.Start()
        Catch ex As Exception
            Error_Handler(ex, "OnStart")
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            Main_Timer.Stop()
            Activity_Logger("Service Successfully Stopped")
        Catch ex As Exception
            Error_Handler(ex, "OnStop")
        End Try
    End Sub

    

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Example Windows Service\Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Example Windows Service\Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
        Catch exc As Exception
            Dim mylog As New EventLog
            If Not mylog.SourceExists("Example Windows Service") Then
                mylog.CreateEventSource("Example Windows Service", "Example Windows Service Log")
            End If
            mylog.Source = "Example Windows Service"
            mylog.WriteEntry("Example Windows Service Log", "Error Handler Failure: " & exc.ToString, EventLogEntryType.Error)
            mylog.Close()
        End Try
    End Sub

    Public Sub Activity_Logger(ByVal message As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Example Windows Service\Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Example Windows Service\Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & message)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity_Logger")
        End Try
    End Sub


    Private Sub Main_Timer_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Main_Timer.Elapsed
        Try
            ' Code_Execute()
        Catch ex As Exception
            Error_Handler(ex, "Code_Execute")
        End Try
    End Sub

    
End Class
