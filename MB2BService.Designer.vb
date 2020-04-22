Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MB2BService
    Inherits System.ServiceProcess.ServiceBase

    'UserService переопределяет метод Dispose для очистки списка компонентов.
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

    ' Главная точка входа процесса
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' В одном процессе может выполняться несколько служб NT. Для добавления
        ' службы в процесс измените следующую строку,
        ' чтобы создавался второй объект службы. Например,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New MB2BService}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Является обязательной для конструктора компонентов
    Private components As System.ComponentModel.IContainer

    ' Примечание: следующая процедура является обязательной для конструктора компонентов
    ' Для ее изменения используйте конструктор компонентов.  
    ' Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.EventLog1 = New System.Diagnostics.EventLog()
        Me.ModbusPort = New System.IO.Ports.SerialPort(Me.components)
        Me.SerialPort = New System.IO.Ports.SerialPort(Me.components)
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'EventLog1
        '
        Me.EventLog1.Log = "ModbusToBase"
        Me.EventLog1.Source = "ModbusToBase"
        '
        'ModbusPort
        '
        Me.ModbusPort.BaudRate = 115200
        Me.ModbusPort.PortName = "COM24"
        '
        'SerialPort
        '
        Me.SerialPort.BaudRate = 115200
        '
        'MB2BService
        '
        Me.ServiceName = "Service1"
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

    Friend WithEvents EventLog1 As EventLog
    Friend WithEvents ModbusPort As IO.Ports.SerialPort
    Friend WithEvents SerialPort As IO.Ports.SerialPort
End Class
