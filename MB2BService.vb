Imports System.Data.OleDb

Public Class MB2BService
    Dim GetDataIsProgress As Boolean
    Private eventId As Integer = 0
    Private ReadOnly debug As Boolean = True
    Private ReadOnly timer As Timers.Timer = New Timers.Timer With {
            .Enabled = True,
            .Interval = 1 * 60000,
            .AutoReset = True
        }
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Добавьте здесь код запуска службы. Этот метод должен настроить все необходимое
        ' для успешной работы службы.
        GetDataIsProgress = False
        If eventId > 60000 Then eventId = 0
        If debug Then EventLog1.WriteEntry("Служба запущена", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
        SaveAlarmToBase("Служба запущена")
        'Создаем таймер 
        'Запускаем Ontimer при срабатывании таймера 
        AddHandler timer.Elapsed, AddressOf OnTimer

        timer.Start()
        '        GetData()
    End Sub

    Public Sub OnTimer(sender As Object, args As Timers.ElapsedEventArgs)
        If eventId > 60000 Then eventId = 0
        If debug Then EventLog1.WriteEntry(Now.ToString & " Запуск события по таймеру", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
        If Not GetDataIsProgress Then
            GetData()
        Else
            If debug Then EventLog1.WriteEntry(Now.ToString & " Предыдущий опрос не окончен", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            SaveAlarmToBase("Предыдущий опрос не окончен")
            'GetDataIsProgress = False
        End If
    End Sub

    Protected Overrides Sub OnStop()
        timer.Stop()
        timer.Close()
        timer.Dispose()
        If debug Then EventLog1.WriteEntry("Служба остановлена", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
        SaveAlarmToBase("Служба остановлена")


    End Sub

    Public Sub GetData()

        Dim BaudRateFromDB, idFromDB As Integer

        Dim ComPortNameFromDB, AddressFromDB, DescriptionFromDB As String

        Dim CollectorIsActive As Boolean

        Dim Collectors As New List(Of Collector)()

        Dim dbConn As New OleDbConnection()

        GetDataIsProgress = True

        Try
            ' Подключаемся к базе
            dbConn.ConnectionString = "Provider=SQLOLEDB; Data Source=localhost\SQLEXPRESS; Initial Catalog=MB2B; Persist Security Info=False; User Id=mb2b; Password=mb2b"
            dbConn.Open()
        Catch ex As Exception
            If debug Then EventLog1.WriteEntry("Ошибка соединения с базой данных: " + ex.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            dbConn.Close()
            GetDataIsProgress = False
            Exit Sub
        End Try

        Try
            ' Делаем запрос к базе 
            Dim dbCommandCollectors As OleDbCommand = dbConn.CreateCommand()
            dbCommandCollectors.CommandText = "SELECT * FROM Settings"
            dbCommandCollectors.ExecuteNonQuery()
            Dim reader As OleDbDataReader = dbCommandCollectors.ExecuteReader()
            ' Читаем данные
            While (reader.Read())
                ' Если пустое значение, то присваиваем переменной 0

                If Not reader.IsDBNull(reader.GetOrdinal("id")) Then
                    idFromDB = reader(reader.GetOrdinal("id"))
                Else
                    idFromDB = 0
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("isActive")) Then
                    CollectorIsActive = reader(reader.GetOrdinal("isActive"))
                Else
                    CollectorIsActive = False
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("ComPortName")) Then
                    ComPortNameFromDB = reader(reader.GetOrdinal("ComPortName"))
                Else
                    ComPortNameFromDB = ""
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("Address")) Then
                    AddressFromDB = reader(reader.GetOrdinal("Address"))
                Else
                    AddressFromDB = ""
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("Description")) Then
                    DescriptionFromDB = reader(reader.GetOrdinal("Description"))
                Else
                    DescriptionFromDB = ""
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("BaudRate")) Then
                    BaudRateFromDB = reader(reader.GetOrdinal("BaudRate"))
                Else
                    BaudRateFromDB = 0
                End If

                If debug Then EventLog1.WriteEntry(Now.ToString _
                                     & " Добавлено устройство  " _
                                     & " ID " & idFromDB & vbCrLf _
                                     & " Компорт: " & ComPortNameFromDB & vbCrLf _
                                     & " Скорость компорта: " & BaudRateFromDB & vbCrLf _
                                     & " Адрес устройства: " & AddressFromDB & vbCrLf _
                                     & " Описание: " & DescriptionFromDB & vbCrLf _
                                     , EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))

                If CollectorIsActive Then Collectors.Add(New Collector() With {.Id = idFromDB,
                                                     .IsActive = CollectorIsActive,
                                                     .ComPortName = ComPortNameFromDB,
                                                     .BaudRate = BaudRateFromDB,
                                                     .Address = AddressFromDB,
                                                     .Description = DescriptionFromDB})
            End While
            ' закрываем ридер
            reader.Close()
        Catch exdb As Exception
            If debug Then EventLog1.WriteEntry("Ошибка чтения БД: " + exdb.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            SaveAlarmToBase("Ошибка чтения БД: " + exdb.Message)
            dbConn.Close()
            GetDataIsProgress = False
            Exit Sub
        End Try
        dbConn.Close()
        ' Начинаем опрашивать счетчики
        For Each Collector As Collector In Collectors
            Try
                If debug Then EventLog1.WriteEntry("ID: " & Collector.Id & " Компорт: " & Collector.ComPortName & "   Адрес счетчика: " & Collector.Address & " Скорость порта: " & Collector.BaudRate, EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                If Not ModbusPort.IsOpen Then
                    ' Открываем компорт
                    ModbusPort.PortName = Collector.ComPortName
                    ModbusPort.ReadTimeout = 5000
                    ModbusPort.BaudRate = Collector.BaudRate
                    Try
                        ModbusPort.Open()
                        ModbusPort.DiscardInBuffer()
                        ModbusPort.DiscardOutBuffer()
                    Catch ex As Exception
                        If debug Then EventLog1.WriteEntry("Ошибка открытия порта " & Collector.ComPortName, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                        SaveAlarmToBase("Ошибка открытия порта " & Collector.ComPortName)
                        ModbusPort.Close()
                        Continue For
                    End Try
                End If
                Try
                    Dim master As Modbus.Device.IModbusSerialMaster = Modbus.Device.ModbusSerialMaster.CreateRtu(ModbusPort)
                    master.Transport.Retries = 0
                    master.Transport.ReadTimeout = 500  'millionsecs

                    Dim DataFromDevice As UShort() = ModbusDataReader(master, Collector.Address, 512, 6)
                    If DataFromDevice.Length <> 0 Then
                        Dim iBacketCounter As ULong = DataFromDevice(1) * &HFFFF + DataFromDevice(0)
                        Dim iWaterGVSCollector As ULong = DataFromDevice(3) * &HFFFF + DataFromDevice(2)
                        Dim iWaterHeatingCollector As ULong = DataFromDevice(5) * &HFFFF + DataFromDevice(4)

                        If debug Then EventLog1.WriteEntry(Now.ToString _
                                                               & " Компорт: " & Collector.ComPortName & vbCrLf _
                                                               & " Адрес устройства: " & Collector.Address & vbCrLf _
                                                               & " Счетчик ковшей: " & iBacketCounter & vbCrLf _
                                                               & " Счетчик горячей воды: " & iWaterGVSCollector & vbCrLf _
                                                               & " Счетчик воды отопления: " & iWaterHeatingCollector & vbCrLf _
                                                                 , EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))

                        SavetoBase(Collector.Id, iBacketCounter, iWaterGVSCollector, iWaterHeatingCollector)

                        ModbusPort.Close()
                    End If
                Catch Ex As Exception
                    If debug Then EventLog1.WriteEntry(Ex.Message, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                    SaveAlarmToBase("Компорт " & Collector.ComPortName & " Адрес " & Collector.Address & Ex.Message)
                    ModbusPort.Close()
                    Continue For
                End Try

            Catch Ex As Exception
                If debug Then EventLog1.WriteEntry(Ex.Message, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                SaveAlarmToBase(Ex.Message)
                ModbusPort.Close()
                Continue For
            End Try
        Next
        GetDataIsProgress = False
    End Sub
    Sub SaveAlarmToBase(AlarmString As String)
        Dim dbConn As New OleDbConnection()

        Try
            dbConn.ConnectionString = "Provider=SQLOLEDB; Data Source=localhost\SQLEXPRESS; Initial Catalog=MB2B; Persist Security Info=False; User Id=mb2b; Password=mb2b"
            dbConn.Open()
        Catch ex As Exception
            If debug Then EventLog1.WriteEntry("Ошибка соединения с базой данных: " + ex.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            Exit Sub
        End Try

        Try
            ' Делаем запрос
            Dim dbCommand As OleDbCommand = dbConn.CreateCommand()
            dbCommand.CommandText = "insert into Alarm ([datetime], [AlarmString]) VALUES (getdate(), ?)"
            dbCommand.Parameters.Add("AlarmString", OleDbType.[VarChar]).Value = AlarmString
            dbCommand.ExecuteNonQuery()
            dbConn.Close()
            Exit Sub
        Catch exdb As Exception
            If debug Then EventLog1.WriteEntry("Ошибка записи ошибки в БД: " + exdb.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            dbConn.Close()
            Exit Sub
        End Try
        ' Закрываем соединение с базой
        dbConn.Close()
    End Sub


    Public Function ModbusDataReader(master As Modbus.Device.IModbusSerialMaster, SlaveAddress As String, StartRegister As String, PointsNumber As UShort) As UShort()

        Dim holding_register As UShort()
        Dim slaveId As Byte = Byte.Parse(SlaveAddress)
        Dim startAddress As UShort = UShort.Parse(StartRegister)

        Try
            holding_register = master.ReadHoldingRegisters(slaveId, startAddress, PointsNumber)
        Catch ex As Exception
            holding_register = {}
            MsgBox(ex.Message & " Неправильный порт или адрес " & " Адрес: " & SlaveAddress)
            If debug Then EventLog1.WriteEntry(ex.Message & " Неправильный порт или адрес " & " Адрес: " & SlaveAddress, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            SaveAlarmToBase(ex.Message & " Неправильный порт или адрес " & " Адрес: " & SlaveAddress)
        End Try

        Return holding_register
    End Function

    Public Sub ModbusDataWriter(master As Modbus.Device.IModbusSerialMaster, SlaveAddress As String, StartRegister As String, WriteData As UShort())
        Dim slaveId As Byte = Byte.Parse(SlaveAddress)
        Dim startAddress As UShort = UShort.Parse(StartRegister)
        Try
            master.WriteMultipleRegisters(slaveId, startAddress, WriteData)
        Catch ex As Exception
            If debug Then EventLog1.WriteEntry(ex.Message & " Неправильный порт или адрес " & " Адрес: " & SlaveAddress, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            SaveAlarmToBase(ex.Message & " Неправильный порт или адрес " & " Адрес: " & SlaveAddress)
        End Try
    End Sub

    Public Function SavetoBase(CollectorId As Integer, BacketCounter As ULong, WaterGVSCounter As ULong, WaterHeatingCounter As ULong) As Boolean
        Dim dbConn As New OleDbConnection()
        ' Подключаемся к базе для записи данных

        If BacketCounter = 0 Or WaterGVSCounter = 0 Or WaterHeatingCounter = 0 Then
            If debug Then
                EventLog1.WriteEntry("Ошибка получения данных " & " Счетчик ковшей: " & BacketCounter & " Счетчик горячей воды: " & WaterGVSCounter & " Счетчик воды отопления: " & WaterHeatingCounter _
                                 , EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                SaveAlarmToBase("Ошибка получения данных " & " Счетчик ковшей: " & BacketCounter & " Счетчик горячей воды: " & WaterGVSCounter & " Счетчик воды отопления: " & WaterHeatingCounter)
            End If
            Return False
        End If
        Try
            dbConn.ConnectionString = "Provider=SQLOLEDB; Data Source=localhost\SQLEXPRESS; Initial Catalog=MB2B; Persist Security Info=False; User Id=mb2b; Password=mb2b"
            dbConn.Open()
        Catch ex As Exception
            If debug Then EventLog1.WriteEntry("Ошибка соединения с базой данных: " + ex.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            Return False
        End Try

        Try
            ' Делаем запрос
            Dim dbCommand As OleDbCommand = dbConn.CreateCommand()
            dbCommand.CommandText = "insert into AccuCollectors ([datetime], [Collectorid],[BacketCounter],[WaterGVSCounter],[WaterHeatingCounter]) VALUES (getdate(), ?, ?, ?, ?)"
            dbCommand.Parameters.Add("CollectorId", OleDbType.Integer).Value = CollectorId
            dbCommand.Parameters.Add("BacketCounter", OleDbType.UnsignedInt).Value = BacketCounter
            dbCommand.Parameters.Add("WaterGVSCounter", OleDbType.UnsignedInt).Value = WaterGVSCounter
            dbCommand.Parameters.Add("WaterHeatingCounter", OleDbType.UnsignedInt).Value = WaterHeatingCounter
            dbCommand.ExecuteNonQuery()
            dbConn.Close()
            Return True
        Catch exdb As Exception
            If debug Then EventLog1.WriteEntry("Ошибка записи в БД: " + exdb.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            dbConn.Close()
            Return False
        End Try
        ' Закрываем соединение с базой
        dbConn.Close()
        Return False
    End Function
End Class

Friend Class Collector
    ' Класс для сохранения данных счетчика
    Public Property Id As Integer
    Public Property IsActive As Boolean
    Public Property ComPortName As String
    Public Property BaudRate As Integer
    Public Property Address As String
    Public Property Description As String
End Class
