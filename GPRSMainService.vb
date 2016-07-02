Option Explicit On
Imports System.IO
Imports System.Collections.Generic
Imports System.Data
Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.ServiceProcess
Imports Microsoft.Win32

Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Globalization

Public Class GPRSMainService
    Public InsertTotblSampling(100) As Boolean
    Public m_DB As New clsDataBase(My.Settings.ICSDBConnectionString, My.Application.Info.DirectoryPath)
    Dim buffer(100) As String
    Public ApplicationAdd As String = My.Application.Info.DirectoryPath
    Public regKey As RegistryKey
    Public WithEvents tmrTRTWMO As System.Timers.Timer
    Public SendTOSwitchQueue As New List(Of Long)
    Public ISWritetoTcpLog As Boolean = False
    Public flgISFileGeneratedBefor As Boolean = False
    Const RegPath = "Software\Partonegar\AMSS_TO_ICS\Config"
    Dim CA As New clsConfigAccess
    Dim Ack_Source As String
    Dim Appver As String
    Public WithEvents Server As New TCPServer("8093", 1000, 1000)
    Protected Overrides Sub OnStart(ByVal args() As String)
        Dim CI As New CultureInfo("en-US")
        Dim DTI As New DateTimeFormatInfo() With {.DateSeparator = "/", .LongDatePattern = "yyyy/MMMM/dd", .ShortDatePattern = "yyyy/MM/dd", .LongTimePattern = "HH:mm:ss", .ShortTimePattern = "HH:mm"}

        CI.DateTimeFormat = DTI
        System.Windows.Forms.Application.CurrentCulture = CI
        Threading.Thread.CurrentThread.CurrentCulture = CI

        Appver = "Version:95.0402.00"
        WriteVersion(Appver)
        WriteSocket("================================================================================", 0)
        WriteSocket("                                 GPRS_Main is Start", 0)
        WriteSocket("================================================================================", 0)
        tmrTRTWMO = New System.Timers.Timer

        Try
            tmrTRTWMO.Interval = 60000 '1000 * GetPollingDuration()
            tmrTRTWMO.Enabled = True

        Catch ex As Exception
        End Try


        '====================================MR Forghani======================================================
    End Sub
    Protected Overrides Sub OnStop()
        WriteSocket("================================================================================", 0)
        WriteSocket("                                 GPRS_Main is stoped", 0)
        WriteSocket("================================================================================", 0)
        ' Add code here to perform any tear-down necessary to stop your service.
        ' WriteLogs(Format(Now, "yyyy/MM/dd HH:mm:ss") & "Stop", -1)
        ' WritSucetStatus(Format(Now, "yyyy/MM/dd HH:mm:ss") & "Stop")
    End Sub
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub

#Region "Write"
    Public Sub checkDirectory()
        Dim dirPath As String = Format(System.DateTime.Now, "yyyyMMdd")
        Try
            If Not Directory.Exists(ApplicationAdd & "\Logs\" & dirPath) Then
                Directory.CreateDirectory(ApplicationAdd & "\Logs\" & dirPath)
            End If
            Dim dirold As String = Format(System.DateTime.Now.AddDays(-7), "yyyyMMdd")
            If Directory.Exists(ApplicationAdd & "\Logs\" & dirold) Then
                Dim files()
                files = Directory.GetFiles(ApplicationAdd & "\Logs\" & dirold)
                For Each File In files
                    System.IO.File.Delete(File)
                Next
                Directory.Delete(ApplicationAdd & "\Logs\" & dirold)
            End If
        Catch ex As Exception
            '  WriteFileLogError("checkDirectory  " & ex.Message)

        End Try
    End Sub
    Public Sub WriteSocket(ByVal log As String, ByVal ID As Integer) '
        Try
            Dim dirPath As String = Format(System.DateTime.Now, "yyyyMMdd")
            checkDirectory()
            If My.Settings.ISWriteLog Then
                Dim sw As New StreamWriter(ApplicationAdd & "\Logs\" & dirPath & "\Client_" & Format(ID, "000") & ".txt", True)
                sw.WriteLine(System.DateTime.Now.ToString & ":" & log)
                sw.Close()
            End If
        Catch ex As Exception
            'WriteFileLogError("WriteSocket  " & ex.Message)
            'WriteFileLogError("Log is  " & log)
        End Try
    End Sub
    Public Sub WriteVersion(ByVal log As String) '
        Try
            If My.Settings.ISWriteLog Then
                Dim sw As New StreamWriter(ApplicationAdd & "\Logs\ver.txt", True)
                sw.WriteLine(System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & ":" & log)
                sw.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub WritePORT(ByVal log As String, ByVal ClientID As Integer) '

        Dim dirPath As String = Format(System.DateTime.Now, "yyyyMMdd")
        Try
            checkDirectory()



            If Server.WMOCode(ClientID) = "" Then

                Dim sw As New StreamWriter(ApplicationAdd & "\Logs\" & dirPath & "\PortClient" & ClientID & ".txt", True)
                sw.WriteLine(System.DateTime.Now.ToString & ":" & log)
                sw.Close()

            Else
                If File.Exists(ApplicationAdd & "\Logs\" & dirPath & "\PortClient" & ClientID & ".txt") Then
                    Dim FsTempfile As FileStream = New FileStream(ApplicationAdd & "\Logs\" & dirPath & "\PortClient" & ClientID & ".txt", FileMode.Open, FileAccess.Read, FileShare.None)
                    Dim FsCorrectFile As FileStream = New FileStream(ApplicationAdd & "\Logs\" & dirPath & "\Port" & Server.WMOCode(ClientID) & "_" & Format(System.DateTime.Now, "yyyyMMdd") & ".txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
                    Dim rsread As StreamReader = New StreamReader(FsTempfile)
                    Dim rsWrite As StreamWriter = New StreamWriter(FsCorrectFile)
                    rsWrite.WriteLine("*****************************************************************************")
                    rsWrite.Write(rsread.ReadToEnd())
                    rsWrite.WriteLine(System.DateTime.Now.ToString & ":" & log)
                    rsWrite.Flush()
                    rsWrite.Close()
                    FsCorrectFile.Dispose()
                    FsTempfile.Dispose()
                    rsread.Close()
                    rsread.Dispose()
                    File.Delete(ApplicationAdd & "\Logs\" & dirPath & "\PortClient" & ClientID & ".txt")
                Else
                    Dim sw As New StreamWriter(ApplicationAdd & "\Logs\" & dirPath & "\Port" & Server.WMOCode(ClientID) & "_" & Format(System.DateTime.Now, "yyyyMMdd") & ".txt", True)
                    sw.WriteLine(System.DateTime.Now.ToString & ":" & log)
                    sw.Close()
                End If
            End If




        Catch ex As Exception
            'WriteFileLogError("WritePORT  " & ex.Message)
            'WriteFileLogError("Log is  " & log)
        End Try
    End Sub
    Public Sub WriteLogs(ByVal log As String, ByVal ClienID As String, ByVal wmocode As String) '
        '    WriteFileLogError("WriteLogs create path  ")
        Dim dirPath As String = Format(System.DateTime.Now, "yyyyMMdd")
        checkDirectory()
        '     WriteFileLogError("WriteLogs create path end  ")
        Try


            '   WriteFileLogError("WriteLogs writing start  ")


            Dim sw As New StreamWriter(ApplicationAdd & "\Logs\" & dirPath & "\" & wmocode & "_" & Format(System.DateTime.Now, "yyyyMMdd") & ".txt", True)
            sw.WriteLine(System.DateTime.Now.ToString & ":" & log)
            sw.Close()



        Catch ex As Exception
            'WriteFileLogError("WriteLogs  " & ex.Message)
            'WriteFileLogError("Log is  " & log)
        End Try

    End Sub

    Public Sub WriteCheck(ByVal log As String) '
        Try
            If My.Settings.ISWriteLog Then
                Dim sw As New StreamWriter(ApplicationAdd & "\Logs\test.txt", True)
                sw.WriteLine(System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & ":" & log)
                sw.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region


    Private Sub tmrTRTWMO_Tick(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles tmrTRTWMO.Elapsed

    End Sub
    Public Sub delay(ByVal ms As Long)
        Dim t As Long = Now.Ticks
        While Now.Ticks - t < ms * Math.Pow(10, 4)
            System.Windows.Forms.Application.DoEvents()
        End While
    End Sub
    Public Enum LAMB
        STX = 2
        ACK = 6
        CR = 13
        ETB = 23
        esc = 27
        DLMIT = 124
    End Enum
#Region "Socket"
    Public Sub DataArrival(ByVal ClientID As Integer, ByVal Data As String, ByVal ip As String) Handles Server.DataArival
        Dim temp As String
        Dim HasAlarm As Boolean = False
        Try

            temp = buffer(ClientID)
            WritePORT(Data, ClientID)
            buffer(ClientID) = Replace(temp & Trim(Data), Chr(0), "")
            '[55102][HWTEST][1]CONNECT->
            If buffer(ClientID).Contains("CONNECT") And buffer(ClientID).Contains(vbCrLf) Then


                ''WriteLogs(buffer(ClientID), ClientID, 11111)
                Dim strsplit = Split(buffer(ClientID), "[")
                Dim index As Integer = strsplit(2).IndexOf("]")
                Dim StationName As String = Mid(strsplit(2), 1, index)
                WritePORT("StationName:" & StationName, ClientID)
                Dim VendorName As String = ""
                Try
                    VendorName = m_DB.usp_GetStation_LoggerType(StationName)
                Catch ex As Exception
                    WritePORT("VendorName error:" & ex.Message, ClientID)
                End Try

                WritePORT("VendorName:" & VendorName, ClientID)

                Dim ObjectID As Integer = GetStation_IndexByStationName(StationName)
                WritePORT("ObjectID:" & ObjectID, ClientID)
                ''WriteLogs("ObjectID:" & ObjectID, ClientID, 11111)
                If ObjectID = -1 Then

                    WritePORT("This Station connect first time", ClientID)
                    ObjectID = LastObject_ID
                    LastObject_ID = ObjectID + 1
                    LoggerType(ObjectID) = New clsLogger(ClientID, VendorName, StationName, ObjectID)
                    AddHandler LoggerType(ObjectID).RTU_Info_Receiver, AddressOf RTU_Info_Receiver
                    AddHandler LoggerType(ObjectID).RTU_Synchronized, AddressOf RTU_Synchronized
                    AddHandler LoggerType(ObjectID).WriteToseroalPort, AddressOf WriteToseroalPort
                    AddHandler LoggerType(ObjectID).RTU_Final_File_Generated, AddressOf RTU_Final_File_Generated
                    AddHandler LoggerType(ObjectID).RTU_Alarm, AddressOf RTU_Alarm
                    Dim tblStation As New System.Data.DataTable
                    m_DB.GETStationInfoByName(tblStation, StationName)
                    LoggerType(ObjectID).StationCode = tblStation.Rows(0).Item("WMOCode")
                    WriteLogs(StationName, ClientID, LoggerType(ObjectID).StationCode)
                    LoggerType(ObjectID).StationID = tblStation.Rows(0).Item("IKAOCode")

                    If strsplit.Length > 3 Then
                        LoggerType(ObjectID).IsReadFromRTU = True
                        If buffer(ClientID).Contains("->") Then
                            HasAlarm = False
                        ElseIf buffer(ClientID).Contains(">") Then
                            HasAlarm = True
                        End If

                    Else
                        LoggerType(ObjectID).IsReadFromRTU = False
                    End If
                    LoggerType(ObjectID).HasAlarm = HasAlarm
                    Try
                        Dim index1 As Integer = strsplit(3).IndexOf("]")
                        Dim UnreadLog As String = Mid(strsplit(3), 1, index1)

                        LoggerType(ObjectID).Unread_Logs = CInt(UnreadLog)
                    Catch ex As Exception
                    End Try
                    Try
                        ' LoggerType(ObjectID).vendor = VendorName
                        'Partonegar = 0
                        'Theodor = 1
                        'Thies = 2
                        'Lambresht = 3
                        'Vaisala = 4
                        'Thiestdl16 = 22
                        Select Case UCase(VendorName)
                            Case "PARTONEGAR"
                                LoggerType(ObjectID).vendor = "Partonegar"
                            Case "THEODOR FRIEDRICHS" 'heodor Friedrichs"
                                LoggerType(ObjectID).vendor = "Theodor"
                            Case "THIES", "THIES_CLIMA"
                                LoggerType(ObjectID).vendor = "Thies"
                            Case "VAISALA" 'Vaisala"
                                LoggerType(ObjectID).vendor = "Vaisala"
                            Case "LAMBRECHT", "LAMBERESHT" 'Lambresht
                                LoggerType(ObjectID).vendor = "LAMBRECHT"
                        End Select
                    Catch ex As Exception
                    End Try
                Else
                    If strsplit.Length > 3 Then
                        LoggerType(ObjectID).IsReadFromRTU = True
                        If buffer(ClientID).Contains("->") Then
                            HasAlarm = False
                        ElseIf buffer(ClientID).Contains(">") Then
                            HasAlarm = True
                        End If

                    Else
                        LoggerType(ObjectID).IsReadFromRTU = False

                    End If

                    Try
                        Dim index1 As Integer = strsplit(3).IndexOf("]")
                        Dim UnreadLog As String = Mid(strsplit(3), 1, index1)

                        LoggerType(ObjectID).Unread_Logs = CInt(UnreadLog)
                    Catch ex As Exception
                    End Try
                    LoggerType(ObjectID).ClientID = ClientID
                    Remove(ObjectID, ClientID)
                End If
                buffer(ClientID) = ""
                LoggerType(ObjectID).HasAlarm = HasAlarm
                WritePORT(StationName & " IsReadFromRTU:" & LoggerType(ObjectID).IsReadFromRTU, ClientID)
                WritePORT(StationName & " HasAlarm:" & HasAlarm, ClientID)
                LoggerType(ObjectID).StartGetInfo()
            Else
                Dim ObjectID As Integer = GetStation_IndexByClientID(ClientID)

                If LoggerType(ObjectID).IsReadFromRTU = True Then
                    LoggerType(ObjectID).tmr_Timeout.Enabled = False
                    LoggerType(ObjectID).tmr_Timeout.Enabled = True
                    LoggerType(ObjectID).RecDataFromRTU(ClientID, buffer(ClientID))
                    'ElseIf UCase(LoggerType(ObjectID).vendor).Contains("LAM") Then
                    '    WriteLogs("RecDataFromRTU  for lam", ClientID, LoggerType(ObjectID).StationCode)
                    '    LoggerType(ObjectID).RecDataFromRTU(ClientID, buffer(ClientID))
                ElseIf UCase(LoggerType(ObjectID).vendor).Contains("LAM") Then
                    LoggerType(ObjectID).tmr_Timeout.Enabled = False
                    LoggerType(ObjectID).tmr_Timeout.Enabled = True
                    LoggerType(ObjectID).RecDataForLamberesht(ClientID, buffer(ClientID))
                Else
                        LoggerType(ObjectID).tmr_Timeout.Enabled = False
                        LoggerType(ObjectID).tmr_Timeout.Enabled = True
                        LoggerType(ObjectID).RecDataForPartonegar(ClientID, buffer(ClientID))
                    End If
                    buffer(ClientID) = ""
                End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub WriteToseroalPort(ByVal buffer As String, ByVal ClientID As Integer)

        Server.SendData(buffer, ClientID)
        '''WriteLogs("sending " & buffer, ClientID, Server.WMOCode(ClientID))
    End Sub

    Public Function InsertReceivedFromStation(ByVal strFile As String, ByVal lngStationCode As Long, ByVal strDate As String, ByVal strTime As String, ByVal FK_ReceiveType As Integer) As String
        '6

        Dim Repid As String = lngStationCode & "01" & Format(System.DateTime.Now, "yyyyMMddHHmm")
        InsertReceivedFromStation = m_DB.InsertReceivedFromStationINICDB(strFile, lngStationCode, strDate, strTime, Repid, 1)
    End Function
    Public Function GetInsertLogsFromLogFileToDB(ByVal RATLogID As String, ByVal FullFileName As String, ByVal Manual As Boolean, ByRef StartTime As DateTime, ByRef EndTime As DateTime, ByRef ErrorLog As String) As Long
        '8

        GetInsertLogsFromLogFileToDB = GetInsertLogsFromLogFileToICSDB(RATLogID, FullFileName, Manual, StartTime, EndTime, ErrorLog)

    End Function
    Public Function SpaceCounter(ByVal strText As String) As Integer
        Dim loc1 As Integer = InStr(strText, Chr(32), CompareMethod.Text)
        Dim Len As Integer = strText.Length - 1
        '  Dim Buffer(Len) As Char
        Dim counter As Integer = 0
        '   Buffer(Len) = Convert.ToChar(strText)
        '77321               XTTT
        For i As Integer = loc1 To Len
            If Mid(strText, i, 1) = Chr(32) Then
                counter += 1
            Else : Exit For
            End If
        Next
        SpaceCounter = counter + loc1 - 1
    End Function
    Public Function CheckDateTimeFormat(ByVal _strTime As String, ByVal _strDate As String) As Boolean
        Try


            _strTime = Trim(_strTime)
            Dim Date1 As DateTime = CDate(_strDate & Space(2) & _strTime)
            _strDate = Format(Date1, "yyyy/MM/dd")
            _strTime = Format(Date1, "HH:mm")
            If Date1 > System.DateTime.Now.AddDays(1) Then
                Return False
            Else

                If _strTime.Length < 5 Then Return False
                If _strTime.Length = 5 Then
                    _strTime = _strTime & ":00"
                End If
                If _strTime.Length = 7 Then
                    _strTime = "0" & _strTime
                End If
                If _strTime(2) <> ":" OrElse _strTime(5) <> ":" Then Return False
                If _strTime(0) <> "0" AndAlso _strTime(0) <> "1" AndAlso _strTime(0) <> "2" Then Return False
                If _strTime(1) <> "0" AndAlso _strTime(1) <> "1" AndAlso _strTime(1) <> "2" AndAlso _strTime(1) <> "3" AndAlso _strTime(1) <> "4" AndAlso _strTime(1) <> "5" AndAlso _strTime(1) <> "6" AndAlso _strTime(1) <> "7" AndAlso _strTime(1) <> "8" AndAlso _strTime(1) <> "9" Then Return False
                If _strTime(3) <> "0" AndAlso _strTime(3) <> "1" AndAlso _strTime(3) <> "2" AndAlso _strTime(3) <> "3" AndAlso _strTime(3) <> "4" AndAlso _strTime(3) <> "5" Then Return False
                If _strTime(4) <> "0" AndAlso _strTime(4) <> "1" AndAlso _strTime(4) <> "2" AndAlso _strTime(4) <> "3" AndAlso _strTime(4) <> "4" AndAlso _strTime(4) <> "5" AndAlso _strTime(4) <> "6" AndAlso _strTime(4) <> "7" AndAlso _strTime(4) <> "8" AndAlso _strTime(4) <> "9" Then Return False
                If _strTime(6) <> "0" AndAlso _strTime(6) <> "1" AndAlso _strTime(6) <> "2" AndAlso _strTime(6) <> "3" AndAlso _strTime(6) <> "4" AndAlso _strTime(6) <> "5" Then Return False
                If _strTime(7) <> "0" AndAlso _strTime(7) <> "1" AndAlso _strTime(7) <> "2" AndAlso _strTime(7) <> "3" AndAlso _strTime(7) <> "4" AndAlso _strTime(7) <> "5" AndAlso _strTime(7) <> "6" AndAlso _strTime(7) <> "7" AndAlso _strTime(7) <> "8" AndAlso _strTime(7) <> "9" Then Return False
                If Val(_strTime(0) & _strTime(1)) > 23 Then Return False
                If Val(_strTime(3) & _strTime(4)) > 59 Then Return False
                If Val(_strTime(6) & _strTime(7)) > 59 Then Return False


                '====
                Dim _months() As Integer = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
                _strDate = Trim(_strDate)
                '2013/12/4

                If _strDate.Length < 10 Then Return False
                If _strDate(4) <> "/" OrElse _strDate(7) <> "/" Then Return False
                If _strDate(0) <> "1" AndAlso _strDate(0) <> "2" Then Return False
                If _strDate(1) <> "0" AndAlso _strDate(1) <> "1" AndAlso _strDate(1) <> "2" AndAlso _strDate(1) <> "3" AndAlso _strDate(1) <> "4" AndAlso _strDate(1) <> "5" AndAlso _strDate(1) <> "6" AndAlso _strDate(1) <> "7" AndAlso _strDate(1) <> "8" AndAlso _strDate(1) <> "9" Then Return False
                If _strDate(2) <> "0" AndAlso _strDate(2) <> "1" AndAlso _strDate(2) <> "2" AndAlso _strDate(2) <> "3" AndAlso _strDate(2) <> "4" AndAlso _strDate(2) <> "5" AndAlso _strDate(2) <> "6" AndAlso _strDate(2) <> "7" AndAlso _strDate(2) <> "8" AndAlso _strDate(2) <> "9" Then Return False
                If _strDate(3) <> "0" AndAlso _strDate(3) <> "1" AndAlso _strDate(3) <> "2" AndAlso _strDate(3) <> "3" AndAlso _strDate(3) <> "4" AndAlso _strDate(3) <> "5" AndAlso _strDate(3) <> "6" AndAlso _strDate(3) <> "7" AndAlso _strDate(3) <> "8" AndAlso _strDate(3) <> "9" Then Return False
                If _strDate(5) <> "0" AndAlso _strDate(5) <> "1" Then Return False
                If _strDate(6) <> "0" AndAlso _strDate(6) <> "1" AndAlso _strDate(6) <> "2" AndAlso _strDate(6) <> "3" AndAlso _strDate(6) <> "4" AndAlso _strDate(6) <> "5" AndAlso _strDate(6) <> "6" AndAlso _strDate(6) <> "7" AndAlso _strDate(6) <> "8" AndAlso _strDate(6) <> "9" Then Return False
                If _strDate(8) <> "0" AndAlso _strDate(8) <> "1" AndAlso _strDate(8) <> "2" AndAlso _strDate(8) <> "3" AndAlso _strDate(8) <> Chr(32) Then Return False
                If _strDate(9) <> "0" AndAlso _strDate(9) <> "1" AndAlso _strDate(9) <> "2" AndAlso _strDate(9) <> "3" AndAlso _strDate(9) <> "4" AndAlso _strDate(9) <> "5" AndAlso _strDate(9) <> "6" AndAlso _strDate(9) <> "7" AndAlso _strDate(9) <> "8" AndAlso _strDate(9) <> "9" Then Return False
                Dim _year As Integer = Val(_strDate(0) & _strDate(1) & _strDate(2) & _strDate(3))
                If _year Mod 4 = 0 Then _months(1) = 29
                If _year Mod 100 = 0 Then _months(1) = 28
                If _year Mod 400 = 0 Then _months(1) = 29
                Dim _month As Integer = Val(_strDate(5) & _strDate(6))
                Dim _day As Integer = Val(_strDate(8) & _strDate(9))
                If _year > 2100 OrElse _year < 1900 Then Return False
                If _month < 1 OrElse _month > 12 Then Return False
                If _day < 1 OrElse _day > _months(_month - 1) Then Return False
                Return True

            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function GetInsertLogsFromLogFileToICSDB(ByVal RATLogID As String, ByVal FullFileName As String, ByVal Manual As Boolean, ByRef StartTime As DateTime, ByRef EndTime As DateTime, ByRef ErrorLogs As String) As Long

        ErrorLogs = ErrorLogs + " GetInsertLogsFromLogFileToICSDB Step 1" & vbCrLf
        Dim fs As New System.IO.FileStream(FullFileName, FileMode.Open, FileAccess.Read, FileShare.Read)
        ErrorLogs = ErrorLogs + " GetInsertLogsFromLogFileToICSDB Step 2" & vbCrLf
        Dim rs As New System.IO.StreamReader(fs)
        'Dim strSpil() As String
        Dim strLine As String
        Dim lngProgress As Integer = 0
        Dim strTemp As String
        Dim strTime As String = ""
        Dim strDate As String = ""
        Dim strStationHeadear As String
        Dim i As Long
        Dim SenHeader(100) As Long
        Dim SenMin(100) As Single
        Dim SenMax(100) As Single
        Dim SenValue(100) As Single
        Dim sglMin As Single
        Dim sglMax As Single
        Dim SpaceCount As Integer = 0
        Dim RainID As Long = -1
        Dim HumidityID As Long = -1
        Dim TemperatureID As Long = -1
        Dim hasRain As Boolean = False
        Dim PreTemp As Single
        strTemp = ""
        '=================================================

        Try


            strLine = ""

            strStationHeadear = rs.ReadLine

            Do
                strLine = rs.ReadLine
            Loop Until InStr(strLine.ToUpper, "DATE", CompareMethod.Binary) > 0
            SpaceCount = SpaceCounter(strLine)
            Dim LineLen As Long = strLine.Length
            Dim WMOCODE As String = Val(Trim(Mid(strStationHeadear, 1, SpaceCount)))
            For i = 2 To LineLen \ SpaceCount - 1
                strTemp = Trim(Mid(strLine, i * SpaceCount + 1, SpaceCount))


                SenHeader(i - 2) = m_DB.GetSenTypeIDByStation_AND_SenNameINICSDB(WMOCODE, strTemp, sglMin, sglMax)
                ErrorLogs = ErrorLogs + " SensorName:" & strTemp & " SenHeader(i - 2):" & SenHeader(i - 2) & vbCrLf
                '   ErrorLogs = ErrorLogs + " SensorName: " & strTemp & vbCrLf
                If SenHeader(i - 2) = 0 Then
                    GetInsertLogsFromLogFileToICSDB = 1
                    '    Exit Function

                End If
                If UCase(strTemp).Contains("RAIN") Or UCase(strTemp).Contains("PRECIPITATION") Then
                    RainID = SenHeader(i - 2)
                End If
                If UCase(strTemp).Contains("RAIN") Or UCase(strTemp).Contains("PRECIPITATION") Then
                    RainID = SenHeader(i - 2)
                End If
                If UCase(strTemp) = "TEMPERATURE" Or UCase(strTemp) = "AIR TEMPERATURE" Then
                    TemperatureID = SenHeader(i - 2)

                    PreTemp = m_DB.usp_tblJoinSensorType_LoggerModule_GetValue(WMOCODE, strTemp)
                End If
                If UCase(strTemp).Contains("HUMIDITY") Or UCase(strTemp).Contains("REL. HUMIDITY") Then
                    HumidityID = SenHeader(i - 2)
                End If
                '     SenMin(i - 2) = sglMin
                '  SenMax(i - 2) = sglMax
            Next i
            '  WriteDCSLog("TempID " & TemperatureID.ToString)
            '    WriteDCSLog("HumidityID " & HumidityID.ToString)
            Dim firstLine As Boolean = True
            ErrorLogs = ErrorLogs + " GetInsertLogsFromLogFileToICSDB Step 3" & vbCrLf
            While (Not rs.EndOfStream)
                strLine = ""
                strLine = rs.ReadLine
                Try
                    strDate = Trim(Mid(strLine, 1, SpaceCount))
                    strTime = Trim(Mid(strLine, SpaceCount + 1, SpaceCount))
                    If firstLine Then
                        Dim date1 As Date = CDate(strDate & "  " & strTime)
                        StartTime = date1
                        firstLine = False
                    End If

                    '  Dim Time1 As Date = Cd(strTime)
                    Dim datetime As New DateTime
                    strTime.Replace("24", "00")
                    Dim SplitTime = Split(strTime, ":")
                    If SplitTime(0).Length = 1 Then
                        strTime = "0" & strTime
                    End If
                    Dim tblSampling As New System.Data.DataTable
                    tblSampling.Columns.Add("SampleDateTime")
                    tblSampling.Columns.Add("SensorValue")
                    tblSampling.Columns.Add("QualityControlLevel")
                    tblSampling.Columns.Add("FK_JoinSensorType_LoggerModuleID")
                    tblSampling.Columns.Add("FK_SendAndReceiveLogID")
                    tblSampling.Columns.Add("WMOCODE")
                    If CheckDateTimeFormat(strTime, strDate) Then
                        '  datetime = Convert.ToDateTime(strDate & Space(2) & strTime)

                        For i = 2 To LineLen \ SpaceCount - 1
                            Try
                                strTemp = Trim(Mid(strLine, i * SpaceCount + 1, SpaceCount))
                                Try
                                    tblSampling.Clear()
                                Catch ex As Exception

                                End Try
                                If SenHeader(i - 2) = RainID And strTemp > 0 Then
                                    hasRain = True
                                End If

                                tblSampling.Rows.Add()
                                tblSampling.Rows(0).Item("SampleDateTime") = CDate(strDate & " " & strTime)
                                tblSampling.Rows(0).Item("SensorValue") = strTemp
                                tblSampling.Rows(0).Item("FK_JoinSensorType_LoggerModuleID") = SenHeader(i - 2)
                                tblSampling.Rows(0).Item("FK_SendAndReceiveLogID") = RATLogID
                                tblSampling.Rows(0).Item("WMOCODE") = WMOCODE
                                If SenHeader(i - 2) = HumidityID And strTemp = 0 Then
                                    '    WriteDCSLog("humidity is 0 ")
                                    m_DB.InsetIntotblSamplingwithQuality(tblSampling, 1)
                                ElseIf SenHeader(i - 2) = TemperatureID Then
                                    '  WriteDCSLog("Temp is 0 ")
                                    If strTemp = 0 And Math.Abs(PreTemp - strTemp) > 10 Then

                                        m_DB.InsetIntotblSamplingwithQuality(tblSampling, 1)
                                    Else
                                        PreTemp = strTemp
                                        m_DB.InsetIntotblSampling1(tblSampling)
                                    End If
                                Else
                                    m_DB.InsetIntotblSampling1(tblSampling)
                                End If
                            Catch ex As Exception

                            End Try

                        Next i

                    Else

                        ' WriteErrorLog(FullFileName & "   Have wrong Date and Time" & vbCrLf)
                        ' WriteErrorLog(strDate & "   " & strTime)
                    End If
                Catch ex As Exception
                    WriteCheck("step2* :" & ex.Message)
                End Try
                Try
                    If Manual Then
                        '  frmDataEntryptr.ProgressBar1.Value = lngProgress
                        lngProgress = lngProgress + 1
                    End If
                Catch ex As Exception
                    ''''WriteCheck("step3* :" & ex.Message)
                End Try
            End While
            Try
                EndTime = strDate & " " & strTime
                If EndTime > Format(System.DateTime.Now.AddMinutes(10), "yyyy/MM/dd") Then
                    For i = 2 To LineLen \ SpaceCount - 1
                        Try
                            strTemp = Trim(Mid(strLine, i * SpaceCount + 1, SpaceCount))

                            m_DB.UpDatetblJoinSensorType_LoggerModuleValue(WMOCODE, strTemp, SenHeader(i - 2), strTime, strDate)
                        Catch ex As Exception
                            ''''WriteCheck("step4* :" & ex.Message)

                        End Try


                    Next i
                End If



            Catch ex As Exception
                ''''WriteCheck("step5* :" & ex.Message)

            End Try
            ErrorLogs = ErrorLogs + " GetInsertLogsFromLogFileToICSDB Step 4" & vbCrLf
            Call m_DB.UpDatetblSendandRecieve(RATLogID)
            Call m_DB.UpDatetblSendandRecieveHasRain(RATLogID, hasRain)
            ErrorLogs = ErrorLogs + " GetInsertLogsFromLogFileToICSDB Step 5" & vbCrLf
            GetInsertLogsFromLogFileToICSDB = 0
        Catch ex As Exception
            GetInsertLogsFromLogFileToICSDB = 1
        Finally
            rs.Close()
            fs.Close()

        End Try
    End Function

    Public Function ReadLogs(ByVal RealFile As String, ByVal StationCode As String, ByRef errorlog As String)
        Dim fs As New System.IO.FileStream(RealFile, FileMode.Open, FileAccess.Read, FileShare.Read)
        Dim rs As New System.IO.StreamReader(fs)
        Dim strda As String = ""
        strda = rs.ReadToEnd
        Dim strstil() = Split(strda, vbCrLf, -1, CompareMethod.Binary)
        If strstil(2) = "" Then

        Else

            Dim SysDateNow As String = Format(System.DateTime.Now, "yyyy/MM/dd")
        Dim SysTimeNow As String = Microsoft.VisualBasic.Format(System.DateTime.Now, "HH:mm")
        Dim lngRATLogID As String
            Try

                Try
                    lngRATLogID = InsertReceivedFromStation(RealFile, StationCode, SysDateNow, SysTimeNow, 3)
                Catch ex As Exception
                    WriteCheck("step101*:" & ex.Message)
                End Try
                '   WriteDCSLog("RatID is " & lngRATLogID)
                Dim Stime, ETime As DateTime
                If lngRATLogID >= 0 Then
                    '  Dim errorlog As String = ""
                    GetInsertLogsFromLogFileToDB(lngRATLogID, RealFile, True, Stime, ETime, errorlog)
                Else
                End If
            Catch ex As Exception
                WriteCheck("step000*:" & ex.Message)
            End Try

        End If
    End Function
    Private Sub RTU_Final_File_Generated(ByVal ObjectID As Integer, ByVal RealFileName As String, ByVal stationCode As Long) ' Handles RTUPort.RTU_Final_File_Generated
        Dim errorlog As String = ""
        WriteLogs("ReadLogs start", ObjectID, stationCode)
        ReadLogs(RealFileName, stationCode, errorlog)
        WriteLogs("errorlog:" & errorlog, ObjectID, stationCode)
        WriteLogs("ReadLogs end", ObjectID, stationCode)
    End Sub

    Private Sub RTU_Synchronized(ByVal objectID As Integer, ByVal vendor As String, ByVal MemoryActive As Boolean)
        Try



            'LoggerType(objectID).StationCode = m_DB.GetStationCodeByName(LoggerType(objectID).StationName)



            'Try
            '    LoggerType(objectID).DataFolder = (GetReceiveFolder() & "\" & LoggerType(objectID).StationCode & "\").Replace("\\", "\")
            '    If Not Directory.Exists(LoggerType(objectID).DataFolder) Then
            '        Directory.CreateDirectory(LoggerType(objectID).DataFolder)
            '    End If
            '    LoggerType(objectID).TempFile = (GetReceiveFolder() & "\" & LoggerType(objectID).StationCode & "\" & LoggerType(objectID).StationCode & "temp_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt").Replace("\\", "\")

            '    LoggerType(objectID).LoggerType = "COMBILOG"
            '    Dim lngErr As Long = 0
            '    LoggerType(objectID).StationID = m_DB.GetStationIDByCode(LoggerType(objectID).StationCode, lngErr)

            '    WriteLogs(LoggerType(objectID).TempFile.ToString, 1, LoggerType(objectID).StationID)
            'Catch ex As Exception

            'End Try
            LoggerType(objectID).GetLogs(MemoryActive)
            'Else

            ''WriteDCSLog("There is not any logs in RTU memory" & vbCrLf)

            ''RTUPort(objectID).Hangup()
            'End If

        Catch ex As Exception
            '   WriteRTUError("RTU_Synchronized Error is  " & " : " & ex.Message)
        End Try
    End Sub

    Public Sub RTU_Alarm(ByVal ObjectID As Integer, clientid As Integer, Alarms As clsLogger.ALARM()) ' Handles RTUPort.RTU_Info_Receiver
        Try
            m_DB.usp_tblAlarms_Insert(Alarms(0), LoggerType(ObjectID).StationCode)
            WriteLogs("Alarm Rec", clientid, LoggerType(ObjectID).StationID)

        Catch ex As Exception

        End Try

    End Sub
    Public ALARMS() As ALARM
    Structure ALARM
        Dim A_Name As String
        Dim A_DateTime As String
        Dim A_Date As String
        Dim A_Time As String
        Dim A_Value As Single
        Dim A_Interval As String
    End Structure
    Public Sub RTU_Info_Receiver(ByVal ObjectID As Long, ByVal LoggerName As String, ByVal Logger_SN As String, ByVal SenNo As Byte, ByVal Paket_Size As Long, ByVal Memory_Active As Boolean, ByVal Unread_Logs As Long, ByVal Vendor As String, ByVal SW As String, ByVal HW As String, ByVal ModulType As String, ByVal MemoryActive As Boolean) ' Handles RTUPort.RTU_Info_Receiver
        Try


            'WriteDCSLog("Unread logs :" & Unread_Logs & vbCrLf)
            'WriteDCSLog("==========================" & vbCrLf)

            '  If GetSnycron() Then
            '

            '     WriteDCSLog(RTUPort(ObjectID).StationName & " : start syncronizing")
            LoggerType(ObjectID).RTU_Synchronizing()

            'Else

            '    '  WriteDCSLog(RTUPort(ObjectID).StationName & " : not set to syncronizing")


            '        'testIf Vendor = "Thiestdl16" Or ((Not MemoryActive) And Vendor = "Thies") Then
            '        If Vendor = "Thiestdl16" Or (Vendor = "Thies") Then
            '            'htlocation
            '            InsertTOtblCallInfo(4, RTUPort(ObjectID).ContactID, CallType)
            '            InsertTOtblCallInfo(5, RTUPort(ObjectID).ContactID, CallType)
            '            updateReadDataPercentIntblCallInfo(RTUPort(ObjectID).ContactID, -1)
            '            RTUPort(ObjectID).GetLogs()
            '        ElseIf Unread_Logs > 0 Then
            '            '   pnlProgress.Height = 50

            '            RTUPort(ObjectID).GetLogs()
            '        Else
            '        ' WriteDCSLog("There is not any logs in RTU memory" & vbCrLf)
            '            '      RTUPort(ObjectID).Goodbye()
            '            '  WriteRTULog("Hang 18")
            '            RTUPort(ObjectID).Hangup()
            '        End If
            '    End If



        Catch ex As Exception

        End Try

    End Sub
    Public Sub Connected(ByVal ClientID As Integer, ByVal ip As String) Handles Server.Connected
        Try
            ' WriteDaily()
            WriteSocket("Client By ID " & ClientID & "  Is Connected" & vbCrLf, ClientID)
            WriteSocket("ip is  " & ip & vbCrLf, ClientID)
            WriteSocket("Client Port is " & ip & vbCrLf, ClientID)
            InsertTotblSampling(ClientID) = False
        Catch ex As Exception
            WriteSocket("Connected" & ex.Message, ClientID)
        End Try
    End Sub
    Public Sub Disconnected(ByVal ClientID As Integer) Handles Server.Disconnected
        Try
            WriteSocket("Client By id  " & ClientID & "  Is disConnected", ClientID)
            'If Server.WMOCode(ClientID) <> "" Then
            '    WriteSocket("Client By ID " & Server.WMOCode(ClientID) & "  Is disConnected" & vbCrLf)
            'End If

        Catch ex As Exception
            WriteSocket("Disconnected" & ex.Message, ClientID)
        End Try
    End Sub
    Public Sub SocketError(ByVal ClientID As Integer, ByVal ER As String, ByVal oNDicconnect As Boolean) Handles Server.SocketError
        'Try
        '    WriteSocket("ClientID " & ClientID)
        'Catch ex As Exception
        '    WriteSocket(" write  SocketError ClientID" & ex.Message)
        'End Try
        Try

            WriteSocket(ER, ClientID)
            If oNDicconnect Then
                WriteSocket("Dicconect in Socket errror", ClientID)
                Server.Disconnect(ClientID)

            End If

        Catch ex As Exception
            WriteSocket(" write  SocketError 2" & ex.Message, ClientID)
            '    WriteSocket(ex.Message)
        End Try
        '  Server.Disconnect(ClientID)
    End Sub

#End Region
End Class

