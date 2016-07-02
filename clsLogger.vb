Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Win32

Public Class clsLogger
    Public Event RTU_Info_Receiver(ByVal ObjectID As Integer, ByVal LoggerName As String, ByVal Logger_SN As String, ByVal SenNo As Byte, ByVal Paket_Size As Long, ByVal Memory_Status As Boolean, ByVal Unread_Logs As Long, ByVal Vendor As String, ByVal SW As String, ByVal HW As String, ByVal ModulType As String, ByVal MemoryActive As Boolean)
    Public Event RTU_Synchronized(ByVal ObjectID As Integer, ByVal Vendor As String, ByVal MemoryActive As Boolean)
    Public Event RTU_Data_Block_Receive(ByVal ObjectID As Integer, ByVal ContactID As Long, ByVal BlockNo As Long, ByVal Percent As Byte, ByVal Vendor As String, ByVal MemoryActive As Boolean)
    Public Event RTU_Final_File_Generated(ByVal ObjectID As Integer, ByVal RealFile As String, ByVal StationCode As Long)
    Public Event WriteToseroalPort(ByVal buffer As String, ByVal ClientID As Integer)
    Public WithEvents tmr_Timeout As New System.Timers.Timer
    Public Event RTU_Alarm(ByVal ObjectID As Integer, clientid As Integer, Alarms As ALARM())
    Dim ApplicationAdd As String = My.Application.Info.DirectoryPath
    Public Enum LAMB
        STX = 2
        ACK = 6
        CR = 13
        ETB = 23
        esc = 27
        DLMIT = 124
    End Enum
    Public ALARMS() As ALARM
    Structure ALARM
        Dim A_Name As String
        Dim A_DateTime As String
        Dim A_Date As String
        Dim A_Time As String
        Dim A_Value As Single
        Dim A_Interval As String
    End Structure
    Public Enum Comunication_States
        PREPARING_PORT = 0
        PORT_READY = 1
        CONNECTING = 2
        CONNECTED = 3
        WAIT_FOR_RTU_NAME = 4
        WAIT_FOR_LOGGER_INFO = 5
        WAIT_FOR_PACKET_SIZE = 6
        WAIT_FOR_UNREAD_LOGS = 7
        WAIT_FOR_GLS = 8
        WAIT_FOR_ONLINE_RESPONSE = 9
        WAIT_FOR_ONLINE_REPEAT = 10
        WAIT_FOR_OFFLINE_RESPONSE = 11
        WAIT_FOR_DATABLOCK = 12
        WAIT_FOR_CMT_ACK = 13
        DISCONNECTING = 14
        SWITCH_TO_COMMAND_MODE = 15
        DISCONNECTED_AND_IDEL = 16
        WAIT_FOR_GOODBYE = 17
        WAIT_FOR_ALARM_CONTENT = 18
        WAIT_FOR_GSL = 19
        WAIT_FOR_ATA = 20
        WAIT_FOR_01E_RESPONSE = 21

        WAIT_FOR_01H_RESPONSE = 22
        WAIT_FOR_01G = 23
        WAIT_FOR_01VRESPONCE = 24
        WaitForReset = 25
        WaitForGetLoggerType = 26
        WaitforVer = 27
        AskforSetVisalaDate = 28
        AskforSetVisalaTime = 29
        AskforSetDateTime = 30
        SetVisalaDate = 31
        SetVisalaTime = 32
        SetDateAndTime = 33
        '    ReadingRTUTime = 36
        WaitforCFG = 34
        ResetRTU = 35
        ReadingRTUTime = 36
        ReadingRTUDate = 37
        Waitforstationname = 38
        Wait01E = 39
        Waitforlogsfromthies = 40
        WaitFor01ZResponse = 41
        WaitFor01ZOKResponse = 42
        WaitforSensorList = 43

        WaitForLambereshtVer = 44
        ReadLambreshtDateTime = 45
        SetLambereshtDateTime = 46
        ReadLogFromLambresht = 47
        GetLambereshtLogCount = 48
        WaitforSULD = 49
        WaitforAnswerSULD = 50
    End Enum
    Dim m_RTUState As Comunication_States
    Dim m_Lamb As LAMB
    Dim str_Buffer As String
    Structure Sensor
        Dim Name As String
        Dim Value As Single
        '=h
        Dim Unit As String
        Dim Code As Long
        Dim Min As Long
        Dim Max As Long
        '======
    End Structure
    Dim srLogs As IO.StreamWriter
    Dim m_Vendor As Logger_Vendor
    Dim m_StationName As String
    Dim m_ClientID As Integer
    Dim m_ObjectID As Integer
    Dim m_SensorCount As Integer
    Dim m_Memory_Active As Boolean
    Dim m_Unread_Logs As Integer
    Dim m_CurentLog As Integer
    Dim m_HWRevision As String
    Dim m_SWRevision As String
    Dim m_ModulType As String
    Dim m_Logger_Name As String
    Dim m_Logger_SN As String
    Dim m_Logger_SenNo As String
    Dim m_LoggerType As String

    Dim m_Temp_File As String
    Dim m_Real_File As String
    Dim m_Bloack_Counter As Long = 0
    Dim m_Data_Persent As Byte = 0
    Dim m_Percent As Integer
    Public m_TempfileCreat As Boolean = False
    Dim SensorIndex As Integer = 0
    Public Sensors(30) As Sensor
    Dim m_datafolder As String
    Dim m_StationCode As Long
    Dim m_StationID As String
    Dim m_strLog As String
    Public m_ISReadFromRtu As Boolean = False


    Property StationID() As String
        Get
            Return m_StationID
        End Get
        Set(ByVal value As String)
            m_StationID = value
        End Set
    End Property
    Dim m_HasAlarm As Boolean
    Property HasAlarm() As Boolean
        Get
            Return m_HasAlarm
        End Get
        Set(ByVal value As Boolean)
            m_HasAlarm = value
        End Set
    End Property
    Public Enum Logger_Vendor
        Partonegar = 0
        Theodor = 1
        Thies = 2
        Lambresht = 3
        Vaisala = 4
        Thiestdl16 = 22
    End Enum
    Property vendor As String
        Get
            Return m_Vendor.ToString
        End Get
        Set(ByVal value As String)
            '   m_Vendor = CType(vendor, Logger_Vendor)

            Select Case UCase(value)
                Case "PARTONEGAR"
                    m_Vendor = 0
                Case "THEODOR FRIEDRICHS" 'heodor Friedrichs"
                    m_Vendor = 1
                Case "THIES"
                    m_Vendor = 2
                Case "VAISALA" 'Vaisala"
                    m_Vendor = 4
                Case "LAMBRECHT"
                    m_Vendor = 3
            End Select
        End Set
    End Property
    Property StationCode() As Long
        Get
            Return m_StationCode
        End Get
        Set(ByVal value As Long)
            m_StationCode = value
        End Set
    End Property
    Property TempFile() As String
        Get
            Return m_Temp_File
        End Get
        Set(ByVal value As String)
            m_Temp_File = value
        End Set
    End Property
    Property StationName() As String
        Get
            Return m_StationName
        End Get
        Set(ByVal value As String)
            m_StationName = value
        End Set
    End Property
    Property RealFile() As String
        Get
            Return m_Real_File
        End Get
        Set(ByVal value As String)
            m_Real_File = value
        End Set
    End Property
    Property DataFolder() As String
        Get
            Return m_datafolder
        End Get
        Set(ByVal value As String)

            m_datafolder = value
        End Set
    End Property

    Public ReadOnly Property ObjectID
        Get
            Return m_ObjectID
        End Get
    End Property
    Public Property ClientID
        Set(ByVal value)
            m_ClientID = value
        End Set
        Get
            Return m_ClientID
        End Get
    End Property
    Public Property Unread_Logs
        Get
            Return m_Unread_Logs
        End Get
        Set(value)
            m_Unread_Logs = value
        End Set
    End Property
    Property LoggerType() As String
        Get
            Return m_LoggerType
        End Get
        Set(ByVal value As String)
            m_LoggerType = value
        End Set
    End Property
    Property IsReadFromRTU() As String
        Get
            Return m_ISReadFromRtu
        End Get
        Set(ByVal value As String)
            m_ISReadFromRtu = value
        End Set
    End Property
    Dim StrBuffer As String = ""
    Public Sub New(ByVal ClientID As Integer, ByVal LoggerType As String, ByVal StationName As String, ByVal ObjectID As Integer)
        m_ObjectID = ObjectID
        Select Case LoggerType
            Case "Partonegar"
                m_Vendor = Logger_Vendor.Partonegar
            Case "Lambresht"
                m_Vendor = Logger_Vendor.Lambresht
            Case "Theodor", "Friedrichs"
                m_Vendor = Logger_Vendor.Theodor
            Case "Thies"

                m_Vendor = Logger_Vendor.Thies
            Case Else

        End Select
        If m_Vendor.ToString = "Partonegar" Then
            m_Memory_Active = False
        End If


        m_StationName = StationName

        m_ClientID = ClientID
        '     m_ObjectID = LastObject_ID


    End Sub
    Public Sub StartGetInfo()
        WriteLogs(m_StationName.ToString & "     StartGetInfo Vendor:" & m_Vendor.ToString & " IsReadFromRTU:" & IsReadFromRTU & " m_HasAlarm:" & m_HasAlarm, ClientID)
        '
        If IsReadFromRTU And (UCase(m_Vendor.ToString).Contains("THEODOR") Or UCase(m_Vendor.ToString).Contains("PARTONEGAR")) And m_HasAlarm = False Then '1
            WriteLogs("StartGetInfo 1", ClientID)
            m_RTUState = Comunication_States.WAIT_FOR_LOGGER_INFO
            RaiseEvent WriteToseroalPort("@$01S" & Chr(13), ClientID)

        ElseIf IsReadFromRTU And (UCase(m_Vendor.ToString).Contains("THEODOR") Or UCase(m_Vendor.ToString).Contains("PARTONEGAR")) And m_HasAlarm = True Then '2

            WriteLogs("StartGetInfo 2", ClientID)
            GetAlarmContent()
        ElseIf IsReadFromRTU And UCase(m_Vendor.ToString).Contains("THIES") Or UCase(m_Vendor.ToString).Contains("VAISALA") Then '3
            '
            WriteLogs("StartGetInfo 3", ClientID)

            RaiseEvent WriteToseroalPort("$GLS?" & vbCrLf, m_ClientID)
            m_RTUState = Comunication_States.WAIT_FOR_GLS


        ElseIf IsReadFromRTU And UCase(m_Vendor.ToString).Contains("LAM") Then
            WriteLogs("StartGetInfo 4", ClientID)
            m_RTUState = Comunication_States.WaitForLambereshtVer  '1l
            RaiseEvent WriteToseroalPort("$GSL?" & vbCrLf, m_ClientID)

        ElseIf UCase(m_Vendor.ToString).Contains("LAM") Then
            WriteLogs("StartGetInfo 5", ClientID)
            m_RTUState = Comunication_States.WaitForLambereshtVer  '1l
            RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "S" & Chr(LAMB.CR), ClientID)



        Else
            WriteLogs("StartGetInfo 6", ClientID)
            ''WriteCheck("step 2: shoro khondan moshakhasate model digar")
            m_RTUState = Comunication_States.WAIT_FOR_LOGGER_INFO
            RaiseEvent WriteToseroalPort("$01S" & Chr(13), ClientID)

        End If

    End Sub
    Public Sub WriteCheck(ByVal log As String) '
        Try
            If My.Settings.ISWriteLog Then
                Dim sw As New StreamWriter(ApplicationAdd & "\Logs\check.txt", True)
                sw.WriteLine(System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & ":" & log)
                sw.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub WriteLogs(ByVal log As String, ByVal ClienID As Long) '
        '    WriteFileLogError("WriteLogs create path  ")
        Dim dirPath As String = Format(System.DateTime.Now, "yyyyMMdd")

        '     WriteFileLogError("WriteLogs create path end  ")
        Try

            Dim sw As New StreamWriter(ApplicationAdd & "\Logs\" & dirPath & "\" & m_StationCode & "_" & Format(System.DateTime.Now, "yyyyMMdd") & ".txt", True)
            sw.WriteLine(System.DateTime.Now.ToString & ":" & log)
            sw.Close()

        Catch ex As Exception
            'WriteFileLogError("WriteLogs  " & ex.Message)
            'WriteFileLogError("Log is  " & log)
        End Try

    End Sub
#Region "Partonegar"
    Public Sub RecDataFromRTU(ByVal clientID As Integer, ByVal Buff As String)

        StrBuffer = StrBuffer + Buff

        Buff = ""
        Try

            If m_RTUState = Comunication_States.WAIT_FOR_LOGGER_INFO And StrBuffer.Contains(vbCr) And StrBuffer.Contains("=") Then
                tmr_Timeout.Enabled = False
                StrBuffer = StrBuffer.Replace(vbCr, "")
                m_Logger_Name = Trim(Mid(StrBuffer, 2, 20))
                m_Logger_SN = Mid(Trim(Right(StrBuffer, 8)), 1, Len(Trim(Right(StrBuffer, 8))) - 2)
                m_Logger_SenNo = ConvHexToDec(Trim(Right(Trim(Right(StrBuffer, 8)), 2)))
                m_SensorCount = m_Logger_SenNo
                '  str_Port_BuffferHistory=str_Port_Bufffer '17
                SensorIndex = 0
                StrBuffer = ""
                Array.Clear(Sensors, 0, Sensors.Length - 1)
                WriteLogs("RTU_SenNo:" & m_Logger_SenNo, clientID)
                ' ''RaiseEvent WriteToseroalPort("$GSL?" & vbCrLf, m_ClientID)
                ' ''m_RTUState = Comunication_States.WAIT_FOR_GSL
                RaiseEvent WriteToseroalPort("$GLS?" & vbCrLf, m_ClientID)
                m_RTUState = Comunication_States.WAIT_FOR_GLS


            ElseIf m_RTUState = Comunication_States.WAIT_FOR_GLS And StrBuffer.Contains(vbCrLf) Then
                '            'Dial = DIAL_RESULTS.CONNECT                    
                tmr_Timeout.Enabled = False

                '   tmr_Timeout.Enabled = True
                If Val(Trim(StrBuffer)) = 0 Then
                    m_Memory_Active = False
                Else
                    m_Memory_Active = True
                End If
                StrBuffer = ""
                WriteLogs("m_Memory_Active:" & m_Memory_Active.ToString, clientID)

                Dim strSensorss As String = GetConfigItem(m_StationCode & "AliasName")

                    If strSensorss <> "" Then
                    Try

                        Array.Clear(Sensors, 0, Sensors.Length - 1)

                        Dim AliasName() = Split(strSensorss, ";")
                        m_SensorCount = AliasName.Length
                        m_Logger_SenNo = m_SensorCount
                        For i As Integer = 0 To AliasName.Count - 1
                            Sensors(i).Name = AliasName(i)
                        Next

                        If UCase(m_Vendor.ToString).Contains("THIES") Or UCase(m_Vendor.ToString).Contains("VAISALA") Then  '"Lamberesht" Then
                            GetLogs(m_Memory_Active)
                        Else
                            m_RTUState = Comunication_States.WAIT_FOR_01VRESPONCE
                            RaiseEvent WriteToseroalPort("@$01V" & vbCr, clientID)
                        End If
                    Catch ex As Exception
                        WriteLogs("error in regestry():" & ex.Message, m_ClientID)
                    End Try
                Else

                    m_RTUState = Comunication_States.WAIT_FOR_GSL
                    RaiseEvent WriteToseroalPort("$GSL?" & vbCrLf, m_ClientID)
                End If

            ElseIf m_RTUState = Comunication_States.WAIT_FOR_GSL And StrBuffer.Contains("OK" & vbCrLf) Then
                WriteLogs("WAIT_FOR_GSL REC OK", clientID)
                tmr_Timeout.Enabled = False
                Try

                    StrBuffer = StrBuffer.Replace(vbCrLf & "OK" & vbCrLf, "")
                    Dim strSpil() = Split(StrBuffer, vbCrLf, -1, CompareMethod.Binary)
                    StrBuffer = ""

                    Array.Clear(Sensors, 0, Sensors.Length - 1)

                    m_SensorCount = strSpil.Length
                    m_Logger_SenNo = m_SensorCount

                    ' End If


                    For i = 0 To m_SensorCount - 1

                        Sensors(i).Name = strSpil(i)
                        Sensors(i).Name = Sensors(i).Name.Replace(Chr(13), "")
                    Next

                Catch ex As Exception
                End Try
                WriteLogs("LoggerType:" & m_Vendor.ToString, clientID)
                If UCase(m_Vendor.ToString).Contains("THIES") Or UCase(m_Vendor.ToString).Contains("VAISALA") Then  '"Lamberesht" Then
                    GetLogs(m_Memory_Active)
                Else
                    m_RTUState = Comunication_States.WAIT_FOR_01VRESPONCE
                    RaiseEvent WriteToseroalPort("@$01V" & vbCr, clientID)
                End If

            ElseIf m_RTUState = Comunication_States.WaitForLambereshtVer And StrBuffer.Contains(vbCr) Then
                tmr_Timeout.Enabled = False
                Try
                    If StrBuffer.Contains("AT+CPOWD=1") Then
                        StrBuffer = ""
                    Else
                        m_Temp_File = (GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode & "temp_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt").Replace("\\", "\")
                        If Not Directory.Exists(GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode) Then
                            Directory.CreateDirectory(GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode)
                        End If
                        m_Real_File = m_Temp_File.Replace("temp", "")

                        srLogs = New IO.StreamWriter(m_Real_File)
                        Dim strLog As String = ""
                        strLog = m_StationCode.ToString & Space(30 - m_StationCode.ToString.Length) & m_StationID & Space(30 - m_StationID.Length) & m_Vendor.ToString & Space(30 - m_Vendor.ToString.Length) & vbCrLf
                        strLog = strLog & "DATE" & Space(30 - "DATE".Length) & "Time" & Space(30 - "Time".Length)
                        StrBuffer = StrBuffer.Replace(vbCrLf & "OK" & vbCrLf, "")
                        Dim strSpil1() = Split(StrBuffer, vbCrLf, -1, CompareMethod.Binary)
                        StrBuffer = ""
                        For i = 0 To strSpil1.Length - 1
                            strLog = strLog & strSpil1(i) & Space(30 - strSpil1(i).Length)
                        Next
                        srLogs.WriteLine(strLog)
                        StrBuffer = ""

                        m_RTUState = Comunication_States.WaitforSULD
                        RaiseEvent WriteToseroalPort("$SULD=" & vbCrLf, clientID)

                    End If
                Catch ex As Exception
                    WriteLogs("step10* : error:" & ex.Message, m_ClientID)
                    StrBuffer = ""
                End Try
            ElseIf m_RTUState = Comunication_States.WaitforSULD And StrBuffer.Contains(">") Then
                Try
                    StrBuffer = ""
                    m_RTUState = Comunication_States.WaitforAnswerSULD
                    RaiseEvent WriteToseroalPort("1" & vbCrLf, clientID)

                Catch ex As Exception
                    WriteLogs("Error step12 : answer suld=1  " & ex.Message, m_ClientID)
                    StrBuffer = ""
                End Try
            ElseIf m_RTUState = Comunication_States.WaitforAnswerSULD And StrBuffer.Contains("OK") Then
                Try
                    StrBuffer = ""
                    m_RTUState = Comunication_States.GetLambereshtLogCount

                    RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "H" & vbCrLf, clientID)


                Catch ex As Exception
                    WriteLogs("step11* : error: " & ex.Message, m_ClientID)
                    StrBuffer = ""
                End Try

            ElseIf m_RTUState = Comunication_States.GetLambereshtLogCount And StrBuffer.Contains(vbCr) Then
                tmr_Timeout.Enabled = False
                Try
                    If StrBuffer.Contains("AT+CPOWD=1") Then
                        StrBuffer = ""
                    Else
                        Dim strSpil() = Split(StrBuffer, "|", -1, vbBinaryCompare)
                        m_Unread_Logs = Val(strSpil(6))
                        m_CurentLog = 1
                        StrBuffer = ""

                        m_RTUState = Comunication_States.ReadLogFromLambresht
                        RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "M" & vbCrLf, clientID)

                        tmr_Timeout.Interval = 30000
                        tmr_Timeout.Enabled = True
                    End If
                Catch ex As Exception
                    WriteLogs("step222* :" & ex.Message, m_ClientID)
                    StrBuffer = ""
                End Try
            ElseIf m_RTUState = Comunication_States.ReadLogFromLambresht And StrBuffer.Contains("AT+CPOWD=1") Then
                ' 
                Try

                    srLogs.Close()
                    StrBuffer = ""
                    RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8
                Catch ex As Exception
                    WriteLogs("step33* :" & ex.Message, m_ClientID)
                    StrBuffer = ""
                End Try
            ElseIf m_RTUState = Comunication_States.ReadLogFromLambresht And StrBuffer.Contains(vbCr) Then
                tmr_Timeout.Enabled = False
                Try
                    WriteLogs("log from logger", clientID)

                    Dim StrLog As String = ""
                    Dim date_header As String = "20" & Mid(StrBuffer, 8, 2) & "/" & Mid(StrBuffer, 6, 2) & "/" & Mid(StrBuffer, 4, 2)
                    Dim time_header As String = Mid(StrBuffer, 12, 2) & ":" & Mid(StrBuffer, 14, 2) & ":" & Mid(StrBuffer, 16, 2)

                    Dim strSpil() = Split(StrBuffer, "|", -1, vbBinaryCompare)
                    Dim strat_step As Integer = 3

                    If StrBuffer.Count = (strSpil.Count - 2) Then
                        strat_step = 1
                    Else
                        strat_step = 3
                    End If
                    For i = strat_step To strSpil.Count - 2 Step strat_step
                        StrLog += Trim(strSpil(i)) & Space(30 - Trim(strSpil(i)).Length)

                        '   strLog += strSpil(i) & ";" 'Trim(strSpil(i)) & Space(25 - Trim(strSpil(i)).Length)
                    Next


                    StrLog = date_header & Space(30 - Trim(date_header).Length) & time_header & Space(30 - Trim(time_header).Length) & StrLog


                    Try
                            srLogs.WriteLine(StrLog)

                        Catch ex As Exception
                            srLogs = New IO.StreamWriter(m_Temp_File)
                            WriteLogs("step100* for IJAD FILE JADID:" & StrLog & ex.Message, m_ClientID)
                        End Try

                        System.Threading.Thread.Sleep(1000)
                    StrBuffer = ""
                    m_CurentLog = m_CurentLog + 1

                    If m_CurentLog <= m_Unread_Logs Then
                        m_RTUState = Comunication_States.ReadLogFromLambresht
                        RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "M" & vbCrLf, clientID)

                        tmr_Timeout.Interval = 30000
                        tmr_Timeout.Enabled = True
                    Else
                        srLogs.Close()

                        RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8
                        '               
                    End If
                Catch ex As Exception
                    WriteLogs("step100 " & ex.Message, m_ClientID)
                End Try


            ElseIf m_RTUState = Comunication_States.WAIT_FOR_01VRESPONCE And Microsoft.VisualBasic.Right(StrBuffer, 1) = Chr(13) Then  '9vbcheck
                tmr_Timeout.Enabled = False
                Dim strrec1 As String = StrBuffer
                strrec1 = DeleteGarbageChars(strrec1)
                '  WritePSTNLog(ObjectID & "  :  " & str_Port_Bufffer)
                ' =FriedrichsCOM1020 .10U4.15
                strrec1 = strrec1.Replace(vbCr, "")
                strrec1 = strrec1.Replace(vbLf, "")
                Try

                    Dim VendorName As String
                    VendorName = Mid(strrec1, 2, 10)
                    Select Case VendorName
                        Case "Partonegar"
                            m_Vendor = Logger_Vendor.Partonegar
                        Case "Lambresht"
                            m_Vendor = Logger_Vendor.Lambresht
                        Case "Theodor", "Friedrichs"
                            m_Vendor = Logger_Vendor.Theodor
                        Case "Thies"

                            m_Vendor = Logger_Vendor.Thies
                        Case Else
                            Exit Sub
                    End Select
                    m_HWRevision = Mid(strrec1, 21, 4)
                    m_SWRevision = Mid(strrec1, 26, 4)
                    m_ModulType = Mid(strrec1, 12, 8)
                    '   str_Port_BuffferHistory = str_Port_Bufffer '24
                    StrBuffer = ""
                    ' '' ''m_RTUState = RTU_STATES.WAIT_FOR_GSL
                    ' '' ''WriteToseroalPort("$GSL?" & vbCrLf)
                    ' '' ''Use Sql To get Sensor List
                    m_RTUState = Comunication_States.CONNECTED
                    RaiseEvent RTU_Info_Receiver(ObjectID, m_Logger_Name, m_Logger_SN, m_Logger_SenNo, 36, m_Memory_Active, m_Unread_Logs, m_Vendor.ToString, m_SWRevision, m_HWRevision, m_ModulType, m_Memory_Active)
                Catch ex As Exception
                End Try
            ElseIf m_RTUState = Comunication_States.WAIT_FOR_01H_RESPONSE And StrBuffer.Contains(vbCr) And StrBuffer.Contains("=") Then  '12vbcheck?
                tmr_Timeout.Enabled = False
                StrBuffer = DeleteGarbageChars(StrBuffer)
                StrBuffer = StrBuffer.Replace("=", "")
                '   RaiseEvent RTU_WriteLog(m_ObjectID, "@$01H  " & Microsoft.VisualBasic.Left(str_Port_Bufffer.Replace(vbCrLf, ""), str_Port_Bufffer.Replace(vbCrLf, "").Length - 3))
                If Microsoft.VisualBasic.Left(StrBuffer.Replace(vbCrLf, ""), StrBuffer.Replace(vbCrLf, "").Length - 3) = Microsoft.VisualBasic.Format(System.DateTime.Now, "yyMMddHHmm") Then
                    '      str_Port_BuffferHistory = str_Port_Bufffer '28
                    StrBuffer = ""
                    RaiseEvent RTU_Synchronized(ObjectID, m_Vendor.ToString, m_Memory_Active)
                Else

                    Try

                        'bln_timeout = False
                        'str_Port_BuffferHistory = str_Port_Bufffer '29
                        StrBuffer = ""
                        m_RTUState = Comunication_States.WAIT_FOR_01G
                        '     RaiseEvent RTU_WriteLog(m_ObjectID, "@$01G   " & Microsoft.VisualBasic.Format(System.DateTime.Now, "yyMMddHHmmss"))
                        RaiseEvent WriteToseroalPort("@$01G" & Microsoft.VisualBasic.Format(System.DateTime.Now, "yyMMddHHmmss") & vbCrLf, m_ClientID)
                    Catch ex As Exception

                    End Try

                End If
                'str_Port_BuffferHistory = str_Port_Bufffer '30
                StrBuffer = ""
            ElseIf m_RTUState = Comunication_States.WAIT_FOR_01G And StrBuffer.Contains(Chr(6)) Then
                tmr_Timeout.Enabled = False
                StrBuffer = ""
                RaiseEvent RTU_Synchronized(ObjectID, m_Vendor.ToString, m_Memory_Active)


            ElseIf m_RTUState = Comunication_States.Waitforlogsfromthies And StrBuffer.Contains("END OF DATA") Then
                tmr_Timeout.Enabled = False

                Try
                    Try
                        WriteLogs("step1 for this endofdata:", m_ClientID)

                        If StrBuffer <> "" Then

                            str_Buffer += StrBuffer
                            StrBuffer = ""

                        End If
                    Catch ex As Exception
                        WriteLogs("step1 for this endofdata:" & ex.Message, m_ClientID)
                    End Try

                    Try
                        srLogs.Write(str_Buffer)
                        str_Buffer = ""
                        WriteLogs("step 2 : IJAD FILE JADID", m_ClientID)

                    Catch ex As Exception
                        srLogs = New IO.StreamWriter(m_Temp_File)
                        srLogs.Write(str_Buffer)
                        str_Buffer = ""
                        WriteLogs("step2 for IJAD FILE JADID:" & ex.Message, m_ClientID)

                    End Try


                    Try
                        srLogs.Close()
                    Catch ex As Exception
                        WriteLogs("step3 for close FILE JADID:" & ex.Message, m_ClientID)
                    End Try

                    Try

                        m_Real_File = m_Temp_File.Replace("temp", "") ' m_datafolder & m_StationCode.ToString & "_" & Format(Now, "yyyyMMdd_HHmmss") & ".log"
                        WriteLogs(" step 4 :  " & m_Real_File, m_ClientID)

                    Catch ex As Exception
                        WriteLogs("step4 for" & ex.Message, m_ClientID)
                    End Try


                    Try
                        Temp2RealValue(m_Temp_File, m_Real_File)
                        WriteLogs("RTU_Final_File_Generated", clientID)
                        RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8


                    Catch ex As Exception
                        WriteLogs("step5 error:" & ex.Message, m_ClientID)
                    End Try

                Catch ex As Exception
                    WriteLogs("step6 error:" & ex.Message, m_ClientID)
                End Try
            ElseIf m_RTUState = Comunication_States.Waitforlogsfromthies And StrBuffer.Contains("AT+CPOWD=1") Then
                tmr_Timeout.Enabled = False

                Try

                    StrBuffer = StrBuffer.Replace("AT+CPOWD=1", "")
                    str_Buffer += StrBuffer
                    StrBuffer = ""
                    srLogs.Write(str_Buffer)
                    str_Buffer = ""
                    WriteLogs("step 22 : IJAD FILE JADID", m_ClientID)

                Catch ex As Exception
                    srLogs = New IO.StreamWriter(m_Temp_File)
                    srLogs.Write(str_Buffer)
                    str_Buffer = ""
                    WriteLogs("step2!*" & str_Buffer & ex.Message, m_ClientID)
                End Try


                Try
                    srLogs.Close()

                Catch ex As Exception
                    WriteLogs("step7 error:" & ex.Message, m_ClientID)
                End Try

                Try

                    m_Real_File = m_Temp_File.Replace("temp", "") ' m_datafolder & m_StationCode.ToString & "_" & Format(Now, "yyyyMMdd_HHmmss") & ".log"
                    WriteLogs(" step 8:  " & m_Real_File, m_ClientID)

                Catch ex As Exception
                    WriteLogs("step9 error:" & ex.Message, m_ClientID)
                End Try


                Try
                    Temp2RealValue(m_Temp_File, m_Real_File)
                    WriteLogs("RTU_File_Generated", clientID)
                    RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                Catch ex As Exception
                    WriteLogs("step10 error:" & ex.Message, m_ClientID)
                End Try

                tmr_Timeout.Interval = 30000
                tmr_Timeout.Enabled = True
            ElseIf m_RTUState = Comunication_States.Waitforlogsfromthies And StrBuffer <> "" Then

                tmr_Timeout.Interval = 60000
                tmr_Timeout.Enabled = True
                WriteLogs("waitforlogs from thies :file temp  ", m_ClientID)

                Try

                    If StrBuffer <> "" Then
                        str_Buffer += StrBuffer
                        StrBuffer = ""


                    End If
                Catch ex As Exception
                    WriteLogs("step11 error:" & ex.Message, m_ClientID)
                End Try

            ElseIf m_RTUState = Comunication_States.WAIT_FOR_DATABLOCK And (StrBuffer.Contains("EOB" & vbCrLf) Or StrBuffer.Contains("EOD" & vbCrLf)) Then
                WriteLogs("Step 1", clientID)
                tmr_Timeout.Enabled = False
                WriteLogs("Step 2", clientID)
                Try


                    If StrBuffer.Contains("EOD" & vbCrLf) Then

                        m_Bloack_Counter += 1
                        m_Data_Persent = 100
                        StrBuffer = StrBuffer.Replace("EOD" & vbCrLf, "")
                        StrBuffer = StrBuffer.Replace("|", vbCrLf)
                        srLogs.Write(StrBuffer)
                        StrBuffer = ""

                        m_RTUState = Comunication_States.WAIT_FOR_CMT_ACK
                        RaiseEvent WriteToseroalPort(vbCrLf & "$CMT" & vbCrLf, clientID)
                        '  ManualCMT()
                    ElseIf StrBuffer.Contains("EOB" & vbCrLf) Then
                        '   tmr_Timeout.Interval = 60000
                        m_Bloack_Counter += 1
                        '   tmr_Timeout.Enabled = True

                        StrBuffer = StrBuffer.Replace("EOB" & vbCrLf, "")
                        StrBuffer = StrBuffer.Replace("|", vbCrLf)
                        srLogs.Write(StrBuffer)
                        StrBuffer = ""
                        m_RTUState = Comunication_States.WAIT_FOR_CMT_ACK
                        RaiseEvent WriteToseroalPort(vbCrLf & "$CMT" & vbCrLf, clientID)
                        '  ManualCMT()
                    End If
                Catch ex As Exception
                    WriteLogs("EOB error:" & ex.Message, clientID)
                End Try
                tmr_Timeout.Interval = 60000
                tmr_Timeout.Enabled = True
            ElseIf m_RTUState = Comunication_States.WAIT_FOR_CMT_ACK And (StrBuffer.Contains("OK" & vbCrLf)) Then
                WriteLogs("WAIT_FOR_CMT_ACK Step 1", clientID)
                tmr_Timeout.Enabled = False


                StrBuffer = ""
                If m_Data_Persent = 100 Then
                    m_RTUState = Comunication_States.CONNECTED
                    'If m_strLog <> "" Then
                    '   srLogs  = New IO.StreamWriter(m_Temp_File)
                    '    srLogs.Write(m_strLog)
                    'End If

                    Try
                        srLogs.Close()
                    Catch ex As Exception
                    End Try
                    Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                    WriteLogs(" fileinfo.Length :" & fileinfo.Length, clientID)
                    If fileinfo.Length > 0 Then


                        Try
                            '  If m_TempfileCreat = False Then 't10

                            m_Real_File = m_Temp_File.Replace("temp", "")
                            Temp2RealValue(m_Temp_File, m_Real_File)
                            '    RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                            WriteLogs("RTU_Final_File_Generated", clientID)
                            RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                            '  End If
                        Catch ex As Exception
                        End Try
                    Else
                        File.Delete(m_Temp_File)
                    End If
                Else
                    ' m_Packet_Size = 36

                    If m_Data_Persent = 100 Then

                        m_RTUState = Comunication_States.CONNECTED
                        'If m_strLog <> "" Then
                        '    srLogs = New IO.StreamWriter(m_Temp_File)
                        '    srLogs.Write(m_strLog)
                        'End If
                        Try
                            srLogs.Close()
                        Catch ex As Exception
                        End Try
                        Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                        If fileinfo.Length > 0 Then


                            Try
                                '  If m_TempfileCreat = False Then 't10

                                m_Real_File = m_Temp_File.Replace("temp", "")
                                Temp2RealValue(m_Temp_File, m_Real_File)
                                '     RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                                WriteLogs("RTU_Final_File_Generated", clientID)
                                RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                                '  End If
                            Catch ex As Exception
                            End Try
                        Else
                            File.Delete(m_Temp_File)
                        End If
                    Else
                        ' RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)


                        m_RTUState = Comunication_States.WAIT_FOR_DATABLOCK
                        RaiseEvent WriteToseroalPort("$RFL" & vbCrLf, m_ClientID) '24hex=36
                        tmr_Timeout.Interval = 30000
                        tmr_Timeout.Enabled = True


                    End If

                End If

            ElseIf m_RTUState = Comunication_States.WaitFor01ZResponse And (StrBuffer.Contains("EOB" & vbCrLf) Or StrBuffer.Contains("EOD" & vbCrLf)) Then
                tmr_Timeout.Enabled = False  '1

                Dim str_Port_BuffferHistory As String = ""
                Try
                    Dim IsErr As Boolean = False
                    Dim logs As String = ""

                    logs = logs & " str_Port_Bufffer  is " & StrBuffer & vbCrLf

                    logs = logs & " call DeleteGarbageChars" & vbCrLf
                    StrBuffer = DeleteGarbageChars(StrBuffer)
                    str_Port_BuffferHistory = StrBuffer '45
                    logs = logs & " end DeleteGarbageChars" & vbCrLf
                    Dim arraylog() = Split(StrBuffer, vbCrLf)

                    StrBuffer = ""
                    Try
                        For i = 0 To arraylog.Length - 1
                            Try


                                If arraylog(i) <> "EOB" Or arraylog(i) <> "EOD" Then
                                    Try

                                        m_Data_Persent = Math.Truncate((m_Bloack_Counter / m_Unread_Logs) * 100)
                                    Catch ex As Exception
                                    End Try

                                    m_Bloack_Counter += 1
                                    arraylog(i) = arraylog(i).Replace("EOB", "")
                                    arraylog(i) = arraylog(i).Replace("EOD", "")
                                    If arraylog(i).Length > 0 Then
                                        srLogs.Write(arraylog(i) & vbCrLf)
                                    End If
                                    '1h


                                End If
                            Catch ex As Exception
                                WriteLogs("WaitFor01ZResponse" & ex.Message, m_ClientID)

                            End Try

                        Next

                    Catch ex As Exception

                        WriteLogs("WaitForResponse: " & ex.Message, m_ClientID)

                    End Try




                    If str_Port_BuffferHistory.Contains("EOD" & vbCrLf) Then
                        m_Data_Persent = 100

                    ElseIf str_Port_BuffferHistory.Contains("EOB" & vbCrLf) Then

                    End If

                    m_RTUState = Comunication_States.WaitFor01ZOKResponse
                    RaiseEvent WriteToseroalPort("@$01ZEOK" & vbCrLf, clientID)
                    tmr_Timeout.Enabled = True
                    tmr_Timeout.Interval = 60000
                Catch ex As Exception

                End Try

            ElseIf m_RTUState = Comunication_States.WaitFor01ZOKResponse And StrBuffer.Contains(Chr(6)) Then
                Try
                    tmr_Timeout.Enabled = False   '2
                    'bln_timeout = False
                    Dim str_Port_BuffferHistory As String = StrBuffer '48
                    StrBuffer = ""
                    If m_Data_Persent = 100 Then
                        ' tmr_Timeout.Enabled = False
                        m_RTUState = Comunication_States.CONNECTED

                        Try
                            srLogs.Close()
                        Catch ex As Exception
                        End Try
                        If m_Temp_File = "" Then
                            m_datafolder = (GetReceiveFolder() & "\" & m_StationCode & "\").Replace("\\", "\")
                            If Not Directory.Exists(m_datafolder) Then
                                Directory.CreateDirectory(m_datafolder)
                            End If
                            m_Temp_File = (GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode & "temp_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt").Replace("\\", "\")
                        End If
                        Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                        WriteLogs(" fileinfo.Length :" & fileinfo.Length, clientID)
                        If fileinfo.Length > 0 Then

                            Try
                                '  If m_TempfileCreat = False Then 't10

                                m_Real_File = m_Temp_File.Replace("temp", "")
                                Temp2RealValue(m_Temp_File, m_Real_File)
                                '     RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                                RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                                '  End If
                            Catch ex As Exception
                            End Try
                        Else
                            File.Delete(m_Temp_File)
                        End If
                        m_RTUState = Comunication_States.WAIT_FOR_GOODBYE
                        RaiseEvent WriteToseroalPort("$DSC" & vbCrLf, clientID)
                    Else

                        'If m_Data_Persent = 100 Then


                        '    m_RTUState = Comunication_States.CONNECTED
                        '    'If m_strLog <> "" Then
                        '    '    srLogs = New IO.StreamWriter(m_Temp_File)
                        '    '    srLogs.Write(m_strLog)
                        '    'End If
                        '    Try
                        '        srLogs.Close()
                        '    Catch ex As Exception
                        '    End Try
                        '    Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                        '    If fileinfo.Length > 0 Then


                        '        Try
                        '            '  If m_TempfileCreat = False Then 't10

                        '            m_Real_File = m_Temp_File.Replace("temp", "")
                        '            Temp2RealValue(m_Temp_File, m_Real_File)
                        '            '     RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                        '            RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                        '            '  End If
                        '        Catch ex As Exception
                        '        End Try
                        '        m_RTUState = Comunication_States.WAIT_FOR_GOODBYE
                        '        RaiseEvent WriteToseroalPort("$DSC" & vbCrLf, clientID)
                        '    Else
                        '        File.Delete(m_Temp_File)
                        '    End If

                        'Else

                        tmr_Timeout.Interval = 30000
                        tmr_Timeout.Enabled = True
                        m_RTUState = Comunication_States.WaitFor01ZResponse
                        RaiseEvent WriteToseroalPort("@$01ZE24" & vbCrLf, clientID)

                    End If


                Catch ex As Exception
                    m_RTUState = Comunication_States.WAIT_FOR_GOODBYE
                    RaiseEvent WriteToseroalPort("$DSC" & vbCrLf, clientID)
                End Try
            ElseIf m_RTUState = Comunication_States.WAIT_FOR_GOODBYE And UCase(StrBuffer).Contains("GOODBYE") Then   '1vbcheck?  hamishe dare ???
                StrBuffer = ""

            ElseIf m_RTUState = Comunication_States.WAIT_FOR_ALARM_CONTENT And (StrBuffer.Contains("|" & vbCrLf) Or StrBuffer.Contains(" " & vbCrLf)) Then   '1vbcheck?  hamishe dare ???
                WriteLogs("ALARM_CONTENT:" & StrBuffer, clientID)
                Try

                    '================================================================================
                    'TEST~>[2010/04/18 11:40:55:RA0:1.6]|
                    'TEST~>[2010/04/18 11:40:55;RA0:1.6;INT:23]|

                    '[2010/04/18 11:40:55;Door:1]|

                    Dim arrayAlarm() As String
                    Dim strDate, strTime As String
                    Dim Value As String



                    If StrBuffer.Contains("|") Then
                        arrayAlarm = Split(StrBuffer, "|")
                        StrBuffer = ""
                        Try

                            For i = 0 To arrayAlarm.Length - 2

                                strDate = Mid(arrayAlarm(i), InStr(arrayAlarm(i), "[", CompareMethod.Text) + 1, 10)
                                strTime = Mid(arrayAlarm(i), InStr(arrayAlarm(i), "[", CompareMethod.Text) + 1 + 10 + 1, 8)
                                ReDim Preserve ALARMS(i)
                                ALARMS(i).A_Date = strDate
                                ALARMS(i).A_Time = strTime
                                ALARMS(i).A_Interval = 800000
                                ALARMS(i).A_DateTime = strDate & " " & strTime
                                Dim len As Integer = arrayAlarm(i).Length - (InStr(arrayAlarm(i), "[", CompareMethod.Text) + 21)
                                arrayAlarm(i) = Microsoft.VisualBasic.Right(arrayAlarm(i), len + 1)
                                Dim AlarmTypeandVale() As String
                                AlarmTypeandVale = Split(arrayAlarm(i), ":")
                                Try
                                    Value = AlarmTypeandVale(1).Replace(";INT", "")
                                    ALARMS(i).A_Value = Convert.ToSingle(Value.Replace("]", ""))
                                    ALARMS(i).A_Interval = Convert.ToInt64(AlarmTypeandVale(2).Replace("]", ""))
                                Catch ex As Exception
                                End Try
                                ALARMS(i).A_Name = AlarmTypeandVale(0)

                            Next
                        Catch ex As Exception
                        End Try
                        '?   m_RTUState = RTU_STATES.CONNECTED
                        RaiseEvent RTU_Alarm(m_ObjectID, clientID, ALARMS)

                    End If
                Catch ex As Exception

                End Try

                m_RTUState = Comunication_States.WAIT_FOR_LOGGER_INFO
                WriteLogs("WAIT_FOR_ALARM_CONTENT Send @$01S", m_ClientID)
                RaiseEvent WriteToseroalPort("@$01S" & Chr(13), clientID)
            ElseIf StrBuffer.Contains("AT+CPOWD=1") Then   '1vbcheck?  hamishe dare ???
                StrBuffer = ""
                Try
                    srLogs.Close()
                Catch ex As Exception
                End Try
                If m_Temp_File = "" Then
                    m_datafolder = (GetReceiveFolder() & "\" & m_StationCode & "\").Replace("\\", "\")
                    If Not Directory.Exists(m_datafolder) Then
                        Directory.CreateDirectory(m_datafolder)
                    End If
                    m_Temp_File = (GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode & "temp_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt").Replace("\\", "\")
                End If
                Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                WriteLogs(" fileinfo.Length :" & fileinfo.Length, clientID)
                If fileinfo.Length > 0 Then

                    Try
                        '  If m_TempfileCreat = False Then 't10

                        m_Real_File = m_Temp_File.Replace("temp", "")
                        Temp2RealValue(m_Temp_File, m_Real_File)
                        '     RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                        RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                        '  End If
                    Catch ex As Exception
                    End Try
                Else
                    File.Delete(m_Temp_File)
                End If
            End If
        Catch ex As InvalidOperationException

        End Try
    End Sub
    Public Sub RecDataForPartonegar(ByVal clientID As Integer, ByVal Buff As String)
        StrBuffer = StrBuffer + Buff

        Buff = ""
        Try


            If m_RTUState = Comunication_States.WAIT_FOR_LOGGER_INFO And StrBuffer.Contains(vbCr) And StrBuffer.Contains("=") Then

                StrBuffer = StrBuffer.Replace(vbCr, "")
                m_Logger_Name = Trim(Mid(StrBuffer, 2, 20))
                m_Logger_SN = Mid(Trim(Right(StrBuffer, 8)), 1, Len(Trim(Right(StrBuffer, 8))) - 2)
                m_Logger_SenNo = ConvHexToDec(Trim(Right(Trim(Right(StrBuffer, 8)), 2)))
                m_SensorCount = m_Logger_SenNo
                '  str_Port_BuffferHistory=str_Port_Bufffer '17
                SensorIndex = 0
                StrBuffer = ""
                Array.Clear(Sensors, 0, Sensors.Length - 1)
                WriteLogs("Logger_SenNo:" & m_Logger_SenNo, clientID)
                RaiseEvent WriteToseroalPort("$01B00" & vbCr, m_ClientID)
                m_RTUState = Comunication_States.WaitforSensorList



            ElseIf m_RTUState = Comunication_States.WaitforSensorList And StrBuffer.Contains(vbCr) Then
                '=?Rain                ???mm    ??
                Dim strSpilit() = Split(StrBuffer, "?")
                StrBuffer = ""
                Sensors(SensorIndex).Name = Trim(strSpilit(1))
                WriteLogs("SensorIndex:" & SensorIndex, clientID)
                WriteLogs("Name:" & Sensors(SensorIndex).Name, clientID)
                If SensorIndex < m_SensorCount - 1 Then
                    SensorIndex = SensorIndex + 1
                    ConvDecToHex(SensorIndex)
                    RaiseEvent WriteToseroalPort("$01B" & ConvDecToHex(SensorIndex) & vbCr, m_ClientID)
                    m_RTUState = Comunication_States.WaitforSensorList



                Else
                    RaiseEvent WriteToseroalPort("$01N" & vbCr, m_ClientID)
                    m_RTUState = Comunication_States.WAIT_FOR_UNREAD_LOGS
                End If



            ElseIf m_RTUState = Comunication_States.WAIT_FOR_UNREAD_LOGS And ((m_Memory_Active And StrBuffer.Contains(vbCrLf)) Or (Not m_Memory_Active And StrBuffer.Contains(vbCr))) Then

                Dim strrec As String = StrBuffer   '7vbcheck?
                Try
                    If m_Memory_Active Then
                    Else
                        strrec = DeleteGarbageChars(strrec)
                        strrec = Trim(strrec).Replace("=", "")
                    End If
                    If IsNumeric(Trim(strrec)) Then
                        m_Unread_Logs = Val(Trim(strrec))
                        '   InsertTOtblCallInfo(9, RTUPort(ObjectID).ContactID, "Periodic", "Unread logs is " & m_Unread_Logs)
                        '   str_Port_BuffferHistory = str_Port_Bufffer '22
                        StrBuffer = ""
                        If m_Vendor = 4 Or m_Vendor = 2 Or m_Vendor = 3 Or m_Vendor = 22 Then
                            m_RTUState = Comunication_States.WaitforVer
                            RaiseEvent WriteToseroalPort("$VER" & vbCrLf, m_ClientID)
                        Else
                            m_RTUState = Comunication_States.WAIT_FOR_01VRESPONCE
                            RaiseEvent WriteToseroalPort("$01V" & vbCr, clientID)
                        End If
                    End If
                Catch ex As Exception

                End Try
            ElseIf m_RTUState = Comunication_States.WAIT_FOR_01VRESPONCE And Microsoft.VisualBasic.Right(StrBuffer, 1) = Chr(13) Then  '9vbcheck

                Dim strrec1 As String = StrBuffer
                strrec1 = DeleteGarbageChars(strrec1)
                '  WritePSTNLog(ObjectID & "  :  " & str_Port_Bufffer)
                ' =FriedrichsCOM1020 .10U4.15
                strrec1 = strrec1.Replace(vbCr, "")
                strrec1 = strrec1.Replace(vbLf, "")
                Try

                    Dim VendorName As String
                    VendorName = Mid(strrec1, 2, 10)
                    Select Case VendorName
                        Case "Partonegar"
                            m_Vendor = Logger_Vendor.Partonegar
                        Case "Lambresht"
                            m_Vendor = Logger_Vendor.Lambresht
                        Case "Theodor", "Friedrichs"
                            m_Vendor = Logger_Vendor.Theodor
                        Case "Thies"

                            m_Vendor = Logger_Vendor.Thies
                        Case Else
                            Exit Sub
                    End Select
                    m_HWRevision = Mid(strrec1, 21, 4)
                    m_SWRevision = Mid(strrec1, 26, 4)
                    m_ModulType = Mid(strrec1, 12, 8)
                    '   str_Port_BuffferHistory = str_Port_Bufffer '24
                    StrBuffer = ""
                    ' '' ''m_RTUState = RTU_STATES.WAIT_FOR_GSL
                    ' '' ''WriteToseroalPort("$GSL?" & vbCrLf)
                    ' '' ''Use Sql To get Sensor List
                    m_RTUState = Comunication_States.CONNECTED
                    RaiseEvent RTU_Info_Receiver(ObjectID, m_Logger_Name, m_Logger_SN, m_Logger_SenNo, 36, m_Memory_Active, m_Unread_Logs, m_Vendor.ToString, m_SWRevision, m_HWRevision, m_ModulType, m_Memory_Active)
                Catch ex As Exception
                End Try
            ElseIf m_RTUState = Comunication_States.WAIT_FOR_01H_RESPONSE And StrBuffer.Contains(vbCr) And StrBuffer.Contains("=") Then  '12vbcheck?
                StrBuffer = DeleteGarbageChars(StrBuffer)
                StrBuffer = StrBuffer.Replace("=", "")
                '   RaiseEvent RTU_WriteLog(m_ObjectID, "@$01H  " & Microsoft.VisualBasic.Left(str_Port_Bufffer.Replace(vbCrLf, ""), str_Port_Bufffer.Replace(vbCrLf, "").Length - 3))
                If Microsoft.VisualBasic.Left(StrBuffer.Replace(vbCrLf, ""), StrBuffer.Replace(vbCrLf, "").Length - 3) = Microsoft.VisualBasic.Format(System.DateTime.Now, "yyMMddHHmm") Then
                    '      str_Port_BuffferHistory = str_Port_Bufffer '28
                    StrBuffer = ""
                    RaiseEvent RTU_Synchronized(ObjectID, m_Vendor.ToString, m_Memory_Active)
                Else

                    Try

                        'bln_timeout = False
                        'str_Port_BuffferHistory = str_Port_Bufffer '29
                        StrBuffer = ""
                        m_RTUState = Comunication_States.WAIT_FOR_01G
                        '     RaiseEvent RTU_WriteLog(m_ObjectID, "@$01G   " & Microsoft.VisualBasic.Format(System.DateTime.Now, "yyMMddHHmmss"))
                        RaiseEvent WriteToseroalPort("$01G" & Microsoft.VisualBasic.Format(System.DateTime.Now, "yyMMddHHmmss") & vbCrLf, m_ClientID)
                    Catch ex As Exception

                    End Try

                End If
                'str_Port_BuffferHistory = str_Port_Bufffer '30
                StrBuffer = ""
            ElseIf m_RTUState = Comunication_States.WAIT_FOR_01G And StrBuffer.Contains(Chr(6)) Then
                StrBuffer = ""
                RaiseEvent RTU_Synchronized(ObjectID, m_Vendor.ToString, m_Memory_Active)
            ElseIf m_RTUState = Comunication_States.WaitFor01ZResponse And (StrBuffer.Contains("EOB" & vbCrLf) Or StrBuffer.Contains("EOD" & vbCrLf)) Then
                WriteLogs("Step 1", clientID)
                tmr_Timeout.Enabled = False  '1
                WriteLogs("Step 2", clientID)
                Dim str_Port_BuffferHistory As String = ""
                Try
                    Dim IsErr As Boolean = False
                    Dim logs As String = ""

                    logs = logs & " str_Port_Bufffer  is " & StrBuffer & vbCrLf

                    logs = logs & " call DeleteGarbageChars" & vbCrLf
                    StrBuffer = DeleteGarbageChars(StrBuffer)
                    str_Port_BuffferHistory = StrBuffer '45
                    logs = logs & " end DeleteGarbageChars" & vbCrLf
                    Dim arraylog() = Split(StrBuffer, vbCrLf)

                    StrBuffer = ""
                    Try
                        For i = 0 To arraylog.Length - 1
                            Try


                                If arraylog(i) <> "EOB" Or arraylog(i) <> "EOD" Then
                                    Try

                                        m_Data_Persent = Math.Truncate((m_Bloack_Counter / m_Unread_Logs) * 100)
                                    Catch ex As Exception
                                    End Try

                                    m_Bloack_Counter += 1
                                    arraylog(i) = arraylog(i).Replace("EOB", "")
                                    arraylog(i) = arraylog(i).Replace("EOD", "")
                                    If arraylog(i).Length > 0 Then
                                        srLogs.Write(arraylog(i) & vbCrLf)
                                    End If
                                    '1h


                                End If
                            Catch ex As Exception
                                WriteLogs(ex.Message, m_ClientID)

                            End Try

                        Next

                    Catch ex As Exception

                        WriteLogs("eob error" & ex.Message, clientID)
                    End Try




                    If str_Port_BuffferHistory.Contains("EOD" & vbCrLf) Then
                        m_Data_Persent = 100

                    ElseIf str_Port_BuffferHistory.Contains("EOB" & vbCrLf) Then

                    End If

                    m_RTUState = Comunication_States.WaitFor01ZOKResponse
                    RaiseEvent WriteToseroalPort("$01ZEOK" & vbCrLf, clientID)
                    tmr_Timeout.Enabled = True
                    tmr_Timeout.Interval = 30000
                Catch ex As Exception
                    WriteLogs("error " & ex.Message, clientID)
                End Try

            ElseIf m_RTUState = Comunication_States.WaitFor01ZOKResponse And StrBuffer.Contains(Chr(6)) Then
                WriteLogs("WaitFor01ZOKResponse step 1", clientID)
                Try
                    tmr_Timeout.Enabled = False   '2
                    'bln_timeout = False
                    Dim str_Port_BuffferHistory As String = StrBuffer '48
                    StrBuffer = ""
                    WriteLogs("WaitFor01ZOKResponse step 2", clientID)
                    If m_Data_Persent = 100 Then
                        ' tmr_Timeout.Enabled = False
                        m_RTUState = Comunication_States.CONNECTED
                        'If m_strLog <> "" Then
                        '    srLogs = New IO.StreamWriter(m_Temp_File)
                        '    srLogs.Write(m_strLog)
                        'End If

                        Try
                            srLogs.Close()
                        Catch ex As Exception
                        End Try
                        Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                        WriteLogs(" fileinfo.Length :" & fileinfo.Length, clientID)
                        If fileinfo.Length > 0 Then


                            Try
                                '  If m_TempfileCreat = False Then 't10

                                m_Real_File = m_Temp_File.Replace("temp", "")
                                Temp2RealValue(m_Temp_File, m_Real_File)
                                '     RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                                RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                                '  End If
                            Catch ex As Exception
                            End Try
                        Else
                            File.Delete(m_Temp_File)
                        End If
                    Else

                        If m_Data_Persent = 100 Then


                            m_RTUState = Comunication_States.CONNECTED
                            'If m_strLog <> "" Then
                            '    srLogs = New IO.StreamWriter(m_Temp_File)
                            '    srLogs.Write(m_strLog)
                            'End If
                            Try
                                srLogs.Close()
                            Catch ex As Exception
                            End Try
                            Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                            If fileinfo.Length > 0 Then


                                Try
                                    '  If m_TempfileCreat = False Then 't10

                                    m_Real_File = m_Temp_File.Replace("temp", "")
                                    Temp2RealValue(m_Temp_File, m_Real_File)
                                    '     RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                                    RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                                    '  End If
                                Catch ex As Exception
                                End Try
                            Else
                                File.Delete(m_Temp_File)
                            End If
                        Else
                            ' RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)

                            tmr_Timeout.Interval = 30000
                            tmr_Timeout.Enabled = True
                            m_RTUState = Comunication_States.WaitFor01ZResponse
                            RaiseEvent WriteToseroalPort("$01ZE24" & vbCrLf, clientID)

                        End If

                    End If
                Catch ex As Exception
                    WriteLogs("WaitFor01ZOKResponse error:" & ex.Message, clientID)
                End Try



            End If
        Catch ex As InvalidOperationException
            WriteLogs("WaitFor01ZOKResponse error 1:" & ex.Message, clientID)
        End Try
    End Sub

    Public Function SensorID2Name(ByVal SID As String) As String
        Select Case SID
            Case "92", "93", "94", "99"
                SensorID2Name = "Temperature"
            Case "10", "11", "12", "14", "16", "17", "18", "19", "1C", "1F", "25", "CE", "E1", "15", "3B"
                SensorID2Name = "Wind Speed"
            Case "27", "28", "29", "2A", "2B", "2C", "CF", "E0", "FC"
                SensorID2Name = "Wind Direction"
            Case "30", "31", "32", "36"
                SensorID2Name = "Precipitation"
            Case "80", "83"
                SensorID2Name = "Rel. Humidity"
            Case "40", "41", "42", "48"
                SensorID2Name = "Radiation"
            Case "A0", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "B6", "B7", "AF"
                SensorID2Name = "Air Pressure"
            Case "61", "64"
                SensorID2Name = "Battery Volt."
            Case "62", "65"
                SensorID2Name = "Battery Curr."
            Case "35", "37", "38"
                SensorID2Name = "Evaporation"
            Case "9A"
                SensorID2Name = "Dry Bulb"
            Case "9B"
                SensorID2Name = "Wet Bulb"
            Case "87"
                SensorID2Name = "Soil Moisture"
            Case "FC"
                SensorID2Name = "Visibility"
            Case Else
                SensorID2Name = "???"
        End Select
    End Function
    Public Sub RecDataForLamberesht(ByVal clientID As Integer, ByVal Buff As String)
        '' dar in ghesmat varde bakhshe lamberesht mishavad az rtu dastor list senesor
        '' migird va ba dastor $Suld varde khandane faramin az khode logger mishavad
        If m_RTUState = Comunication_States.ReadLogFromLambresht Then
            tmr_Timeout.Enabled = False
            tmr_Timeout.Interval = 10000
            tmr_Timeout.Enabled = True
        End If


        StrBuffer = StrBuffer + Buff
        Buff = ""
        Try
            If m_RTUState = Comunication_States.WaitForLambereshtVer And StrBuffer.Contains(vbCr) Then



                m_Vendor = Logger_Vendor.Lambresht ' "LAMBRECHT"
                Dim strSpil() = Split(StrBuffer, "_", -1, vbBinaryCompare)
                m_Logger_Name = Trim(Mid(strSpil(0), 2))
                m_HWRevision = Trim(strSpil(1))
                m_SWRevision = "-"
                StrBuffer = ""
                m_RTUState = Comunication_States.WaitforSensorList
                RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "H" & Chr(LAMB.CR), clientID)



            ElseIf m_RTUState = Comunication_States.WaitforSensorList And StrBuffer.Contains(vbCr) Then
                Try
                    Dim strSpil() = Split(StrBuffer, "|", -1, vbBinaryCompare)
                    m_Logger_SenNo = strSpil.Count - 9
                    SensorIndex = 0
                    For i = 8 To strSpil.Count - 2
                        Sensors(SensorIndex).Name = SensorID2Name(Mid(strSpil(i), 3, 2)) & Mid(strSpil(i), 1, 2)
                        SensorIndex = SensorIndex + 1

                    Next
                Catch ex As Exception

                Finally

                End Try
                StrBuffer = ""
                m_RTUState = Comunication_States.ReadLambreshtDateTime
                RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "A" & Chr(LAMB.CR), clientID)



            ElseIf m_RTUState = Comunication_States.ReadLambreshtDateTime And StrBuffer.Contains(vbCr) Then
                If StrBuffer.Contains(Chr(LAMB.CR)) Then
                    StrBuffer = StrBuffer.Substring(2, 14)
                    'WriteLogs("Raw Date/Time:" & res_buffer)
                    Dim strdatetime As String = "20" & Mid(StrBuffer, 5, 2) & "/" & Mid(StrBuffer, 3, 2) & "/" & Mid(StrBuffer, 1, 2) & " " & Mid(StrBuffer, 9, 2) & ":" & Mid(StrBuffer, 11, 2) '& ":" & Mid(StrBuffer, 13, 2)

                    'If strdatetime = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm") Then
                    '    'settime

                    '    Dim strDay As String
                    '    Dim strMonth As String
                    '    Dim strYear As String
                    '    Dim strDOW As String
                    '    Dim strHour As String
                    '    Dim strMinute As String
                    '    Dim strSecond As String
                    '    Dim strHOSecond As String
                    '    Dim strCommand As String
                    '    Dim bk As Long = 4
                    '    Try
                    '        strDay = Format(Val(System.DateTime.Now.Day), "0#")
                    '        bk = 5
                    '        strMonth = Format(Val(System.DateTime.Now.Month), "0#")
                    '        bk = 6
                    '        strYear = System.DateTime.Now.Year.ToString.Substring(2, 2)
                    '        bk = 7
                    '        strHour = Mid(Format(System.DateTime.Now, "HHmmss"), 1, 2)
                    '        bk = 8
                    '        strMinute = Mid(Format(System.DateTime.Now, "HHmmss"), 3, 2)
                    '        bk = 9
                    '        strSecond = Mid(Format(System.DateTime.Now, "HHmmss"), 5, 2)
                    '        bk = 10
                    '        strHOSecond = "00"
                    '        strDOW = Format(Val(System.DateTime.Now.DayOfWeek), "0#")
                    '        bk = 11
                    '        strCommand = "J" & "00" & strSecond & strMinute & strHour & strDOW & strDay & strMonth & strYear
                    '        'WriteLogs("Set Clock Command:" & strCommand)
                    '        bk = 12
                    '    Catch ex As Exception

                    '    End Try
                    '    StrBuffer = ""
                    '    m_RTUState = Comunication_States.SetLambereshtDateTime
                    '    RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & strCommand & Chr(LAMB.ETB) & CheckSumGen(Chr(LAMB.STX) & Chr(48)), clientID)

                    'Else
                    'readlog
                    StrBuffer = ""
                    m_RTUState = Comunication_States.GetLambereshtLogCount
                    RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "H" & Chr(LAMB.CR), clientID)
                    '  End If


                End If

                'ElseIf StrBuffer.Contains(Chr(LAMB.STX) & Chr(LAMB.ACK) & Chr(LAMB.CR)) And StrBuffer.Contains(vbCr) Then
                '    '===============================================================
                '    ' Read a log from first logger
                '    '===============================================================
                '    'ReadLogFromLambresht
                '    StrBuffer = ""
                '    m_RTUState = Comunication_States.GetLambereshtLogCount
                '    RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "H" & Chr(LAMB.CR), clientID)
                '    'delay_ms(200)
            ElseIf m_RTUState = Comunication_States.GetLambereshtLogCount And StrBuffer.Contains(vbCr) Then
                Dim strSensors As String = GetConfigItem(m_StationCode & "AliasName")

                Dim strSpil() = Split(StrBuffer, "|", -1, vbBinaryCompare)
                m_Unread_Logs = Val(strSpil(6))
                m_CurentLog = 1
                m_Temp_File = (GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode & "temp_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt").Replace("\\", "\")

                If Not Directory.Exists(GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode) Then
                    Directory.CreateDirectory(GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode)
                End If
                m_Real_File = m_Temp_File.Replace("temp", "")
                srLogs = New IO.StreamWriter(m_Real_File)
                Dim strLog As String = ""
                strLog = m_StationCode.ToString & Space(30 - m_StationCode.ToString.Length) & m_StationID & Space(30 - m_StationID.Length) & m_Vendor.ToString & Space(30 - m_Vendor.ToString.Length) & vbCrLf
                strLog = strLog & "DATE" & Space(30 - "DATE".Length) & "Time" & Space(30 - "Time".Length)
                Dim AliasName() = Split(strSensors, ";")
                For i As Integer = 0 To AliasName.Count - 1
                    strLog = strLog & AliasName(i) & Space(30 - AliasName(i).Length)
                Next


                srLogs.WriteLine(strLog)
                StrBuffer = ""
                m_RTUState = Comunication_States.ReadLogFromLambresht
                RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "M" & Chr(LAMB.CR), clientID)
                tmr_Timeout.Interval = 30000
                tmr_Timeout.Enabled = True
            ElseIf m_RTUState = Comunication_States.ReadLogFromLambresht And StrBuffer.Contains(vbCr) Then



                Dim date_header As String = "20" & Mid(StrBuffer, 8, 2) & "/" & Mid(StrBuffer, 6, 2) & "/" & Mid(StrBuffer, 4, 2)
                Dim time_header As String = Mid(StrBuffer, 12, 2) & ":" & Mid(StrBuffer, 14, 2) & ":" & Mid(StrBuffer, 16, 2)
                ' Dim date_time_header As String = "20" & Mid(StrBuffer, 8, 2) & Mid(StrBuffer, 6, 2) & Mid(StrBuffer, 4, 2) & Mid(StrBuffer, 12, 2) & Mid(StrBuffer, 14, 2) & Mid(StrBuffer, 16, 2) & ";"
                Dim strLog As String = ""
                'strLog = m_StationCode.ToString & Space(30 - m_StationCode.ToString.Length) & m_StationID & Space(30 - m_StationID.Length) & m_Vendor.ToString & Space(30 - m_Vendor.ToString.Length) & vbCrLf
                'strLog = strLog & "DATE" & Space(30 - "DATE".Length) & "Time" & Space(30 - "Time".Length)
                'Dim AliasName() = Split(strSensors, ";")
                'For i As Integer = 0 To AliasName.Count - 1
                '    strLog = strLog & AliasName(i) & Space(30 - AliasName(i).Length)
                'Next

                ' strLog = strLog & vbCrLf
                '  WriteErrorLog("temp" & Sensors(i).Name)
                Dim strSpil() = Split(StrBuffer, "|", -1, vbBinaryCompare)
                Dim strat_step As Integer = 3

                If StrBuffer.Count = (strSpil.Count - 2) Then
                    strat_step = 1
                Else
                    strat_step = 3
                End If
                For i = strat_step To strSpil.Count - 2 Step strat_step
                    strLog += Trim(strSpil(i)) & Space(30 - Trim(strSpil(i)).Length)
                    '   strLog += strSpil(i) & ";" 'Trim(strSpil(i)) & Space(25 - Trim(strSpil(i)).Length)
                Next

                strLog = date_header & Space(30 - Trim(date_header).Length) & time_header & Space(30 - Trim(time_header).Length) & strLog


                srLogs.WriteLine(strLog)

                System.Threading.Thread.Sleep(1000)

                '    GenerateLambreshtLogs(m_Temp_File, m_Real_File)
                StrBuffer = ""
                m_CurentLog = m_CurentLog + 1
                If m_CurentLog <= m_Unread_Logs Then
                    m_RTUState = Comunication_States.ReadLogFromLambresht
                    RaiseEvent WriteToseroalPort(Chr(LAMB.STX) & Chr(48) & Chr(49) & Chr(LAMB.esc) & "M" & Chr(LAMB.CR), clientID)
                    tmr_Timeout.Interval = 30000
                    tmr_Timeout.Enabled = True
                Else
                    srLogs.Close()
                    RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8
                    '                  MsgBox("")
                End If
            ElseIf m_RTUState = Comunication_States.ReadLogFromLambresht And StrBuffer.Contains("AT+CPOWD=1") Then
                ' 
                Try
                    srLogs.Close()
                    RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8
                Catch ex As Exception

                End Try


                '                  MsgBox("")

            End If
        Catch ex As InvalidOperationException

        End Try
    End Sub

    Public Function GetConfigItem(ByVal item As String) As Object
        Dim key As RegistryKey = Registry.LocalMachine
        Dim subkey As RegistryKey
        Try

            '77865AliasName

            subkey = key.OpenSubKey("Software\Partonegar\GPRS", True)
            GetConfigItem = CStr(subkey.GetValue(item))

            subkey.Close()
        Catch ex As Exception
            GetConfigItem = ""
        End Try


        key.Close()

    End Function
    Public Function CheckSumGen(ByVal strIn As String) As String
        Dim i As Long
        Dim lngCS As Long
        lngCS = 0
        For i = 1 To Len(strIn)
            lngCS = lngCS Xor Asc(Mid(strIn, i, 1))
        Next i
        CheckSumGen = Format(lngCS, "00#")
    End Function
    Private Sub tmr_Timeout_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles tmr_Timeout.Elapsed
        tmr_Timeout.Enabled = False
        WriteLogs("vard shodan be time out", m_ClientID)
        Try


            If m_RTUState = Comunication_States.WaitFor01ZResponse Or m_RTUState = Comunication_States.WaitFor01ZOKResponse Then
                Try
                    ''WriteCheck("waitfor logs from thies in tmr-timeou")
                    srLogs.Close()
                    srLogs.Dispose()
                Catch ex As Exception
                    ''WriteCheck(" error waitfor logs from thies in tmr-timeou" & ex.Message)
                End Try

                Dim fileinfo As New System.IO.FileInfo(m_Temp_File)
                If fileinfo.Length > 0 Then


                    Try
                        '  If m_TempfileCreat = False Then 't10
                        ''WriteCheck("in tmr-timer out for real file  m_Real_File = m_Temp_File.Replace(temp")
                        m_Real_File = m_Temp_File.Replace("temp", "")
                        ''WriteCheck("in tmr-timer out for real file   m_Temp_File.vojod nadarad")

                        Temp2RealValue(m_Temp_File, m_Real_File)
                        '     RaiseEvent RTU_Data_Block_Receive(ObjectID, ContactID, m_Bloack_Counter * m_Packet_Size, m_Data_Persent, m_Vendor.ToString, m_Memory_Active)
                        RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

                        '  End If
                    Catch ex As Exception
                        ''WriteCheck("error in tmr-timer out for real file " & ex.Message)
                    End Try
                End If

            ElseIf m_RTUState = Comunication_States.Waitforlogsfromthies Then

                Try
                    WriteLogs("Timeout : IJAD FILE JADID", m_ClientID)
                    srLogs.Write(str_Buffer)
                    str_Buffer = ""

                Catch ex As Exception
                    srLogs = New IO.StreamWriter(m_Temp_File)
                    srLogs.Write(str_Buffer)
                    str_Buffer = ""
                    WriteLogs("step23 for IJAD FILE JADID:" & ex.Message, m_ClientID)

                End Try


                Try
                    srLogs.Close()

                Catch ex As Exception
                    WriteLogs("step73 error:" & ex.Message, m_ClientID)
                End Try

                Try

                    m_Real_File = m_Temp_File.Replace("temp", "") ' m_datafolder & m_StationCode.ToString & "_" & Format(Now, "yyyyMMdd_HHmmss") & ".log"
                    WriteLogs(" step 83:  " & m_Real_File, m_ClientID)

                Catch ex As Exception
                    WriteLogs("step93 error:" & ex.Message, m_ClientID)

                End Try


                Try

                    Temp2RealValue(m_Temp_File, m_Real_File)
                    WriteLogs(" Time out : RTU_Final_File_Generated", ClientID)
                    RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8


                Catch ex As Exception
                    WriteLogs("step103 error:" & ex.Message, m_ClientID)

                End Try

            ElseIf m_RTUState = Comunication_States.ReadLogFromLambresht Then

                WriteLogs("Timeout : IJAD FILE JADID FOR LAMBERESHT", m_ClientID)

                RaiseEvent RTU_Final_File_Generated(ObjectID, m_Real_File, m_StationCode) '8

            Else
                File.Delete(m_Temp_File)
            End If

        Catch ex As Exception
            WriteLogs("error in TimeOut:" & ex.Message, m_ClientID)

        End Try
    End Sub



    'Step 6-1-1   RTU_Synchronized


    'Step 6-1-2


    Public Function GetSnycron() As Boolean
        'If My.Settings.RTUSync Then
        '    GetSnycron = True
        'Else
        '    GetSnycron = False
        'End If
    End Function
    Public Sub RTU_Synchronizing()
        Try



            If m_Vendor = Logger_Vendor.Partonegar Or m_Vendor = Logger_Vendor.Theodor Then


                StrBuffer = ""
                If IsReadFromRTU Then
                    m_RTUState = Comunication_States.WAIT_FOR_01H_RESPONSE
                    RaiseEvent WriteToseroalPort("@$01H" & vbCrLf, m_ClientID)
                Else
                    m_RTUState = Comunication_States.WAIT_FOR_01H_RESPONSE
                    RaiseEvent WriteToseroalPort("$01H" & vbCrLf, m_ClientID)
                End If


                ' End If

            End If
        Catch ex As Exception
            '   RaiseEvent RTU_WriteError(m_ObjectID, "RTU_Synchronizing" & ex.Message)
        End Try

        '

    End Sub
    'Public Sub RTU_Data_Block_Receive(ByVal ObjectID As Long, ByVal ContactID As Long, ByVal Block_Counter As Long, ByVal Percent As Byte, ByVal Vendor As String, ByVal MemoryActive As Boolean)
    '    'Try
    '    '    If Not ReadDataStart Then
    '    '        If RTUPort(ObjectID).ContactID = -1 Then
    '    '            While flagGetConntactID
    '    '                Application.DoEvents()
    '    '            End While
    '    '            flagGetConntactID = True
    '    '            RTUPort(ObjectID).ContactID = GetCallNUm() + 1
    '    '            InsertTOtblContact(RTUPort(ObjectID).ContactID, RTUPort(ObjectID).StationName, RTUPort(ObjectID).StationCode, ObjectID, "Periodic")
    '    '            flagGetConntactID = False
    '    '        End If
    '    '        'If Not (Vendor = "Thiestdl16" Or ((Not MemoryActive) And Vendor = "Thies")) Then
    '    '        If Not (Vendor = "Thiestdl16" Or (Vendor = "Thies")) Then
    '    '            InsertTOtblCallInfo(4, RTUPort(ObjectID).ContactID, CallType)
    '    '            InsertTOtblCallInfo(5, RTUPort(ObjectID).ContactID, CallType)
    '    '            updateReadDataPercentIntblCallInfo(RTUPort(ObjectID).ContactID, 100)
    '    '        End If

    '    '        ReadDataStart = True
    '    '    End If
    '    '    WriteDCSLog("Block (" & Block_Counter & ") has been received" & Percent.ToString & "%" & vbCrLf)
    '    '    '    If Not (Vendor = "Thiestdl16" Or (Not MemoryActive And Vendor = "Thies")) Then
    '    '    If Not (Vendor = "Thiestdl16" Or (Vendor = "Thies")) Then
    '    '        Dim intPercent As Single = Math.Round(Percent, 0)
    '    '        updateReadDataPercentIntblCallInfo(RTUPort(ObjectID).ContactID, intPercent)
    '    '    End If
    '    '    If Percent = 100 Then
    '    '        WriteDCSLog("Data receiption has been finished" & vbCrLf)
    '    '        ' WriteRTULog("Hang 19")
    '    '        flagHangingStart = True
    '    '        RTUPort(ObjectID).Hangup()
    '    '    End If
    '    'Catch ex As Exception
    '    '    WriteRTUError("RTU_Data_Block_Receive Error is  " & " : " & ex.Message)
    '    'End Try

    'End Sub
    Public Sub GetAlarmContent()
        Try


            tmr_Timeout.Interval = 60000
            tmr_Timeout.Enabled = False

            tmr_Timeout.Enabled = True
            WriteLogs("Send $GAV?", ClientID)
            m_RTUState = Comunication_States.WAIT_FOR_ALARM_CONTENT
            RaiseEvent WriteToseroalPort("$GAV?" & vbCrLf, m_ClientID) '24hex=36



        Catch ex As Exception

        End Try
    End Sub
    Public Function GetLogs(MemoryActive As Boolean) As Long
        tmr_Timeout.Enabled = False

        Try
            m_datafolder = (GetReceiveFolder() & "\" & m_StationCode & "\").Replace("\\", "\")
            If Not Directory.Exists(m_datafolder) Then
                Directory.CreateDirectory(m_datafolder)
            End If


            m_Temp_File = (GetReceiveFolder() & "\" & m_StationCode & "\" & m_StationCode & "temp_" & Format(Now, "yyyyMMdd_HHmm") & "00" & ".txt").Replace("\\", "\")

            Dim lngErr As Long = 0


        Catch ex As Exception
            WriteLogs("GetLogs error 00:" & ex.Message, m_ClientID)
        End Try

        Try

            m_Bloack_Counter = 0
            m_Data_Persent = 0

            Try

                WriteLogs("m_Temp_File:" & m_Temp_File.ToString, m_ClientID)
            Catch ex As Exception
                WriteLogs("GetLogs error 0:" & ex.Message, m_ClientID)
            End Try

            Try

                '   WriteLogs("GetLogs error 0:" & ex.Message, m_ClientID)
                If IsReadFromRTU Then

                    If UCase(m_Vendor.ToString).Contains("THIES") Then
                        WriteLogs("Step 1 get log of thies", m_ClientID)
                        srLogs = New IO.StreamWriter(m_Temp_File)
                        Dim dt As DateTime
                        Dim LastReadOut As String = ""
                        Dim first_line As Boolean

                        Try
                            m_RTUState = Comunication_States.Waitforlogsfromthies
                            RaiseEvent WriteToseroalPort("@" & Chr(2) & "ds" & GetNextDateTimePointerForThies(m_StationCode, m_ClientID) & Chr(3) & vbCrLf, ClientID)

                        Catch ex As Exception
                            WriteLogs("GetNextDateTimePointerForThies():" & ex.Message, m_ClientID)

                        End Try

                    ElseIf UCase(m_Vendor.ToString).Contains("VAISALA") Or MemoryActive = True Then
                        WriteLogs("Step 1 get log from VAISALA", m_ClientID)
                        m_RTUState = Comunication_States.WAIT_FOR_DATABLOCK
                        RaiseEvent WriteToseroalPort("$RFL" & vbCrLf, m_ClientID) '24hex=36

                    Else 'If LoggerType = "Partonegar" Then
                        m_RTUState = Comunication_States.WaitFor01ZResponse
                        RaiseEvent WriteToseroalPort("@$01ZE24" & vbCrLf, m_ClientID) '24hex=36
                    End If


                Else
                    m_RTUState = Comunication_States.WaitFor01ZResponse
                    RaiseEvent WriteToseroalPort("$01ZE24" & vbCrLf, m_ClientID) '24hex=36
                End If


                    tmr_Timeout.Enabled = True
                tmr_Timeout.Interval = 30000

                If UCase(m_Vendor.ToString).Contains("THIES") Then

                Else

                    srLogs = New IO.StreamWriter(m_Temp_File)
                    WriteLogs("GetLogs start **", m_ClientID)
                End If

            Catch ex As Exception
                GetLogs = 0
                WriteLogs("GetLogs error 1:" & ex.Message, m_ClientID)
            End Try
        Catch ex As Exception
            WriteLogs("GetLogs error 2:" & ex.Message, m_ClientID)
        End Try

    End Function
#Region "Theis"
    Public Function ConvDateTimeToThiesCoding(ByVal DateString As String, ByVal TimeString As String) As String
        Dim lngDay As Long
        Dim lngMonth As Long
        Dim lngYear As Long
        Dim lngHour As Long
        Dim lngMinute As Long
        '2012/11/28
        '15:27
        lngMinute = Val(Mid(TimeString, 4, 2)) + 28
        lngHour = Val(Mid(TimeString, 1, 2)) + 28
        lngYear = Val(Mid(DateString, 3, 2)) + 28
        lngMonth = Val(Mid(DateString, 6, 2)) + 28
        lngDay = Val(Mid(DateString, 9, 2)) + 28
        'ConvDateTimeToThiesCoding = Format(lngDay, "0#") & " " & Format(lngMonth, "0#") & " " & Format(lngYear, "0#") & " " & Format(lngHour, "0#") & " " & Format(lngMinute, "0#") & " "
        ConvDateTimeToThiesCoding = Chr(lngDay) & Chr(lngMonth) & Chr(lngYear) & Chr(lngHour) & Chr(lngMinute)
    End Function
    Public m_DB As New clsDataBase(My.Settings.ICSDBConnectionString, My.Application.Info.DirectoryPath)
    Public Function GetNextDateTimePointerForThies(WMOCODE As String, ClientID As String) As String
        Dim d As DateTime
        'Dim Dur As Date
        'Dur = "00:01"
        Try
            Dim LastSampleDateTime As String = m_DB.GetLastSampleDateTimeBYWMOCode(WMOCODE)

            'd = GetLastReadOutRecordFromThies()
            d = CDate(LastSampleDateTime)
            '''''''d = CDate("2016-05-09 09:00:00.000")

            WriteLogs("GetLastReadOutRecordFromThies():" & Format(d, "yyyy/MM/dd") & " " & Format(d, "HH:mm"), ClientID)

        Catch ex As Exception
            WriteLogs("d = GetLastReadOutRecordFromThies():" & ex.Message, ClientID)
        End Try
        Try
            d = d.AddMinutes(1)
        Catch ex As Exception
            ' WriteLogs("d = d.AddMinutes(1):" & ex.Message)
        End Try

        GetNextDateTimePointerForThies = ConvDateTimeToThiesCoding(Format(d, "yyyy/MM/dd"), Format(d, "HH:mm"))
        Try
            WriteLogs("2result " & GetNextDateTimePointerForThies, ClientID)
        Catch ex As Exception

        End Try

        'GetNextDateTimePointerForThies = ConvDateTimeToThiesCoding(Format(Now, "yyyy/MM/dd"), "00:00")
    End Function

    'Public Function GetLastReadOutRecordFromThies() As Date
    '    Dim strRegRet As String = ""
    '    Try
    '        'GetLastReadOutRecordFromThies = Format(Now, "yyyy/MM/dd") & " 00:00"
    '        strRegRet = GetLastLogtimeForThiestdl16byWmoCode(m_StationCode) ' CA.GetConfigItem(RegPath, "LASTREADOUT")
    '        If Trim(strRegRet) <> "" Then
    '            GetLastReadOutRecordFromThies = strRegRet
    '        Else
    '            GetLastReadOutRecordFromThies = Format(Now, "yyyy/MM/dd") & " 00:00"
    '        End If
    '    Catch ex As Exception
    '        GetLastReadOutRecordFromThies = Format(Now, "yyyy/MM/dd") & " 00:00"
    '    End Try

    'End Function

    'Public Function GetLastLogtimeForThiestdl16byWmoCode(ByVal WmoCode As String) As String
    '    Dim regkey As RegistryKey
    '    Try
    '        regkey = Registry.LocalMachine.OpenSubKey("Software\Rtu_MainService")
    '        GetLastLogtimeForThiestdl16byWmoCode = regkey.GetValue(WmoCode & "LASTREADOUT", Format(System.DateTime.Now, "yyyy/MM/dd") & " 00:00")

    '    Catch ex As Exception
    '        Writeport(ex.Message, "")
    '        GetLastLogtimeForThiestdl16byWmoCode = Format(System.DateTime.Now, "yyyy/MM/dd") & " 00:00"
    '    End Try


    'End Function

#End Region

    'Private Sub RTU_Final_File_Generated(ByVal ObjectID As Integer, ByVal ContactID As Long, ByVal vendor As String, ByVal memoryactive As Boolean) ' Handles RTUPort.RTU_Final_File_Generated
    '    'Try
    '    '    If vendor = "Thiestdl16" Or vendor = "Thies" Then
    '    '        '  If vendor = "Thiestdl16" Or (Not memoryactive And vendor = "Thies") Then
    '    '        updateReadDataPercentIntblCallInfo(RTUPort(ObjectID).ContactID, -2)
    '    '    End If

    '    '    If Not flagHangingStart Then
    '    '        '   WriteRTULog("Hang 21")
    '    '        RTUPort(ObjectID).Hangup()
    '    '    End If

    '    'Catch ex As Exception
    '    '    WriteRTUError("RTU_Final_File_Generated Error is  " & " : " & ex.Message)
    '    'End Try

    'End Sub
#End Region
    Public Function DeleteGarbageChars(ByVal strIn As String) As String
        DeleteGarbageChars = strIn
        Try


            Dim loc As Long
            loc = strIn.IndexOf("=")
            Return Mid(strIn, loc + 1, strIn.Length - loc)
        Catch ex As Exception

        End Try
    End Function
    Public Function ConvHexToDec(ByVal strHex As String) As Byte
        Try


            Select Case UCase(strHex)
                Case "00"
                    ConvHexToDec = 0
                Case "01"
                    ConvHexToDec = 1
                Case "02"
                    ConvHexToDec = 2
                Case "03"
                    ConvHexToDec = 3
                Case "04"
                    ConvHexToDec = 4
                Case "05"
                    ConvHexToDec = 5
                Case "06"
                    ConvHexToDec = 6
                Case "07"
                    ConvHexToDec = 7
                Case "08"
                    ConvHexToDec = 8
                Case "09"
                    ConvHexToDec = 9
                Case "0A"
                    ConvHexToDec = 10
                Case "0B"
                    ConvHexToDec = 11
                Case "0C"
                    ConvHexToDec = 12
                Case "0D"
                    ConvHexToDec = 13
                Case "0E"
                    ConvHexToDec = 14
                Case "0F"
                    ConvHexToDec = 15
                Case "10"
                    ConvHexToDec = 16
                Case "11"
                    ConvHexToDec = 17
                Case "12"
                    ConvHexToDec = 18
                Case "13"
                    ConvHexToDec = 19
                Case "14"
                    ConvHexToDec = 20
                Case "15"
                    ConvHexToDec = 21
                Case "16"
                    ConvHexToDec = 22
                Case "17"
                    ConvHexToDec = 23
                Case "18"
                    ConvHexToDec = 24
                Case "19"
                    ConvHexToDec = 25
                Case "1A"
                    ConvHexToDec = 26
                Case "1B"
                    ConvHexToDec = 27
                Case "1C"
                    ConvHexToDec = 28
                Case "1D"
                    ConvHexToDec = 29
                Case "1E"
                    ConvHexToDec = 30
                Case "1F"
                    ConvHexToDec = 31
                Case "20"
                    ConvHexToDec = 32
            End Select
        Catch ex As Exception
            '   RaiseEvent RTU_WriteError(m_ObjectID, "ConvHexToDec " & ex.Message)
        End Try
    End Function
    Public Function ConvDecToHex(ByVal Dec As Byte) As String
        Try


            ConvDecToHex = ""
            Select Case UCase(Dec)
                Case 0
                    ConvDecToHex = "00"
                Case 1
                    ConvDecToHex = "01"
                Case 2
                    ConvDecToHex = "02"
                Case 3
                    ConvDecToHex = "03"
                Case 4
                    ConvDecToHex = "04"
                Case 5
                    ConvDecToHex = "05"
                Case 5
                    ConvDecToHex = "06"
                Case 6
                    ConvDecToHex = "07"
                Case 7
                    ConvDecToHex = "08"
                Case 8
                    ConvDecToHex = "09"
                Case 10
                    ConvDecToHex = "0A"
                Case 11
                    ConvDecToHex = "0B"
                Case 12
                    ConvDecToHex = "0C"
                Case 13
                    ConvDecToHex = "0D"
                Case 14
                    ConvDecToHex = "0E"
                Case 15
                    ConvDecToHex = "0F"
                Case 16
                    ConvDecToHex = "10"
                Case 17
                    ConvDecToHex = "11"
                Case 18
                    ConvDecToHex = "12"
                Case 19
                    ConvDecToHex = "13"
                Case 20
                    ConvDecToHex = "14"
                Case 21
                    ConvDecToHex = "15"
                Case 22
                    ConvDecToHex = "16"
                Case 23
                    ConvDecToHex = "17"
                Case 24
                    ConvDecToHex = "18"
                Case 25
                    ConvDecToHex = "19"
                Case 26
                    ConvDecToHex = "1A"
                Case 27
                    ConvDecToHex = "1B"
                Case 28
                    ConvDecToHex = "1C"
                Case 29
                    ConvDecToHex = "1D"
                Case 30
                    ConvDecToHex = "1E"
                Case 31
                    ConvDecToHex = "1F"
                Case 32
                    ConvDecToHex = "20"
            End Select
        Catch ex As Exception
            ' RaiseEvent RTU_WriteError(m_ObjectID, "ConvDecToHex " & ex.Message)
        End Try
    End Function
    Public Sub GenerateLambreshtLogs(ByVal HexFullFileName As String, ByVal FileName As String)
        Try
            m_TempfileCreat = True
            Dim strLog As String
            Dim strTemp As String
            Dim strSpil() As String
            Dim FsTempfile As FileStream = New FileStream(HexFullFileName, FileMode.Open, FileAccess.Read, FileShare.None)
            Dim FsCorrectFile As FileStream = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
            Dim rsread As StreamReader = New StreamReader(FsTempfile)
            Dim rsWrite As StreamWriter = New StreamWriter(FsCorrectFile)
            Try


                Dim str_SID As String
                Dim lng_SC As String
                Dim LoggerType As String = m_LoggerType
                Dim str_Number As String
                str_SID = m_StationID
                lng_SC = m_StationCode
                str_Number = ""
                Dim ISCreatHeader As Boolean = False
                m_Vendor = 3
                LoggerType = m_Vendor.ToString
                Try
                    rsWrite.WriteLine(lng_SC.ToString & Space(30 - lng_SC.ToString.Length) & str_SID & Space(30 - str_SID.Length) & LoggerType & Space(30 - LoggerType.Length))

                Catch ex As Exception

                End Try

                rsWrite.Write("DATE" & Space(30 - "DATE".Length) & "Time" & Space(30 - "Time".Length))
                ' WriteErrorLog("SensorCount= ")


                Try
                    For i As Integer = 0 To m_SensorCount

                        Try


                            If Sensors(i).Name <> Nothing Then
                                If UCase(Sensors(i).Name).Contains("TEMP") And (Not UCase(Sensors(i).Name).Contains("TEMPERATURE")) Then
                                    'Soil Temp 5cm
                                    'Soil Temperature 5cm
                                    Sensors(i).Name = Sensors(i).Name.Replace("Temp", "Temperature")

                                End If
                                'If UCase(Sensors(i).Name) = "TEMPERATURE" Then '1
                                '    Sensors(i).Name = "Air Temperature"
                                'End If
                                'If UCase(Sensors(i).Name) = "HUMIDITY" Then
                                '    Sensors(i).Name = "rel. Humidity"
                                'End If
                                rsWrite.Write(Sensors(i).Name & Space(30 - Sensors(i).Name.Length))
                                '  WriteErrorLog("temp" & Sensors(i).Name)
                            End If
                        Catch ex As Exception
                            WriteLogs("temp 1" & ex.Message, m_ClientID)

                            WriteLogs("temp 1" & Sensors(i).Name.Length, m_ClientID)
                            Sensors(i).Name = "NA"
                            rsWrite.Write(Sensors(i).Name & Space(30 - Sensors(i).Name.Length))
                        End Try
                    Next
                Catch ex As Exception
                    '   WriteErrorLog("temp 2" & ex.Message)
                End Try

                rsWrite.WriteLine("")
                Dim LastLogDateTime As String = ""
                While Not rsread.EndOfStream

                    strLog = rsread.ReadLine
                    If strLog.Contains("[") Then
                    Else

                        strLog = strLog.Replace(Chr(10), "")
                        strLog = strLog.Replace(Chr(13), "")
                        ' strLog = strLog.Replace("M101", "")

                        strTemp = strLog

                        If strTemp <> "" Then
                            'M1 21061204044000;37.96;17.3;394;39.80;40.96;0.9;42.4;994.0;5.1;5.6;5.8;0.00
                            'M1 01 11 11 03 125000;8.75;73.2;3.8;272.4;118;774.6;0.03;-0.007;0.000|
                            '111103125000;8.75;73.2;3.8;272.4;118;774.6;0.03;-0.007;0.000|
                            '   strTemp = strTemp.Replace("M101", "")
                            strSpil = Split(strTemp, ";", -1, vbBinaryCompare)
                            Dim H As Long
                            If LastLogDateTime = strSpil(0) Then

                            Else
                                LastLogDateTime = strSpil(0)
                                strSpil(0) = Right(strSpil(0), 14)
                                '103121302095000
                                rsWrite.Write("20" & Mid(strSpil(0), 5, 2) & "/" & Mid(strSpil(0), 3, 2) & "/" & Mid(strSpil(0), 1, 2) & Space(20))
                                '111103125000;8.75;73.2;3.8;272.4;118;774.6;0.03;-0.007;0.000|
                                rsWrite.Write(Mid(strSpil(0), 9, 2) & ":" & Mid(strSpil(0), 11, 2) & ":" & Mid(strSpil(0), 13, 2) & Space(22))
                                For H = 1 To UBound(strSpil)
                                    strTemp = strSpil(H)
                                    '  If m_Vendor = Logger_Vendor.Partonegar Then
                                    rsWrite.Write(Math.Round(Val(strTemp), 1) & Space(30 - Len(Trim(Math.Round(Val(strTemp), 1)))))
                                    '

                                Next H
                                rsWrite.WriteLine("")
                                ' Print #1, ""
                                Application.DoEvents()
                            End If
                        End If
                    End If

                End While
            Catch ex As Exception
                WriteLogs("GenerateLambreshtLogs 0" & ex.Message, m_ClientID)
            End Try
            rsWrite.Flush()
            rsWrite.Close()
            FsCorrectFile.Dispose()
            FsTempfile.Dispose()
            rsread.Close()
            rsread.Dispose()
        Catch ex As Exception
            '  RaiseEvent RTU_WriteError(m_ObjectID, "GenerateLambreshtLogs" & ex.Message)
            WriteLogs("GenerateLambreshtLogs 1" & ex.Message, m_ClientID)
        End Try
    End Sub
    Public Function Temp2RealValue(ByVal HexFullFileName As String, ByVal FileName As String) As Long

        Try



            m_TempfileCreat = True

            Temp2RealValue = 0
            WriteLogs(m_Vendor.ToString, m_ClientID)
            If m_Vendor = Logger_Vendor.Vaisala Then
                '  WriteErrorLog("Vaisala")

                '    GenerateVisalaLogs(HexFullFileName, FileName)
            ElseIf m_Vendor = Logger_Vendor.Lambresht Then
                WriteLogs("Lambresht", m_ClientID)
                Dim t As New Threading.Thread(DirectCast(Sub() GenerateLambreshtLogs(HexFullFileName, FileName), Threading.ThreadStart))
                t.Start()
                t.Join()


                'ElseIf m_Vendor = Logger_Vendor.Thiestdl16 Then ' Or (m_Vendor = Logger_Vendor.Thies And Not m_Memory_Active) Then
                '    'Writeport("Generate file start", m_RTUState.ToString, StationName, 0)
                '    GenerateThiesLogsfromDatalogger(HexFullFileName, FileName)
                '    'Writeport("Generate file end", m_RTUState.ToString, StationName, 0)
            ElseIf m_Vendor = Logger_Vendor.Thies Then
                WriteLogs("Generate file start", m_ClientID)
                GenerateThiesLogsfromDatalogger(HexFullFileName, FileName)
                WriteLogs("Generate file End", m_ClientID)
                'test   GenerateThiesLogs(HexFullFileName, FileName)
            ElseIf m_Vendor = Logger_Vendor.Vaisala Then
                '  WriteErrorLog("Vaisala")

                GenerateVisalaLogs(HexFullFileName, FileName)
            Else
                Dim t As New Threading.Thread(DirectCast(Sub() GeneratetheodorLogs(HexFullFileName, FileName), Threading.ThreadStart))
                t.Start()
                t.Join()
                '  GeneratetheodorLogs(HexFullFileName, FileName)
            End If


        Catch ex As Exception

        End Try
    End Function
    Public Sub GenerateVisalaLogs(ByVal HexFullFileName As String, ByVal FileName As String)
        Try


            Dim strLog As String
            Dim strTemp As String
            Dim strSpil() As String
            Dim FsTempfile As FileStream = New FileStream(HexFullFileName, FileMode.Open, FileAccess.Read, FileShare.None)
            Dim FsCorrectFile As FileStream = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
            Dim rsread As StreamReader = New StreamReader(FsTempfile)
            Dim rsWrite As StreamWriter = New StreamWriter(FsCorrectFile)
            Try
                Dim str_SID As String
                Dim lng_SC As String
                Dim LoggerType As String = m_LoggerType
                '    Dim str_Number As String
                str_SID = m_StationID
                lng_SC = m_StationCode
                '   str_Number = m_Dial_Number
                Dim ISCreatHeader As Boolean = False
                ' WriteErrorLog("Vaisala")
                Try
                    rsWrite.WriteLine(lng_SC.ToString & Space(25 - lng_SC.ToString.Length) & str_SID & Space(25 - str_SID.Length) & "Vaisala" & Space(25 - "Vaisala".Length))
                Catch ex As Exception
                    ' ArrayLog(0, 0) = ""
                End Try
                rsWrite.Write("DATE" & Space(25 - "DATE".Length) & "Time" & Space(25 - "Time".Length))
                While Not rsread.EndOfStream
                    strLog = rsread.ReadLine
                    While Not strLog.Contains("[") And Not rsread.EndOfStream
                        strLog = strLog & rsread.ReadLine
                    End While
                    Dim strSplit() = Split(strLog, "[")
                    strLog = strSplit(0)
                    strTemp = strLog

                    If strTemp <> "" Then
                        strSpil = Split(strTemp, ";", -1, vbBinaryCompare)
                        If Not ISCreatHeader Then
                            Try
                                For i As Integer = 3 To strSpil.Length - 1
                                    Dim SensorName()
                                    If strSpil(i).Contains(":") Then
                                        SensorName = Split(strSpil(i), ":")
                                        If SensorName(0) <> Nothing Then
                                            rsWrite.Write(GetVisalasensorName(SensorName(0)) & Space(25 - GetVisalasensorName(SensorName(0)).Length))
                                            Sensors(i - 3).Name = GetVisalasensorName(SensorName(0))
                                        End If
                                    Else
                                        SensorName = Split(strSpil(i), "10m")
                                        If SensorName(0) <> Nothing Then
                                            rsWrite.Write(GetVisalasensorName(SensorName(0)) & Space(25 - GetVisalasensorName(SensorName(0)).Length))
                                            Sensors(i - 3).Name = GetVisalasensorName(SensorName(0))

                                        End If
                                    End If
                                Next
                            Catch ex As Exception

                            End Try
                            ISCreatHeader = True
                            rsWrite.WriteLine("")
                        End If
                        Dim H As Long
                        Dim StrDate As String = Mid(strSpil(1), 3, strSpil(1).Length - 1)
                        Dim StrTime As String = Mid(strSpil(2), 3, strSpil(2).Length - 1)
                        StrDate = StrDate.Replace(":", "")
                        StrTime = StrTime.Replace(":", "")
                        rsWrite.Write("20" & Mid(StrDate, 1, 2) & "/" & Mid(StrDate, 3, 2) & "/" & Mid(StrDate, 5, 2) & Space(25 - Len(20 & Mid(StrDate, 1, 2) & "/" & Mid(StrDate, 3, 2) & "/" & Mid(StrDate, 5, 2))))
                        rsWrite.Write(Mid(StrTime, 1, 2) & ":" & Mid(StrTime, 3, 2) & ":" & Mid(StrTime, 5, 2) & Space(25 - Len(Mid(StrTime, 1, 2) & ":" & Mid(StrTime, 3, 2) & ":" & Mid(StrTime, 5, 2))))
                        For H = 3 To strSpil.Length - 1
                            'SM410MA:0.0;
                            'SM510MA:0.2;
                            'SM610MA:0.1
                            Dim Sensor()
                            If strSpil(H).Contains(":") Then
                                Try
                                    Sensor = Split(strSpil(H), ":")
                                    strTemp = Sensor(1)
                                Catch ex As Exception
                                    strTemp = 9999
                                End Try
                            Else
                                Sensor = Split(strSpil(H), "10m")
                                strTemp = Sensor(1)
                            End If
                            rsWrite.Write(Math.Round(Val(strTemp), 1) & Space(25 - Len(Trim(Math.Round(Val(strTemp), 1)))))
                        Next H
                        rsWrite.WriteLine("")
                        Application.DoEvents()
                    End If

                End While
            Catch ex As Exception
                WriteLogs("GenerateVisalaLogs 0" & ex.Message, m_ClientID)
            End Try
            rsWrite.Flush()
            rsWrite.Close()
            FsCorrectFile.Dispose()
            FsTempfile.Dispose()
            rsread.Close()
            rsread.Dispose()
        Catch ex As Exception
            'RaiseEvent RTU_WriteError(m_ObjectID, "GenerateVisalaLogs " & ex.Message)
            WriteLogs("GenerateVisalaLogs 1" & ex.Message, m_ClientID)
        End Try
    End Sub
    Public Function GetVisalasensorName(ByVal Header As String) As String
        Try


            Select Case UCase(Header)
                Case "WD10MA", "QuickWind_1.DirAvg"
                    GetVisalasensorName = "Wind Direction"
                Case "WS10MA"
                    GetVisalasensorName = "Wind Speed"
                Case "TA10MA"
                    GetVisalasensorName = "Air Temperature"
                Case "RH10MA"
                    GetVisalasensorName = "Relative Humidity"
                Case "DP10MA"
                    GetVisalasensorName = "Dewpoint"
                Case "WT10MA"
                    GetVisalasensorName = "Wet Temperature"
                Case "QFE10MA", "QFE10IA"
                    GetVisalasensorName = "Air Pressure"
                Case "QFF10MA"
                    GetVisalasensorName = "QFF"
                Case "PRSUM"
                    GetVisalasensorName = "Precipitation-Sum"
                Case "SR10MA"
                    GetVisalasensorName = "Radiation"
                Case "TG110MA"
                    GetVisalasensorName = "Surface Temperature" '"Soil Temperature(1)" 'SURFACE TEMPERATURE
                Case "TG210MA"
                    GetVisalasensorName = "Soil Temperature(2)"
                Case "TG310MA"
                    GetVisalasensorName = "Soil Temperature(3)"
                Case "TG410MA"
                    GetVisalasensorName = "Soil Temperature(4)"
                Case "TG510MA"
                    GetVisalasensorName = "Soil Temperature(5)"
                Case "TG610MA"
                    GetVisalasensorName = "Soil Temperature(6)"
                Case "TG710MA"
                    GetVisalasensorName = "Soil Temperature(7)"
                Case "SM110MA"
                    GetVisalasensorName = "Soil Moisture1"
                Case "SM210MA"
                    GetVisalasensorName = "Soil Moisture2"
                Case "SM310MA"
                    GetVisalasensorName = "Soil Moisture3"
                Case "SM410MA"
                    GetVisalasensorName = "Soil Moisture4"
                Case "SM510MA"
                    GetVisalasensorName = "Soil Moisture5"
                Case "SM610MA"
                    GetVisalasensorName = "Soil Moisture6"
                Case "SM710MA"
                    GetVisalasensorName = "Soil Moisture7"
                Case "SRSUM"
                    GetVisalasensorName = "Radiation-SUM"
                Case "SDSUM"
                    GetVisalasensorName = "Sun-Shine"
                Case "WH10MA"
                    GetVisalasensorName = "NA"
                Case "EVSUM"
                    GetVisalasensorName = "Evaporation"
                Case "TW10MA"
                    GetVisalasensorName = "NA"
                Case "WINDRUN"
                    GetVisalasensorName = "WIND-RUN"
                Case "QNH10MA"
                    GetVisalasensorName = "QNH"
                    'Case "QFE"
                    '    GetVisalasensorName = "Air Pressure"
                Case Else
                    GetVisalasensorName = Header


            End Select
        Catch ex As Exception
            '     RaiseEvent RTU_WriteError(m_ObjectID, "GetVisalasensorName " & ex.Message)
            WriteLogs("GetVisalasensorName 1" & ex.Message, m_ClientID)
        End Try
    End Function

    Public Sub GenerateThiesLogsfromDatalogger(ByVal HexFullFileName As String, ByVal FileName As String)
        Try


            m_TempfileCreat = True
            Dim strLog As String
            Dim strTemp As String

            Dim FsTempfile As FileStream = New FileStream(HexFullFileName, FileMode.Open, FileAccess.Read, FileShare.None)
            Dim FsCorrectFile As FileStream = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
            Dim rsread As StreamReader = New StreamReader(FsTempfile)
            Dim rsWrite As StreamWriter = New StreamWriter(FsCorrectFile)
            Try


                Dim str_SID As String
                Dim lng_SC As String
                Dim LoggerType As String = m_LoggerType
                Dim str_Number As String
                str_SID = m_StationID
                lng_SC = m_StationCode
                '  str_Number = m_Dial_Number
                Dim ISCreatHeader As Boolean = False


                Try
                    rsWrite.WriteLine(lng_SC.ToString & Space(30 - lng_SC.ToString.Length) & str_SID & Space(30 - str_SID.Length) & "THIES" & Space(30 - "THIES".Length))
                Catch ex As Exception

                End Try

                rsWrite.Write("DATE" & Space(30 - "DATE".Length) & "Time" & Space(30 - "Time".Length))
                ' Writeport("Sensor count   " & SensorCount, m_RTUState.ToString, "test", 0)
                Try
                    For i As Integer = 0 To Sensors.Length - 1
                        If Sensors(i).Name <> Nothing Then 'And Not UCase(Sensors(i).Name).Contains("NA") Then
                            Try
                                If UCase(Sensors(i).Name).Contains("TEMP") And (Not UCase(Sensors(i).Name).Contains("TEMPERATURE")) Then
                                    'Soil Temp 20cm
                                    Sensors(i).Name = Trim(Sensors(i).Name.Replace("Temp", "Temperature"))

                                End If
                            Catch ex As Exception
                                WriteLogs("1  " & ex.Message, m_ClientID)
                            End Try
                            '
                            Try
                                rsWrite.Write(Sensors(i).Name & Space(30 - Sensors(i).Name.Length))

                                WriteLogs(Sensors(i).Name & Space(30 - Sensors(i).Name.Length) & "end name", m_ClientID)
                            Catch ex As Exception
                                WriteLogs("2  " & ex.Message, m_ClientID)
                                WriteLogs("21  " & "Sensors(i).Name  " & Sensors(i).Name & "  ", m_ClientID)
                                WriteLogs("21  " & "Sensors(i).Name.len  " & Sensors(i).Name.Length & "  ", m_ClientID)
                            End Try

                            '  Writevisalatest(Sensors(i).Name & vbCrLf)
                        End If
                    Next
                Catch ex As Exception
                    WriteLogs("3  " & ex.Message, m_ClientID)
                End Try
                rsWrite.WriteLine("")





                While Not rsread.EndOfStream
                    Try


                        Dim ArrayValue As String() = New String(0) {}
                        strLog = ""
                        'While Not strLog.Contains("[") And Not rsread.EndOfStream
                        '    strLog = strLog & rsread.ReadLine & vbCrLf
                        'End While

                        strLog = rsread.ReadLine
                        If strLog.Contains("Data") Then
                            strLog = rsread.ReadLine
                        End If
                        If strLog.Contains("END OF DATA") Then
                            rsWrite.Flush()
                            rsWrite.Close()
                            FsCorrectFile.Dispose()
                            FsTempfile.Dispose()
                            rsread.Close()
                            rsread.Dispose()
                            Exit Sub
                        End If
                        GetSensorValueListForThies(ArrayValue, strLog)
                        strTemp = strLog

                        Dim StrDate As String = ArrayValue(ArrayValue.Length - 2)
                        Dim StrTime As String = ArrayValue(ArrayValue.Length - 1)

                        Dim strsplit() = Split(StrDate, ".")
                        StrDate = "20" & strsplit(2) & "/" & strsplit(1) & "/" & strsplit(0)
                        rsWrite.Write(StrDate & Space(30 - StrDate.Length))
                        rsWrite.Write(StrTime & Space(30 - StrTime.Length))
                        For H = 1 To ArrayValue.Length - 3
                            Try
                                If IsNumeric(ArrayValue(H)) Then
                                    rsWrite.Write(Math.Round(CDbl(ArrayValue(H)), 1) & Space(30 - Len(Trim(Math.Round(CDbl(ArrayValue(H)), 1)))))
                                Else
                                    rsWrite.Write("99999" & Space(30 - "99999".Length))
                                End If
                            Catch ex As Exception
                                rsWrite.Write("99999" & Space(30 - "99999".Length))
                            End Try
                            'If ArrayValue(H).Contains("?") Then
                            'Else
                            '    '  If Not UCase(Sensors(H).Name).Contains("NA") Then
                            '    ' rsWrite.Write(ArrayValue(H) & Space(30 - Len(Trim(ArrayValue(H)))))
                            '    '   End If
                            'End If
                        Next H
                        ' rowindex = rowindex + 1
                        rsWrite.WriteLine("")
                        ' Print #1, ""
                        Application.DoEvents()
                    Catch ex As Exception

                    End Try
                End While
            Catch ex As Exception
                WriteLogs("GenerateThiesLogs 0 " & ex.Message, m_ClientID)
            End Try
            rsWrite.Flush()
            rsWrite.Close()
            FsCorrectFile.Dispose()
            FsTempfile.Dispose()
            rsread.Close()
            rsread.Dispose()



        Catch ex As Exception
            WriteLogs("GenerateThiesLogs 1 " & ex.Message, m_ClientID)
        End Try





    End Sub
    Public Sub GetSensorValueListForThies(ByRef ArrayValue As String(), ByVal StrTemp As String)
        Dim index As Integer = 0
        Dim value As String = ""
        For i = 0 To StrTemp.Length - 1
            Try
                If StrTemp(i) = Chr(32) Then
                    ReDim Preserve ArrayValue(index)
                    ArrayValue(index) = value
                    index = index + 1
                    value = ""
                    While StrTemp(i) = Chr(32) And i <= StrTemp.Length - 1
                        i = i + 1
                    End While
                    If StrTemp(i) <> Chr(32) Then
                        i = i - 1
                    End If
                Else
                    value = value & StrTemp(i)
                End If
            Catch ex As Exception

            End Try

        Next
        ReDim Preserve ArrayValue(index)
        ArrayValue(index) = value
    End Sub
    Public Sub GeneratetheodorLogs(ByVal HexFullFileName As String, ByVal FileName As String)

        Try
            m_TempfileCreat = True
            Dim strLog As String
            Dim strTemp As String
            Dim strSpil() As String
            Dim FsTempfile As FileStream = New FileStream(HexFullFileName, FileMode.Open, FileAccess.Read, FileShare.None)
            Dim FsCorrectFile As FileStream = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
            Dim rsread As StreamReader = New StreamReader(FsTempfile)
            Dim rsWrite As StreamWriter = New StreamWriter(FsCorrectFile)

            Try


                Dim str_SID As String
                Dim lng_SC As String
                Dim LoggerType As String = m_Vendor.ToString
                Dim str_Number As String
                str_SID = m_StationID
                lng_SC = m_StationCode
                Dim ISCreatHeader As Boolean = False

                Try
                    rsWrite.WriteLine(lng_SC.ToString & Space(25 - lng_SC.ToString.Length) & str_SID & Space(25 - str_SID.Length) & LoggerType & Space(25 - LoggerType.Length))

                Catch ex As Exception

                End Try

                rsWrite.Write("DATE" & Space(25 - "DATE".Length) & "Time" & Space(25 - "Time".Length))

                Try
                    For i As Integer = 0 To Sensors.Length - 1
                        Try
                            If Sensors(i).Name <> Nothing Then
                                If UCase(Sensors(i).Name).Contains("TEMP") And (Not UCase(Sensors(i).Name).Contains("TEMPERATURE")) Then
                                    Sensors(i).Name = Sensors(i).Name.Replace("Temp", "Temperature")
                                End If
                                Sensors(i).Name = Sensors(i).Name.Replace(Chr(13), "")
                                rsWrite.Write(Sensors(i).Name & Space(25 - Sensors(i).Name.Length))
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                Catch ex As Exception

                End Try

                rsWrite.WriteLine("")
                While Not rsread.EndOfStream

                    strLog = rsread.ReadLine
                    strLog = strLog.Replace(Chr(10), "")
                    strLog = strLog.Replace(Chr(13), "")
                    strLog = strLog.Replace("=1", "")
                    If Trim(strLog.Length) > 17 Then
                        strTemp = strLog
                        If Asc(Mid(strLog, 1, 1)) <> 0 Then
                            If strTemp <> "" Then
                                strSpil = Split(strTemp, ";", -1, vbBinaryCompare)
                                Dim H As Long
                                rsWrite.Write("20" & Mid(strSpil(0), 1, 2) & "/" & Mid(strSpil(0), 3, 2) & "/" & Mid(strSpil(0), 5, 2) & Space(25 - Len(20 & Mid(strSpil(0), 1, 2) & "/" & Mid(strSpil(0), 3, 2) & "/" & Mid(strSpil(0), 5, 2))))
                                rsWrite.Write(Mid(strSpil(0), 7, 2) & ":" & Mid(strSpil(0), 9, 2) & ":" & Mid(strSpil(0), 11, 2) & Space(25 - Len(Mid(strSpil(0), 7, 2) & ":" & Mid(strSpil(0), 9, 2) & ":" & Mid(strSpil(0), 11, 2))))
                                For H = 1 To UBound(strSpil) - 1
                                    strTemp = strSpil(H)
                                    If m_Vendor = Logger_Vendor.Partonegar Then
                                        rsWrite.Write(Math.Round(Val(strTemp), 1) & Space(25 - Len(Trim(Math.Round(Val(strTemp), 1)))))
                                    Else
                                        rsWrite.Write(Math.Round(HexStringToSingle(strTemp), 1) & Space(25 - Len(Trim(Math.Round(HexStringToSingle(strTemp), 1)))))

                                    End If

                                Next H
                                rsWrite.WriteLine("")
                                ' Print #1, ""
                                Application.DoEvents()
                            End If
                        End If
                    End If
                End While
            Catch ex As Exception

            End Try
            rsWrite.Flush()
            rsWrite.Close()
            FsCorrectFile.Dispose()
            FsTempfile.Dispose()
            rsread.Close()
            rsread.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Public Function HexStringToSingle(ByVal strHex As String) As Single
        Try




            Dim strTemp As String = ""
            Dim i As Long
            Dim Sign As Long
            Dim strExponent As String
            Dim strMantisa As String
            Dim lngExponent As Long
            Dim sglMantisa As Single
            Dim strBin As String = ""
            For i = 1 To Len(strHex)
                strBin = strBin & HexNibleToBin(Asc(Mid(strHex, i, 1)))
            Next i
            If Mid(strBin, 1, 1) = 0 Then
                Sign = 1
            Else
                Sign = -1
            End If
            strExponent = Mid(strBin, 2, 8)
            strMantisa = Mid(strBin, 10, 23)
            lngExponent = ByteBinToUnsignedDec(strExponent) - 127
            sglMantisa = 1 + FracByteBinToUnsignedDec(strMantisa)

            If lngExponent <> -127 Then
                HexStringToSingle = Sign * 2 ^ (lngExponent) * (sglMantisa)
            Else
                If sglMantisa = 1 Then
                    HexStringToSingle = 0
                Else
                    HexStringToSingle = Sign * 2 ^ (lngExponent + 1) * (sglMantisa - 1)
                End If
            End If
        Catch ex As Exception
            '   RaiseEvent RTU_WriteError(m_ObjectID, "HexStringToSingle" & ex.Message)
        End Try
    End Function
    Public Function HexNibleToBin(ByVal intHex As Integer) As String
        Try


            Dim strTemp As String = ""
            intHex = intHex
            If intHex = 48 Then
                strTemp = "0000"
            ElseIf intHex = 49 Then
                strTemp = "0001"
            ElseIf intHex = 50 Then
                strTemp = "0010"
            ElseIf intHex = 51 Then
                strTemp = "0011"
            ElseIf intHex = 52 Then
                strTemp = "0100"
            ElseIf intHex = 53 Then
                strTemp = "0101"
            ElseIf intHex = 54 Then
                strTemp = "0110"
            ElseIf intHex = 55 Then
                strTemp = "0111"
            ElseIf intHex = 56 Then
                strTemp = "1000"
            ElseIf intHex = 57 Then
                strTemp = "1001"
            ElseIf intHex = 65 Then
                strTemp = "1010"
            ElseIf intHex = 66 Then
                strTemp = "1011"
            ElseIf intHex = 67 Then
                strTemp = "1100"
            ElseIf intHex = 68 Then
                strTemp = "1101"
            ElseIf intHex = 69 Then
                strTemp = "1110"
            ElseIf intHex = 70 Then
                strTemp = "1111"
            ElseIf intHex = 97 Then
                strTemp = "1010"
            ElseIf intHex = 98 Then
                strTemp = "1011"
            ElseIf intHex = 99 Then
                strTemp = "1100"
            ElseIf intHex = 100 Then
                strTemp = "1101"
            ElseIf intHex = 101 Then
                strTemp = "1110"
            ElseIf intHex = 102 Then
                strTemp = "1111"
            End If
            HexNibleToBin = strTemp
        Catch ex As Exception
            '  RaiseEvent RTU_WriteError(m_ObjectID, "HexNibleToBin" & ex.Message)
        End Try
    End Function

    Public Function ByteBinToUnsignedDec(ByVal strBits As String) As Byte
        Try


            Dim W As Byte
            Dim i As Byte
            Dim strBin As String
            Dim bytRet As Byte
            strBin = strBits
            bytRet = 0
            W = 128
            For i = 1 To 8
                bytRet = bytRet + W * Val(Mid(strBin, i, 1))
                W = W / 2
            Next i
            ByteBinToUnsignedDec = bytRet
        Catch ex As Exception
            '   RaiseEvent RTU_WriteError(m_ObjectID, "ByteBinToUnsignedDec" & ex.Message)
        End Try
    End Function

    Public Function FracByteBinToUnsignedDec(ByVal strFracBin As String) As Single
        Try


            Dim W As Single
            Dim i As Byte
            Dim sglRet As Single
            sglRet = 0
            W = 0.5
            For i = 1 To 23
                sglRet = sglRet + W * Val(Mid(strFracBin, i, 1))
                W = W / 2
            Next i
            FracByteBinToUnsignedDec = sglRet
        Catch ex As Exception
            '      RaiseEvent RTU_WriteError(m_ObjectID, "FracByteBinToUnsignedDec" & ex.Message)
        End Try
    End Function

    Private Function Server() As Object
        Throw New NotImplementedException
    End Function

End Class
