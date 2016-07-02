Imports Microsoft.Win32
Public Class clsConfigAccess
    Public Function GetConfigItem(ByVal KeyAddress As String, ByVal item As String) As Object
        Dim key As RegistryKey = Registry.LocalMachine
        Dim subkey As RegistryKey
        Try
            Select Case UCase(item)
                Case "POLLING_INTERVAL"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CLng(subkey.GetValue(item))
                        If GetConfigItem = 0 Then
                            GetConfigItem = 30000
                        End If
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = 30000
                    End Try                
                Case "PROVINCE_ICAOID"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "????"
                    End Try
                Case "SOURCE"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "D:\SOURCE\"
                    End Try
                Case "BACKUP"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "D:\BACKUP\"
                    End Try
                Case "DESTINATION"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "D:\DESTINATION\"
                    End Try
                Case "GARBAGE"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "D:\GARBAGE\"
                    End Try
                Case "ICSDBCONNECTIONSTRING"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "Data Source=.;Initial Catalog=NRCDB;User ID=sa;Password=0912admin12!@"
                    End Try
                Case "COMPORT"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "COM1"
                    End Try
                Case "RTSENABLED"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = False
                    End Try
                Case "BAUDRATE"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "9600"
                    End Try
                Case "STATIONNAME"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "StationName"
                    End Try
                Case "WMOCODE"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "WMOCode"
                    End Try
                Case "VENDOR"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "Vendor"
                    End Try
                Case "IKAOCODE", "IKAOID"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue("IKAOID")
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "IKAOID"
                    End Try
                Case "STATIONTYPE"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "SYNOPTIC"
                    End Try
                Case "SENDERLOGGINGENABLED"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = True
                    End Try
                Case "ACK_REC_TO_ICSLOGGINGENABLED"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = True
                    End Try
                Case "LOGGINGPATH"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "C:\"
                    End Try
                Case "LOCALLOGS"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "D:\Logs\"
                    End Try
                Case "MODULE_COUNT"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CByte(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = 1
                    End Try
                Case "ONLINESAMPLINGINTERVAL"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CLng(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = 1000
                    End Try
                Case "DATABITS"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "8"
                    End Try
                Case "PARITY"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "N"
                    End Try
                Case "HANDSHAKING"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "NONE"
                    End Try
                Case "PROTOCOLID"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "1"
                    End Try
                Case "LOGREADINGENABLED"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = True
                    End Try
                Case "SERVERIP"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "127.0.0.1"
                    End Try
                Case "SERVERPORT"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "8093"
                    End Try
                Case "LASTRDOSRC"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "DB"
                    End Try
                Case "SYNCH_DCS"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = True
                    End Try
                Case "SYNCHDCSINTERVAL"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CLng(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = 3
                    End Try
                Case "FTPALIASNAME"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "/LOGS/"
                    End Try
                Case "FTPUSER"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "FTPUSER"
                    End Try
                Case "FTPPASSWORD"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = "FTPPASSWORD"
                    End Try
                Case "SENDER_MAINTASK_INTERVAL"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CLng(subkey.GetValue(item))
                        If GetConfigItem = 0 Then
                            GetConfigItem = 10000
                        End If
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = 10000
                    End Try
                Case "AUTOSETLOGGERCLOCK"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = False
                    End Try
                Case "LASTREADOUT"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = False
                    End Try
                Case "HASEXTREMEVALUE"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = subkey.GetValue(item)
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = False
                    End Try
                Case "FORMULALOGSHOW"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = False
                    End Try
                Case "READSYNOPS"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = True
                    End Try
                Case "READMETARS"
                    Try
                        subkey = key.OpenSubKey(KeyAddress, True)
                        GetConfigItem = CBool(subkey.GetValue(item))
                        subkey.Close()
                    Catch ex As Exception
                        GetConfigItem = True
                    End Try
                Case Else
                    GetConfigItem = ""

            End Select

        Catch ex As Exception
            GetConfigItem = "???"
        Finally
            key.Close()
        End Try
    End Function
End Class

'Imports Microsoft.Win32
'Public Class clsConfigAccess   
'    Public Function GetConfigItem(ByVal KeyAddress As String, ByVal item As String) As Object
'        Dim key As RegistryKey = Registry.LocalMachine
'        Dim subkey As RegistryKey
'        Try


'            Select Case UCase(item)

'                Case "ICSDBCONNECTIONSTRING"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "Data Source=.;Initial Catalog=ICSDB;User ID=sa;Password=admin11"
'                    End Try

'                Case "RUNFIRSTTIME"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = CBool(subkey.GetValue(item))
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = True
'                    End Try


'                Case "SERVICELOGGINGENABLED"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = CBool(subkey.GetValue(item))
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = True
'                    End Try
'                Case "LOGGINGPATH"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "C:\AWSLogs\"
'                    End Try


'                Case "LOGREADINGENABLED"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = CBool(subkey.GetValue(item))
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = True
'                    End Try
'                Case "SERVERIP"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "127.0.0.1"
'                    End Try
'                Case "SERVERPORT"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "5500"
'                    End Try

'                Case "USER"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "USER"
'                    End Try
'                Case "PASSWORD"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "PASSWORD"
'                    End Try
'                Case "DESTINATION"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "c:\"
'                    End Try
'                Case "ICSFEED"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = subkey.GetValue(item)
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = "c:\"
'                    End Try
'                Case "SENDER_MAINTASK_INTERVAL"
'                    Try
'                        subkey = key.OpenSubKey(KeyAddress, True)
'                        GetConfigItem = CLng(subkey.GetValue(item))
'                        If GetConfigItem = 0 Then
'                            GetConfigItem = 10000
'                        End If
'                        subkey.Close()
'                    Catch ex As Exception
'                        GetConfigItem = 10000
'                    End Try
'                    'Sender_MainTask_Interval
'                Case Else
'                    GetConfigItem = ""
'            End Select
'        Catch ex As Exception
'            GetConfigItem = "???"
'        Finally
'            key.Close()
'        End Try
'    End Function
'End Class
