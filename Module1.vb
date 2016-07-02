Imports Microsoft.Win32
Imports System.Windows.Forms

Module Module1
    Public LoggerType(100) As clsLogger
    Public Station_Index As Integer = 0
    Public regKey As RegistryKey

    Public Function GetReceiveFolder() As String
        'Dim strRecAdd As String = ""
        'GetReceiveFolder = ""
        'Try
        '    regKey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\IRIMO-ICS-FT")
        '    If Not IsNothing(regKey) Then
        '        strRecAdd = regKey.GetValue("ReceiveFolder", "")
        '        If strRecAdd = "" Then
        '            GetReceiveFolder = GetAppPath()
        '        Else
        '            GetReceiveFolder = strRecAdd
        '        End If
        '    Else
        '        GetReceiveFolder = GetAppPath()
        '    End If
        'Catch ex As Exception
        'End Tryd
        GetReceiveFolder = "D:\Logs"
    End Function
    Public Sub Remove(ByVal ObjectID As Integer, ByVal ClientID As Integer)
        Try


            For i = 0 To LoggerType.Length - 1
                If LoggerType(i).ClientID = ClientID And LoggerType(i).ObjectID <> i Then
                    LoggerType(i).ClientID = -1
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub
    Public Function GetAppPath() As String
        Dim strSpil() As String
        strSpil = Split(Application.ExecutablePath, "\", -1, CompareMethod.Binary)
        GetAppPath = Mid(Application.ExecutablePath, 1, Application.ExecutablePath.Length - strSpil(UBound(strSpil)).Length)
    End Function
    Public LastObject_ID As Integer = 0
    Public Function GetStation_IndexByStationName(ByVal StationName As String) As Integer

        GetStation_IndexByStationName = -1
        Try


            For i = 0 To 100

                If LoggerType(i).StationName = StationName Then
                    GetStation_IndexByStationName = i
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Public Function GetStation_IndexByClientID(ByVal ClientID As String) As Integer

        GetStation_IndexByClientID = -1
        Try


            For i = 0 To 100

                If LoggerType(i).ClientID = ClientID Then
                    GetStation_IndexByClientID = LoggerType(i).ObjectID
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

End Module
