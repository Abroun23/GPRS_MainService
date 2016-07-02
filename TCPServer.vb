Imports System.Net, System.Net.Sockets

Imports System.Text.ASCIIEncoding
Imports System.IO
Public Class TCPServer

#Region "Filed"
    '  Public flagParsReport As Boolean = False
    'Dim RecbufferSize As Integer = 1023
    Dim tcpListener As Socket = Nothing 'New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    Dim MaximonClientCount As Integer = 0
    Dim Connection_Error As String = 0
    Dim _Port As String
    '  Dim ArrayConnectedClient As New List(Of Integer)
    Dim appADD As String = My.Application.Info.DirectoryPath
    Dim lastIndex As Integer = 0

    Public Structure Client_Type
        Public ID As Integer
        Public Ip As String
        Public Recbytes() As Byte '= New Byte(1023) {}
        Public CSocket As Socket
        Public ISConnect As Boolean
        Public _WMOCODE As String
        '    Dim _flagReadReport As Boolean
        Public _flagDisconnecting As Boolean
    End Structure
    Dim Clients(100) As Client_Type
#End Region
#Region "Body"
    Public Sub New(ByVal Port As String, ByVal MaxClient As Integer, ByVal MaxPending As Integer)
        Try


            ReDim Clients(MaxClient)
            'ReDim Preserve Clients(MaimonClientCount)
            'For i = 0 To MaxClient - 1
            '    Try
            '        Dim Clients(i) As Client_Type
            '        '  m_Client.Recbytes = New Byte(10000) {}
            '        'm_Client.ISConnect = false
            '        '     m_Client.CSocket = TcpListener.EndAccept(ar)
            '        'm_Client.Ip = ip.Address.ToString
            '        'm_Client._WMOCODE = ""
            '        ''   m_Client._flagReadReport = False
            '        'm_Client._flagDisconnecting = False
            '        'Clients(ID) = m_Client
            '        '  Clients(i) = New Client_Type()

            '    Catch ex As Exception

            '    End Try


            'Next


            Dim ip As New IPEndPoint(IPAddress.Any, Port)
            _Port = Port
            MaximonClientCount = MaxClient
            tcpListener = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            tcpListener.Bind(ip)
            tcpListener.Listen(MaxPending)
            tcpListener.BeginAccept(New AsyncCallback(AddressOf OnAccept), tcpListener)
        Catch ex As Exception

        End Try
    End Sub
    Private Function GetClientID() As Integer
        '     RaiseEvent SocketError(0, "lastindex is   " & lastIndex, False)
        'If lastIndex >= MaximonClientCount Then
        '    lastIndex = 0
        'End If
        '  RaiseEvent SocketError(0, "lastindex is  " & lastIndex, False)

        'For i = lastIndex To MaximonClientCount
        GetClientID = -1
        For i = 0 To MaximonClientCount
            Try

                If Clients(i).ISConnect = False Then
                    GetClientID = i
                    Exit Function
                Else
                    Dim Status As Boolean
                    If Not CheckClient(i) Then
                        Clients(i).ISConnect = False
                        RaiseEvent SocketError(i, "CheckClient():" & Status.ToString, False)

                    End If


                End If
            Catch ex As Exception
                RaiseEvent SocketError(0, "GetClientID() error:" & ex.Message, False)
                GetClientID = i
                '  lastIndex = i + 
                Exit Function

            End Try
        Next i

    End Function
    Public Function CheckClient(ByVal ID As Integer) As Boolean
        Dim bytes1(0) As Byte
        bytes1(0) = 0
        Try
            Clients(ID).CSocket.BeginSend(bytes1, 0, bytes1.Length, SocketFlags.None, New AsyncCallback(AddressOf OnSend), ID)
            CheckClient = True
        Catch ex As Exception
            CheckClient = False
        End Try
    End Function
    Private Sub OnAccept(ByVal ar As IAsyncResult)
        Dim ID As Integer = GetClientID()
        Dim m_Client As Client_Type
        Try
            RaiseEvent SocketError(ID, "===============================OnAccepting:" & ID & "===============================", False)
            If ID <> -1 Then
                'ReDim Preserve Clients(ID)
                m_Client.Recbytes = New Byte(10000) {}
                Try
                    m_Client.ISConnect = True
                    ' 
                    ' ArrayConnectedClient.Insert(ID, ID)
                    RaiseEvent SocketError(ID, "Client by ID " & ID & " is Accepted.", False)
                    m_Client.CSocket = tcpListener.EndAccept(ar)

                    Dim ip As IPEndPoint = m_Client.CSocket.RemoteEndPoint
                    m_Client.Ip = ip.Address.ToString
                    m_Client._WMOCODE = ""
                    '   m_Client._flagReadReport = False
                    m_Client._flagDisconnecting = False
                    Clients(ID) = m_Client
                    '   m_Client.CSocket.
                    RaiseEvent Connected(ID, ip.Address.ToString)
                    RaiseEvent SocketError(ID, "Client by ID " & ID & " IP=" & ip.Address.ToString, False)
                    RaiseEvent SocketError(ID, "Client by ID " & ID & " is Ready to use", False)
                Catch ex As SocketException
                    RaiseEvent SocketError(ID, "OnAccept err00:" & ex.Message, False)
                Catch ex1 As Exception
                    RaiseEvent SocketError(ID, "OnAccept err01:" & ex1.Message, False)
                End Try
                Try
                    Array.Clear(Clients(ID).Recbytes, 0, Clients(ID).Recbytes.Length)
                Catch ex As Exception

                End Try
                Try
                    Clients(ID).CSocket.BeginReceive(Clients(ID).Recbytes, 0, Clients(ID).Recbytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), ID)
                    RaiseEvent SocketError(ID, "Client by ID " & ID & " is ready to recieved data.", False)
                Catch ex As SocketException
                    RaiseEvent SocketError(ID, "OnAccept err10:" & ex.Message, False)
                Catch ex1 As Exception
                    RaiseEvent SocketError(ID, "OnAccept err11:" & ex1.Message, False)
                End Try
            Else
                tcpListener.EndAccept(ar).Disconnect(True)
            End If
        Catch ex As Exception
            RaiseEvent SocketError(ID, "OnAccept err2:" & ex.Message, False)
            '  Writ("OnAccept err:" & ex.Message, ID)
        End Try
        Try

            tcpListener.BeginAccept(New AsyncCallback(AddressOf OnAccept), tcpListener)
            RaiseEvent SocketError(ID, "tcpListener is ready ", False)
        Catch ex As SocketException
            RaiseEvent SocketError(ID, "OnAccept err30:" & ex.Message, False)
        Catch ex1 As Exception
            RaiseEvent SocketError(ID, "OnAccept err31:" & ex1.Message, False)
        End Try
        'RaiseEvent SocketError(ID, "OnAccept finish", False)
        RaiseEvent SocketError(ID, "+++++++++++++++++++++++++++++ Accept finish +++++++++++++++++++++++++++++", False)
    End Sub
    Private Sub OnSend(ByVal ar As IAsyncResult)
        'hh
        Dim ID As Integer = ar.AsyncState
        Dim er As SocketError
        Try

            Clients(ID).CSocket.EndSend(ar, er)
            If er <> Sockets.SocketError.Success Then
                RaiseEvent SocketError(ID, "OnSend error:" & er.ToString, True)
            End If

            '  Clients(ID).CSocket.BeginReceive(Clients(ID).Recbytes, 0, Clients(ID).Recbytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), ID)
        Catch ex As Exception
            RaiseEvent SocketError(ID, "OnSend error2:" & ex.Message, True)
        End Try
    End Sub

    Private Sub OnDisConnect(ByVal ar As IAsyncResult)

        Dim id As Integer = ar.AsyncState

        Try
            Clients(id).CSocket.EndDisconnect(ar)

            Clients(id).Ip = ""
            Clients(id)._WMOCODE = ""
            Clients(id).ISConnect = False
            RaiseEvent SocketError(id, "Clieint by id " & id & " Set isconnect=False.", False)
            RaiseEvent Disconnected(id)

            RaiseEvent SocketError(id, "Clieint by id " & id & " Removed.", False)
            '  Clients(id).CSocket = Nothing
            '  Clients(id) = Nothing
            GC.Collect()
        Catch ex As Exception
            '  WritesocketLog("OnDisConnect 1" & ex.Message)
            RaiseEvent SocketError(id, "OnDisConnect error:" & ex.Message, True)
        End Try
    End Sub

    Private Sub OnRecieve(ByVal ar As IAsyncResult)

        Dim Status1 As String = ""
        Dim flager As Boolean = False
        Dim ID As Integer = -1
        Try
            Try
                ID = ar.AsyncState
                'If Clients(ID).CSocket Is Nothing Then

                '    'Clients(ID).CSocket.Disconnect(False)
                'End If
                Dim flagDisconnecting As Boolean = False
                Try
                    flagDisconnecting = Clients(ID)._flagDisconnecting
                Catch ex As Exception
                    RaiseEvent SocketError(ID, "OnRecieve(889):" & ex.Message, False)
                End Try
                If flagDisconnecting Then
                    RaiseEvent SocketError(ID, "Clientby id " & ID & " Recieved some thing during disconnecting.", False)
                Else
                    Try
                        '    ID = ar.AsyncState

                        '    RaiseEvent SocketError(ID, "Client   " & ID & "Recieved some thing during disconnecting.", False)

                        If Clients(ID).CSocket.Poll(1, SelectMode.SelectRead) And Clients(ID).CSocket.Available = 0 Then
                            RaiseEvent SocketError(ID, " OnRecieve is for disconnect ", False)
                            Try
                                '    Status1 = Status1 & " Disconnect SignaL REC  " & vbCrLf
                                Disconnect(ID)
                            Catch ex As Exception
                                RaiseEvent SocketError(ID, " Disconnect in recieve  error  " & ex.Message, False)
                                flager = True
                            End Try
                            Exit Sub
                        Else
                        End If
                        '   Status = Status & "Read Buffer Start  " & vbCrLf
                        Dim er As SocketError
                        Dim Len As Integer = Clients(ID).CSocket.EndReceive(ar, er)
                        '    RaiseEvent SocketError(ID, "OnRecieve Data Len " & Len, False)
                        '  Status = Status & " Len is " & Len & vbCrLf
                        If er <> Sockets.SocketError.Success Then
                            RaiseEvent SocketError(ID, "OnRecieve error0:" & er.ToString, True)
                        Else
                            If Len > 0 Then
                                Dim message As String
                                Try
                                    message = System.Text.ASCIIEncoding.ASCII.GetString(Clients(ID).Recbytes, 0, Len)
                                    '  Status = Status & "read message end " & vbCrLf & message
                                    System.Threading.Thread.Sleep(200)
                                Catch ex As Exception
                                    RaiseEvent SocketError(ID, "OnRecieve error1:" & er.ToString, True)
                                End Try
                                Try
                                    '  Status = Status & "Clients.Length=" & Clients.Length & vbCrLf
                                    '   Status = Status & "ID=" & ID & vbCrLf & message & vbCrLf
                                    Array.Clear(Clients(ID).Recbytes, 0, Clients(ID).Recbytes.Length)

                                Catch ex As Exception
                                    flager = True
                                    ' Status = Status & "clear buffer  error" & vbCrLf & message
                                    RaiseEvent SocketError(ID, "clear error:" & ex.Message, False)
                                End Try

                                Try
                                    If message <> "" Then
                                        RaiseEvent DataArival(ID, message, Clients(ID).Ip)
                                    End If
                                Catch ex As Exception
                                    flager = True
                                    RaiseEvent SocketError(ID, "OnRecieve error0:" & ex.Message, False)
                                    '  RaiseEvent SocketError(ID, Status, False)
                                End Try

                                '   Status1 = Status1 & "RaiseDataArival"

                                Try
                                    Clients(ID).CSocket.BeginReceive(Clients(ID).Recbytes, 0, Clients(ID).Recbytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), ID)
                                Catch ex As Exception
                                    flager = True
                                    RaiseEvent SocketError(ID, "OnRecieve error 1:" & ex.Message, False)
                                    '   RaiseEvent SocketError(ID, Status1, False)
                                End Try
                                '   Status1 = Status1 & "ready for rec " & vbCrLf

                            Else
                                Array.Clear(Clients(ID).Recbytes, 0, Clients(ID).Recbytes.Length)
                                Clients(ID).CSocket.BeginReceive(Clients(ID).Recbytes, 0, Clients(ID).Recbytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), ID)
                            End If
                        End If
                    Catch e1 As SocketException
                        flager = True
                        RaiseEvent SocketError(ID, "OnRecieve socket error:" & e1.Message, False)
                        '   RaiseEvent SocketError(ID, Status1, False)

                    Catch ex As Exception
                        flager = True
                        '  RaiseEvent SocketError(ID, Status1, False)
                        RaiseEvent SocketError(ID, "OnRecieve(890):" & ex.Message, False)
                    End Try


                End If
            Catch ex As Exception
                flager = True
                RaiseEvent SocketError(ID, "OnRecieve(891):" & ex.Message, False)
            End Try

            If flager Then
                RaiseEvent SocketError(ID, "Recieve Finish", False)
            End If
        Catch ex As InvalidOperationException
            RaiseEvent SocketError(ID, "Recieve Finish:" & ex.Message, False)
        End Try
    End Sub
#End Region
#Region "Methods"

    Public Sub SendData(ByVal Data As String, ByVal ClientID As Integer)
        ' WriteSocket(Data, ClientID)
        Dim bytes1 As Byte() = ASCII.GetBytes(Data)
        Try
            Clients(ClientID).CSocket.BeginSend(bytes1, 0, bytes1.Length, SocketFlags.None, New AsyncCallback(AddressOf OnSend), ClientID)

        Catch ex As Exception
            RaiseEvent SocketError(ClientID, "SendData Error:" & ex.Message, False)
        End Try
    End Sub
    Public Sub Disconnect(ByVal ID As Integer)
        '        .BeginSend(bytes1, 0, bytes1.Length, SocketFlags.None, New AsyncCallback(AddressOf OnSend), client)

        '   WritesocketLog("Disconnect 1")
        Try


            RaiseEvent SocketError(ID, "Clieint by id " & ID & " check for  disconnect.", False)
            If Clients(ID).ISConnect = True Then
                RaiseEvent SocketError(ID, "Disconnect 1", False)
                Clients(ID).CSocket.Shutdown(SocketShutdown.Both)
                RaiseEvent SocketError(ID, "Clieint by id " & ID & " Start BeginDisconnect.", False)
                Clients(ID)._flagDisconnecting = True
                Clients(ID).CSocket.BeginDisconnect(True, New AsyncCallback(AddressOf OnDisConnect), ID)
            Else

                RaiseEvent SocketError(ID, "Disconnect 1", False)
                Clients(ID).Ip = ""
                Clients(ID).ISConnect = False
                RaiseEvent SocketError(ID, "Clieint by id " & ID & " Set isconnect=False.", False)
                '   ArrayConnectedClient.Remove(ID)
                ' Clients(ID).CSocket = Nothing
                ' Clients(ID) = Nothing
                RaiseEvent Disconnected(ID)

                GC.Collect()
            End If
        Catch ex As Exception
            RaiseEvent SocketError(ID, "Disconnect Error:" & ex.Message, False)
        End Try
    End Sub
    'Public Sub Listen()
    '    'Dim ip As New IPEndPoint(IPAddress.Any, _Port)
    '    'tcpListener.Bind(ip)
    '    'tcpListener.Listen(100)
    '    'tcpListener.BeginAccept(New AsyncCallback(AddressOf OnAccept), tcpListener)

    'End Sub
#End Region
#Region "Events"
    Public Event DataArival(ByVal ClientID As Integer, ByVal Data As String, ByVal ip As String)
    Public Event SocketError(ByVal ClientID As Integer, ByVal ER As String, ByVal IsDisconnect As Boolean)
    ' Public Event SocketError1(ByVal ClientID As Integer, ByVal ER As SocketError)
    Public Event Disconnected(ByVal ClientID As Integer)
    Public Event Connected(ByVal ClientID As Integer, ByVal ip As String)
#End Region
#Region "Peroperties"""
    Public Property WMOCode(ByVal id As Integer)
        'hh
        Get
            Try


                Return Clients(id)._WMOCODE
            Catch ex As Exception
                Return "99999"
                RaiseEvent SocketError(id, "Property error get wmo", False)
                Clients(id)._flagDisconnecting = True
            End Try

        End Get
        Set(ByVal value)
            Try
                Clients(id)._WMOCODE = value
            Catch ex As Exception
                RaiseEvent SocketError(id, "Property error set wmo", False)
            End Try

        End Set
    End Property
    '_flagDisconnecting
    Public Property Port()
        Get
            Return _Port
        End Get
        Set(ByVal value)
            _Port = value
        End Set
    End Property


#End Region


End Class

