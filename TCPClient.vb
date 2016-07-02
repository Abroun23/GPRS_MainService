
'Imports System.Net, System.Net.Sockets
'Imports System.Text.ASCIIEncoding
'Imports System.Windows.Forms
'Imports System.IO
''
'Imports vb = Microsoft.VisualBasic
'Imports System.Threading

'Public Class TcpClient
'    Dim Client As Socket
'    Dim _host As String
'    Dim _port As Integer
'    Dim _objectID As Long
'    Dim bytes As Byte() = New Byte(1023) {}
'    Private Shared disconnectDone As New ManualResetEvent(False)


'    Dim WithEvents tmrTimeOut As System.Timers.Timer
'    Dim WithEvents tmrcheckCounter As System.Timers.Timer
'    Dim fladrec As Boolean = False
'    Dim _ISCONNECT As Boolean = False
'    Enum _State
'        Conneting = 1
'        DisConnecting = 2
'        Sendingdata = 3
'        RecieveingData = 4
'    End Enum
'    Property ObjectID() As Long
'        Get
'            Return _objectID

'        End Get
'        Set(ByVal value As Long)
'            _objectID = value
'        End Set
'    End Property




'    Dim Mystate As _State
'    Public Sub New(ByVal host As String, ByVal port As String)
'        Control.CheckForIllegalCrossThreadCalls = False
'        _host = host
'        _port = port
'        Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
'        tmrTimeOut = New System.Timers.Timer
'        tmrTimeOut.Interval = 60000
'        tmrcheckCounter = New System.Timers.Timer
'        tmrcheckCounter.Interval = 60000
'        tmrcheckCounter.Enabled = True
'    End Sub
'    Public Event DataArrival(ByVal Data As String)
'    Public Event Disconnected()
'    '   Public Event SocketError(ByVal er As SocketError)
'    Public Event ExceptionError(ByVal er As String)
'    Public Event Connected()
'    Public Event TimeOut(ByVal state As _State)
'#Region "Body"
'    Private Sub OnConnect(ByVal ar As IAsyncResult)
'        Try
'            Client.EndConnect(ar)
'        Catch ex As Exception

'        End Try

'        Try

'            _ISCONNECT = True
'            tmrTimeOut.Enabled = False
'            Client.BeginReceive(bytes, 0, bytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), Client)
'            RaiseEvent Connected()
'        Catch ex As SocketException
'            tmrTimeOut.Enabled = False
'            RaiseEvent ExceptionError(ex.Message)
'        End Try
'    End Sub
'    Private Sub OnSend(ByVal ar As IAsyncResult)
'        Dim er As SocketError
'        tmrTimeOut.Enabled = False
'        Try
'            If Client.Connected Then
'                Client.EndSend(ar, er)
'                If er <> Sockets.SocketError.Success Then
'                    ' Else
'                    RaiseEvent ExceptionError(er.ToString)
'                End If
'                'Client.BeginReceive(bytes, 0, bytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), Client)
'            End If
'            '  tmrTimeOut.Enabled = False
'        Catch ex As Exception
'            RaiseEvent ExceptionError(ex.Message)
'        End Try
'    End Sub
'    Private Sub OnDisconnect(ByVal ar As IAsyncResult)
'        'tmrTimeOut.Enabled = False
'        'Try
'        '    client.EndDisconnect(ar)
'        '    client.BeginDisconnect(SocketShutdown.Both)
'        '    client = Nothing
'        'Catch ex As Exception
'        'End Try
'        '_ISCONNECT = False
'        'RaiseEvent Disconnected()


'        ' Dim client As Socket = CType(ar.AsyncState, Socket)

'        Try
'            Client.EndDisconnect(ar)
'        Catch ex As Exception
'            '  RaiseEvent Exception(ex)
'        End Try

'        disconnectDone.Set()

'        _ISCONNECT = False
'        'Client = Nothing
'        RaiseEvent Disconnected()

'    End Sub

'    Private Sub OnRecieve(ByVal ar As IAsyncResult)
'        Dim er As SocketError

'        Try
'            If Client.Poll(1, SelectMode.SelectRead) And Client.Available = 0 Then
'                _ISCONNECT = False
'                Disconnect()
'                Exit Sub
'            End If
'            If Client.Connected Then
'                fladrec = True
'                Client = ar.AsyncState
'                Dim len As Integer = Client.EndReceive(ar, er)
'                If len > 0 Then


'                    If er = Sockets.SocketError.Success Then
'                        Dim message As String = System.Text.ASCIIEncoding.ASCII.GetString(bytes, 0, len)
'                        Array.Clear(bytes, 0, bytes.Length)
'                        Thread.Sleep(1000)
'                        RaiseEvent DataArrival(message)
'                    Else
'                        RaiseEvent ExceptionError(er.ToString)
'                    End If
'                    '*****
'                    Array.Clear(bytes, 0, bytes.Length - 1)
'                    Try


'                        Client.BeginReceive(bytes, 0, bytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), Client)

'                    Catch ex As Exception
'                        RaiseEvent ExceptionError(ex.Message)
'                    End Try
'                    '*****
'                End If
'            End If
'        Catch ex As Exception
'            RaiseEvent ExceptionError(ex.Message)
'        End Try
'    End Sub

'#End Region
'#Region "Property"
'    Public Property host()
'        Get
'            Return host
'        End Get
'        Set(ByVal value)
'            _host = value
'        End Set
'    End Property
'    Public Property Port()
'        Get
'            Return _port
'        End Get
'        Set(ByVal value)
'            _port = value
'        End Set
'    End Property
'    Public ReadOnly Property IsConnected As Boolean
'        Get
'            Return _ISCONNECT
'        End Get
'    End Property

'#End Region

'    Private Sub tmr_Timeout_tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrTimeOut.Elapsed
'        ' fladrec = False
'        tmrTimeOut.Enabled = False
'        '   Connect()
'        RaiseEvent TimeOut(Mystate)
'    End Sub

'#Region "Methodes"
'    'Public Function IsConnected() As Boolean
'    '    If _ISCONNECT = True Then
'    '        'Try
'    '        '    Dim bytes1 As Byte() = ASCII.GetBytes(vbCrLf)
'    '        '    Dim er As SocketError
'    '        '    Client.Send(bytes1, 0, bytes1.Length, SocketFlags.None, er)
'    '        '    If er = Sockets.SocketError.Success Then
'    '        '        _ISCONNECT = True
'    '        '    Else
'    '        '        _ISCONNECT = False
'    '        '        Disconnect()
'    '        '    End If
'    '        'Catch ex As Exception
'    '        '    _ISCONNECT = False
'    '        '    ' RaiseEvent Disconnect(_ObjectID)
'    '        '    Disconnect()
'    '        'End Try
'    '    Else
'    '        _ISCONNECT = False

'    '    End If
'    '    IsConnected = _ISCONNECT
'    ' End Function
'    Public Sub SendData(ByVal Data As String)
'        Mystate = _State.Sendingdata
'        tmrTimeOut.Enabled = True
'        Dim bytes1 As Byte()
'        bytes1 = ASCII.GetBytes(Data)
'        Try

'            Client.BeginSend(bytes1, 0, bytes1.Length, SocketFlags.None, New AsyncCallback(AddressOf OnSend), Client)
'        Catch ex As Exception
'            tmrTimeOut.Enabled = False
'            RaiseEvent ExceptionError(ex.Message)
'        End Try
'    End Sub
'    Public Sub Disconnect()
'        '   Client = Nothing
'        '  Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)


'        Mystate = _State.DisConnecting
'        tmrTimeOut.Enabled = True
'        Try

'            Client.Shutdown(SocketShutdown.Both)

'            Client.BeginDisconnect(True, New AsyncCallback(AddressOf OnDisconnect), Nothing)
'            '   Client.Disconnec(True)
'            ' disconnectDone.WaitOne
'        Catch ex As Exception
'            tmrTimeOut.Enabled = False
'            'Client = Nothing
'            'Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
'            'RaiseEvent Disconnected()
'        Finally
'            'Client.Close()
'            'Client.Dispose()
'        End Try
'    End Sub
'    Public Sub Connect()
'        '     Client.Close()
'        'Client.Dispose()
'        Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
'        Mystate = _State.Conneting
'        tmrTimeOut.Enabled = True
'        Dim IP As IPAddress = IPAddress.Parse(_host)
'        Dim xIpEndPoint As IPEndPoint = New IPEndPoint(IP, _port)
'        Try
'            Client.BeginConnect(xIpEndPoint, New AsyncCallback(AddressOf OnConnect), Nothing)
'        Catch ex As Exception
'            tmrTimeOut.Enabled = False
'            RaiseEvent ExceptionError(ex.Message)
'        End Try
'    End Sub
'#End Region

'End Class


Imports System.Net, System.Net.Sockets
Imports System.Text.ASCIIEncoding
Imports System.Windows.Forms
Imports System.IO
'
Imports vb = Microsoft.VisualBasic
Imports System.Threading

Public Class TcpClient
    Dim Client As Socket
    Dim _host As String = "80.191.68.213"
    Dim _port As Integer = "5000"
    Dim _objectID As Long
    Dim bytes As Byte() = New Byte(1023) {}
    Private Shared disconnectDone As New ManualResetEvent(False)


    Dim WithEvents tmrTimeOut As System.Timers.Timer
    Dim WithEvents tmrcheckCounter As System.Timers.Timer
    Dim fladrec As Boolean = False
    Dim _ISCONNECT As Boolean = False
    Enum _State
        Conneting = 1
        DisConnecting = 2
        Sendingdata = 3
        RecieveingData = 4
    End Enum
    Property ObjectID() As Long
        Get
            Return _objectID

        End Get
        Set(ByVal value As Long)
            _objectID = value
        End Set
    End Property




    Dim Mystate As _State
    Public Sub New(ByVal host As String, ByVal port As String)
        _host = host
        _port = port
        Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        tmrTimeOut = New System.Timers.Timer
        tmrTimeOut.Interval = 60000
        tmrcheckCounter = New System.Timers.Timer
        tmrcheckCounter.Interval = 60000
        tmrcheckCounter.Enabled = True
    End Sub
    Public Event DataArrival(ByVal Data As String)
    Public Event Disconnected()
    '   Public Event SocketError(ByVal er As SocketError)
    Public Event ExceptionError(ByVal er As String)
    Public Event Connected()
    Public Event TimeOut(ByVal state As _State)
#Region "Body"
    Private Sub OnConnect(ByVal ar As IAsyncResult)
        Try
            Client.EndConnect(ar)
            _ISCONNECT = True
            tmrTimeOut.Enabled = False

            RaiseEvent Connected()
            Client.BeginReceive(bytes, 0, bytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), Client)

        Catch ex As SocketException
            RaiseEvent ExceptionError("OnConnect" & ex.Message)
        End Try
    End Sub
    Private Sub OnSend(ByVal ar As IAsyncResult)
        Dim er As SocketError
        tmrTimeOut.Enabled = False
        Try
            If Client.Connected Then
                Client.EndSend(ar, er)
                If er <> Sockets.SocketError.Success Then
                    ' Else
                    RaiseEvent ExceptionError("OnSend" & er.ToString)
                End If
                'Client.BeginReceive(bytes, 0, bytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), Client)
            End If
            '  tmrTimeOut.Enabled = False
        Catch ex As Exception
            RaiseEvent ExceptionError("OnSend" & ex.Message)
        End Try
    End Sub
    Private Sub OnDisconnect(ByVal ar As IAsyncResult)
        'tmrTimeOut.Enabled = False
        'Try
        '    client.EndDisconnect(ar)
        '    client.BeginDisconnect(SocketShutdown.Both)
        '    client = Nothing
        'Catch ex As Exception
        'End Try
        '_ISCONNECT = False
        'RaiseEvent Disconnected()


        ' Dim client As Socket = CType(ar.AsyncState, Socket)

        Try
            Client.EndDisconnect(ar)
        Catch ex As Exception
            disconnectDone.Set()
            RaiseEvent ExceptionError("OnDisconnect" & ex.Message)
        End Try



        _ISCONNECT = False
        RaiseEvent Disconnected()

    End Sub

    Private Sub OnRecieve(ByVal ar As IAsyncResult)
        Dim er As SocketError

        Try
            If Client.Poll(1, SelectMode.SelectRead) And Client.Available = 0 Then
                _ISCONNECT = False
                Disconnect()
                Exit Sub
            End If
            If Client.Connected Then
                fladrec = True
                Client = ar.AsyncState

                Dim len As Integer = Client.EndReceive(ar, er)
                If er = Sockets.SocketError.Success Then
                    RaiseEvent ExceptionError("OnRecieveer" & er.ToString)
                ElseIf len > 0 Then


                    If er = Sockets.SocketError.Success Then
                        Dim message As String = System.Text.ASCIIEncoding.ASCII.GetString(bytes)
                        Array.Clear(bytes, 0, bytes.Length)
                        Thread.Sleep(1000)
                        RaiseEvent DataArrival(message)
                    Else
                        RaiseEvent ExceptionError("OnRecieve" & er.ToString)
                    End If
                    '*****
                    Array.Clear(bytes, 0, bytes.Length - 1)
                    Try


                        Client.BeginReceive(bytes, 0, bytes.Length, SocketFlags.None, New AsyncCallback(AddressOf OnRecieve), Client)

                    Catch ex As Exception
                        RaiseEvent ExceptionError("OnRecieve 1" & er.ToString)
                    End Try
                    '*****
                End If
            End If
        Catch ex As Exception
            RaiseEvent ExceptionError("OnRecieve2 " & ex.Message)
        End Try
    End Sub

#End Region
#Region "Property"
    Public Property host()
        Get
            Return host
        End Get
        Set(ByVal value)
            _host = value
        End Set
    End Property
    Public Property Port()
        Get
            Return _port
        End Get
        Set(ByVal value)
            _port = value
        End Set
    End Property
    Public ReadOnly Property IsConnected As Boolean
        Get
            Return _ISCONNECT
        End Get
    End Property

#End Region

    Private Sub tmr_Timeout_tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrTimeOut.Elapsed
        ' fladrec = False
        tmrTimeOut.Enabled = False
        '   Connect()
        RaiseEvent TimeOut(Mystate)
    End Sub

#Region "Methodes"
    'Public Function IsConnected() As Boolean
    '    If _ISCONNECT = True Then
    '        'Try
    '        '    Dim bytes1 As Byte() = ASCII.GetBytes(vbCrLf)
    '        '    Dim er As SocketError
    '        '    Client.Send(bytes1, 0, bytes1.Length, SocketFlags.None, er)
    '        '    If er = Sockets.SocketError.Success Then
    '        '        _ISCONNECT = True
    '        '    Else
    '        '        _ISCONNECT = False
    '        '        Disconnect()
    '        '    End If
    '        'Catch ex As Exception
    '        '    _ISCONNECT = False
    '        '    ' RaiseEvent Disconnect(_ObjectID)
    '        '    Disconnect()
    '        'End Try
    '    Else
    '        _ISCONNECT = False

    '    End If
    '    IsConnected = _ISCONNECT
    ' End Function
    Public Sub SendData(ByVal Data As String)
        Mystate = _State.Sendingdata
        tmrTimeOut.Enabled = True
        Dim bytes1 As Byte()
        bytes1 = ASCII.GetBytes(Data)
        Try

            Client.BeginSend(bytes1, 0, bytes1.Length, SocketFlags.None, New AsyncCallback(AddressOf OnSend), Client)
        Catch ex As Exception
            tmrTimeOut.Enabled = False
            RaiseEvent ExceptionError("SendData" & ex.Message)
        End Try
    End Sub
    Public Sub Disconnect()
        '   Client = Nothing
        '  Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)


        Mystate = _State.DisConnecting
        tmrTimeOut.Enabled = True
        Try
            'Clie
            Client.Shutdown(SocketShutdown.Both)

            Client.BeginDisconnect(True, New AsyncCallback(AddressOf OndisConnect), Nothing)
            '   Client.Disconnec(True)
            ' disconnectDone.WaitOne
        Catch ex As Exception
            tmrTimeOut.Enabled = False
            'Client = Nothing
            'Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            RaiseEvent ExceptionError("Disconnect" & ex.Message)
        End Try
    End Sub
    Public Sub Connect()
        Mystate = _State.Conneting
        tmrTimeOut.Enabled = True
        Dim IP As IPAddress = IPAddress.Parse(_host)
        Dim xIpEndPoint As IPEndPoint = New IPEndPoint(IP, _port)
        Try
            Client.BeginConnect(xIpEndPoint, New AsyncCallback(AddressOf OnConnect), Nothing)

        Catch ex As Exception
            tmrTimeOut.Enabled = False
            RaiseEvent ExceptionError("Connect" & ex.Message)
        End Try
    End Sub
#End Region

End Class

