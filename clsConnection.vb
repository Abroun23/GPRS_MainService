Public Class clsConnection
#Region "Field"
    Dim m_ServerName As String
    Dim DataSource As String
    Dim Userid As String
    Dim Password As String
    Dim Integrated_Security As Boolean
    Dim Persist_Security_Info As Boolean
    Dim ShowPassInConnectionString As Boolean
    Dim SecurityType As ConnectionAutherithy
    Dim ConnectionTimeOut As Integer
    Dim m_connectionString As String
    'Data Source=helia-pc;Initial Catalog=HM_DB;Persist Security Info=True;User ID=sa;Connect Timeout=500
    Enum ConnectionAutherithy
        SqlServer = 1 ' ' Integrated_Security=true  UserId and pass
        Windows = 2  ' Integrated_Security=true


    End Enum
#End Region
#Region "Events"
    Public Event WriteLog(ByVal log As String)
    Public Event Connection_Error(ByVal Errorlog As String)
#End Region
#Region "body"
    Sub New(ByVal ConnectionString As String)

        '     Con.ConnectionString = ConnectionString
        ParsConnection(ConnectionString)
        m_connectionString = ConnectionString
    End Sub
    Sub New(ByVal ServerName As String, ByVal DataSource As String, ByVal USERID As String, ByVal Password As String, Optional ByVal TimeOut As Integer = 15)
        m_ServerName = ServerName
        DataSource = DataSource
        USERID = USERID
        Password = Password
        ConnectionTimeOut = TimeOut
        SecurityType = ConnectionAutherithy.SqlServer

        m_connectionString = GenerateConnection()
    End Sub
    Sub New(ByVal ServerName As String, ByVal DataSource As String, Optional ByVal TimeOut As Integer = 15)
        m_ServerName = ServerName
        DataSource = DataSource
        ConnectionTimeOut = TimeOut
        SecurityType = ConnectionAutherithy.Windows
        m_connectionString = GenerateConnection()
    End Sub
    Public Sub ParsConnection(ByVal ConnectionString As String)
        Dim ConnectionInfo() = Split(ConnectionString, ";")
        For i As Integer = 1 To ConnectionInfo.Length - 1
            Dim strsplit() = Split(ConnectionInfo(0), "=")
            '  ServerName = strsplit(1)
            Select Case Trim(UCase(strsplit(0)))
                Case "DATASOURCE"
                    m_ServerName = strsplit(1)
                Case "INITIALCATALOG"
                    DataSource = strsplit(1)
                Case "PERSISTSECURITYINFO"
                    Persist_Security_Info = strsplit(1)
                Case "INTEGRATEDSECURITY"
                    Integrated_Security = strsplit(1)
                Case "USERID"
                    Userid = strsplit(1)
                Case "PASSWORD"
                    Password = strsplit(1)
                Case "CONNECTIONTIMEOUT"
                    ConnectionTimeOut = strsplit(1)

            End Select
        Next i
    End Sub
    Public Function GenerateConnection() As String
        'Data Source=helia-pc;Initial Catalog=HM_DB;Persist Security Info=True;User ID=sa;Connect Timeout=500
        m_connectionString = "Data Source=" & m_ServerName & ";Initial Catalog=" & DataSource
        If SecurityType = ConnectionAutherithy.Windows Then
            m_connectionString = +";Integerated security=True;Connect Timeout=" & ConnectionTimeOut
        Else
            m_connectionString = +";Persist Security Info=True;User ID=" & Userid & ";Password=" & Password & ";Connect Timeout=" & ConnectionTimeOut

        End If
        GenerateConnection = connectionString
    End Function

#End Region
#Region "Methods"
    'Public Sub Open()
    '    Try

    '        If con.State = ConnectionState.Open Then
    '            con.Close()
    '        End If
    '        con.Open()
    '        RaiseEvent WriteLog("Connection is Opened. ")
    '    Catch ex As Exception
    '        RaiseEvent Connection_Error(ex.Message)
    '    End Try
    'End Sub
    'Public Sub Close()
    '    con.Close()
    '    RaiseEvent WriteLog("Connection is Closed. ")
    'End Sub
#End Region
#Region "Peroperties"
    ReadOnly Property ConnectionString()
        Get
            Return m_connectionString
        End Get
    End Property

    ReadOnly Property ServerName()
        Get
            Return m_ServerName
        End Get
    End Property


#End Region
End Class
