Imports System.IO
Public Class clsDataBase
#Region "Feild"
    Dim ErrorLogADD As String
    Dim TblSensorsUpdate1 As Boolean = False
    Dim m_Connection As clsConnection
    Dim ApplicationAdd As String=My.Application.Info.DirectoryPath
#End Region
    '#Region "Event"
    '    Public Event DataBaseErrors(ByVal StrError As String)
    '#End Region
#Region "Body"

    Sub New(ByVal ConnectionString As String, ByVal _ErrorLogADD As String)
        'm_ConnectionString = ConnectionString
        ErrorLogADD = _ErrorLogADD
        m_Connection = New clsConnection(ConnectionString)
        '   WriteLogs("Start")
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ServerName"></param>
    ''' <param name="DataSource"></param>
    ''' <param name="USERID"></param>
    ''' <param name="Password"></param>
    ''' <param name="TimeOut"></param>
    ''' <remarks>Time Out Default=15</remarks>
    Sub New(ByVal ServerName As String, ByVal DataSource As String, ByVal USERID As String, ByVal Password As String, ByVal _ErrorLogADD As String, Optional ByVal TimeOut As Integer = 15)
        m_Connection = New clsConnection(ServerName, DataSource, USERID, Password, TimeOut)
        ErrorLogADD = _ErrorLogADD
    End Sub
    Sub New(ByVal ServerName As String, ByVal DataSource As String, ByVal _ErrorLogADD As String, Optional ByVal TimeOut As Integer = 15)
        m_Connection = New clsConnection(ServerName, DataSource, TimeOut)
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region
#Region "Methods"
    Public Function updateIntotblReport_ReportMessage(ByVal ReportContent As String, ByVal ID As String) As Long




        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim m_CON As New SqlClient.SqlConnection
        Dim flagupdate As Boolean = False
        m_CON.ConnectionString = m_Connection.ConnectionString
        Dim Count As Integer = 0
        Dim CMD As New SqlClient.SqlCommand
        CMD.Connection = m_CON
        If m_CON.State = ConnectionState.Closed Then
            m_CON.Open()
        End If
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD

        Try



            CMD.Parameters.Clear()
            CMD.CommandText = "usp_tblReports_update_Center_ReportContent"
            CMD.Parameters.AddWithValue("@ReportID", ID)
            CMD.Parameters.AddWithValue("@ReportContent", ReportContent)

            CMD.ExecuteNonQuery()
            ' ID.Add(.Item("ReportID"))
            '     WriteLogs(ID.Count)
            '   ID.Add(.Item("MetarID"))
            '  ReDim Preserve ID(Count)
            '  ID(Count) = .Item("MetarID")

        Catch ex As Exception
            WriteLogs("usp_tblReports_update_Center_ReportContent :" & ex.Message)
        End Try

        m_CON.Close()
        Return 0
    End Function
#Region "NDD"
    'usp_tblProductsForStationsSendingQueue_updateByWmoCode
    Public Sub usp_tblProductsForStationsSendingQueue_updateByWmoCode(ByVal WMOCode As String, ByVal path As String)
        Try

            Dim CMD As New SqlClient.SqlCommand
            Dim m_DataAdapter As New SqlClient.SqlDataAdapter
            Dim m_CON As New SqlClient.SqlConnection
            Dim m_Dataset As New DataSet
            CMD.Connection = m_CON
            m_CON.ConnectionString = m_Connection.ConnectionString
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = "usp_tblProductsForStationsSendingQueue_updateByWmoCode"
            '@DirectoryPath,@DirectorySize,@DrectoryModify1
            m_DataAdapter.SelectCommand = CMD
            CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
            WriteLogs("usp_tblProductsForStationsSendingQueue_updateByWmoCode " & path)
            WriteLogs("usp_tblProductsForStationsSendingQueue_updateByWmoCode " & WMOCode)
            CMD.Parameters.AddWithValue("@Path", path)
            CMD.ExecuteNonQuery()
        Catch ex As Exception
            WriteLogs("usp_tblProductsForStationsSendingQueue_updateByWmoCode" & ex.Message)
        End Try
    End Sub
    Public Sub usp_tblProductsForStationsSendingQueue_SelectByWmoCode(ByRef tblInfo As DataTable, ByVal WMOCode As String)
        Try

            Dim CMD As New SqlClient.SqlCommand
            Dim m_DataAdapter As New SqlClient.SqlDataAdapter
            Dim m_CON As New SqlClient.SqlConnection
            Dim m_Dataset As New DataSet
            CMD.Connection = m_CON
            m_CON.ConnectionString = m_Connection.ConnectionString
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = "usp_tblProductsForStationsSendingQueue_SelectByWmoCode"
            '@DirectoryPath,@DirectorySize,@DrectoryModify
            m_DataAdapter.SelectCommand = CMD
            CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
            m_DataAdapter.Fill(m_Dataset, "tblInfo")
            tblInfo = m_Dataset.Tables("tblInfo")
        Catch ex As Exception
            WriteLogs("usp_tblProductsForStationsSendingQueue_SelectByWmoCode" & ex.Message)
        End Try
    End Sub
#End Region
#Region "checkACK"
    Public Function Update_tblSendToSwitchQueue(ByVal ReportID As String, ByVal DestinationName As String, ByVal FK_ReportTypeID As Long, ByVal WMOCode As String, ByVal UDT As String) As Long
        Dim lngRes As Long = 0
        Dim con As New SqlClient.SqlConnection
        con.ConnectionString = My.Settings.ICSDBConnectionString
        Dim Count As Integer = 0
        Dim CMD As New SqlClient.SqlCommand
        CMD.Connection = con
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If


        CMD.CommandType = CommandType.StoredProcedure
        'CMD.CommandText = "UPDATE tblSendToSwitchQueue " & _
        '                  "SET ACK_Received_DateTime='" & UDT & "' " & _
        '                  " WHERE FK_ReportID ='" & ReportID & "'"
        CMD.CommandText = "usp_tblSendToSwitchAuthority_update_ACK_Received_DateTime"
        CMD.CommandTimeout = 0
        Try
            CMD.Parameters.Clear()
            CMD.Parameters.AddWithValue("@FK_ReportID", ReportID)
            CMD.Parameters.AddWithValue("@DestinationName", DestinationName)
            CMD.Parameters.AddWithValue("@ACK_Received_DateTime", UDT)
            CMD.Parameters.AddWithValue("@FK_ReportTypeID", FK_ReportTypeID)
            CMD.Parameters.AddWithValue("@Wmocode", WMOCode)
            CMD.ExecuteNonQuery()
            Update_tblSendToSwitchQueue = 0
        Catch ex As Exception
            Update_tblSendToSwitchQueue = 1
            WriteLogs("Update_tblSendToSwitchQueue(" & ReportID & "," & UDT & "):" & ex.Message)
        End Try
    End Function
    Public Function GetReportID_By_MSGNO(ByVal yyyyMMddHH As String, ByVal MSGNO As Integer, ByRef diff As Single, ByVal Reptype As Integer) As String
        Dim strRes As String
        Dim CMD As New SqlClient.SqlCommand
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim con As New SqlClient.SqlConnection
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        CMD.CommandType = CommandType.Text
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If

        Try
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = "usp_tblReports_SelectByMSGNOAndType"
            CMD.Parameters.AddWithValue("@MSGNO", MSGNO)
            CMD.Parameters.AddWithValue("@yyyyMMddHH", yyyyMMddHH)
            CMD.Parameters.AddWithValue("@Reptype", Reptype)
            m_DataAdapter.SelectCommand = CMD
            Try
                m_Dataset.Tables("tblReports").Clear()
            Catch ex As Exception

            End Try
            m_DataAdapter.Fill(m_Dataset, "tblReports")
            If m_Dataset.Tables("tblReports").Rows.Count > 0 Then
                GetReportID_By_MSGNO = m_Dataset.Tables("tblReports").Rows(0).Item("ReportID")
                diff = m_Dataset.Tables("tblReports").Rows(0).Item("diff")
            Else
                GetReportID_By_MSGNO = ""
            End If
        Catch ex As Exception
            GetReportID_By_MSGNO = ""
            WriteLogs("CommandText:" & CMD.CommandText)
            WriteLogs("GetReportID_By_MSGNO(" & yyyyMMddHH & "," & MSGNO & "):" & ex.Message)
        End Try
    End Function
    Public Function GetAllDestination(ByRef Destination As DataTable) As Long
        Dim strRes As String
        Dim CMD As New SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        con.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = con
        CMD.CommandType = CommandType.StoredProcedure
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If
        Try
            CMD.CommandText = "usp_tblDestination_selectAll"
            m_DataAdapter.SelectCommand = CMD
            Try
                m_Dataset.Tables("tblDestination").Clear()
            Catch ex As Exception

            End Try
            m_DataAdapter.Fill(m_Dataset, "tblDestination")
            Destination = m_Dataset.Tables("tblDestination")
            GetAllDestination = m_Dataset.Tables("tblDestination").Rows.Count

        Catch ex As Exception
            GetAllDestination = -1
            WriteLogs("GetAllDestination():" & ex.Message)
        End Try
    End Function
#End Region

#Region "Destination-dcs"
    Public Sub InsertIntotblSendToSwithQeue(ByVal tblInsertIntotblSendToSwithQeue As System.Data.DataTable)
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Try
            For i = 0 To tblInsertIntotblSendToSwithQeue.Rows.Count - 1
                CMD.Parameters.Clear()

                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.ConnectionString = My.Settings.ICSDBConnectionString
                CMD.Connection = con
                con.Open()

                CMD.CommandType = CommandType.StoredProcedure
                CMD.CommandText = "usp_tblSendToSwitchQueue_Insert"
                CMD.Parameters.AddWithValue("@FK_SendToSwitchAuthorityID", tblInsertIntotblSendToSwithQeue.Rows(i).Item("FK_SendToSwitchAuthorityID"))
                CMD.Parameters.AddWithValue("@FK_ReportID", tblInsertIntotblSendToSwithQeue.Rows(i).Item("FK_ReportID"))
                CMD.Parameters.AddWithValue("@Sent", tblInsertIntotblSendToSwithQeue.Rows(i).Item("Sent"))

                CMD.ExecuteNonQuery()
            Next

        Catch ex As Exception
            WriteLogs("usp_tblSendToSwitchQueue_Insert" & ex.Message)

        End Try
    End Sub
    Public Sub GetDestinationByWMOCodeAndReportTypeName(ByVal WMOCode As String, ByVal ReportType As String, ByRef tblDestination As System.Data.DataTable)

        'db15
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_GetDestinationInfoByWMOcodeAndReportType"
            CMD.Parameters.AddWithValue("@ReportType", ReportType)
            CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
            Da.SelectCommand = CMD
            Da.Fill(Ds, "tblDes")
            tblDestination = Ds.Tables("tblDes")
        Catch ex As Exception
            WriteLogs("GetDestinationByWMOCodeAndReportTypeName" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Sub

#End Region
#Region "Config"
    Public Function InsertIntotblUser(ByVal tblUser As DataTable) As Long



        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        m_CON.Open()
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure

        CMD.CommandText = "Usp_tblUser_Insert"
        CMD.Parameters.AddWithValue("@UserName", tblUser.Rows(0).Item("UserName"))
        CMD.Parameters.AddWithValue("@UserID", tblUser.Rows(0).Item("UserID"))
        CMD.Parameters.AddWithValue("@UserPassword", tblUser.Rows(0).Item("UserPassword"))
        CMD.Parameters.AddWithValue("@UserEmail", tblUser.Rows(0).Item("UserEmail"))
        CMD.Parameters.AddWithValue("@UserMobile", tblUser.Rows(0).Item("UserMobile"))
        CMD.Parameters.AddWithValue("@UserTel", tblUser.Rows(0).Item("UserTel"))
        CMD.Parameters.AddWithValue("@UserIP", tblUser.Rows(0).Item("UserIP"))
        CMD.Parameters.AddWithValue("@Userport", tblUser.Rows(0).Item("Userport"))
        CMD.Parameters.AddWithValue("@GroupID", tblUser.Rows(0).Item("FK_GroupID"))

        Return CLng(CMD.ExecuteScalar)
        m_CON.Close()
    End Function
    Public Sub UpdatetblUser(ByVal tblUser As DataTable)


        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        m_CON.Open()
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure

        CMD.CommandText = "usp_tbluser_Update"
        CMD.Parameters.AddWithValue("@UserName", tblUser.Rows(0).Item("UserName"))
        CMD.Parameters.AddWithValue("@UserPassword", tblUser.Rows(0).Item("UserPassword"))
        CMD.Parameters.AddWithValue("@UserEmail", tblUser.Rows(0).Item("UserEmail"))
        CMD.Parameters.AddWithValue("@UserMobile", tblUser.Rows(0).Item("UserMobile"))
        CMD.Parameters.AddWithValue("@UserTel", tblUser.Rows(0).Item("UserTel"))
        CMD.Parameters.AddWithValue("@UserIP", tblUser.Rows(0).Item("UserIP"))
        CMD.Parameters.AddWithValue("@Userport", tblUser.Rows(0).Item("Userport"))
        CMD.Parameters.AddWithValue("@UserID", tblUser.Rows(0).Item("UserID"))

        CMD.ExecuteNonQuery()
        m_CON.Close()
    End Sub
    Public Function ISUserExist(ByVal UserID As String) As Boolean
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        m_CON.Open()
        Dim rdr As SqlClient.SqlDataReader
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandText = "usp_ISUserExistForStation"
        CMD.Parameters.AddWithValue("@UserID", UserID)
        rdr = CMD.ExecuteReader
        Try
            If rdr.HasRows Then
                rdr.Read()
                ISUserExist = True
            Else
                ISUserExist = False
            End If
        Catch ex As Exception
            ISUserExist = False
        Finally
            rdr.Close()
            m_CON.Close()
        End Try
    End Function
    Public Function updatetblJoinSensorType_LoggerModule(ByVal tbljoin As DataTable, ByVal JionSentype_LoggerModuleID As Integer) As Long
        'usp_tblStations_DeleteRow
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = m_CON
        m_CON.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandTimeout = 0
        CMD.CommandText = "usp_tblJoinSensorType_LoggerModule_UpdateCenter"
        CMD.Parameters.AddWithValue("@JoinSensorType_LoggerModuleID", tbljoin.Rows(0).Item("JoinSensorType_LoggerModuleID"))
        CMD.Parameters.AddWithValue("@ChannelName", tbljoin.Rows(0).Item("ChannelName"))
        CMD.Parameters.AddWithValue("@SensorAliasName", tbljoin.Rows(0).Item("SensorAliasName"))
        '  CMD.Parameters.AddWithValue("@SensorLocalMinValue", tbljoin.Rows(0).Item("SensorLocalMinValue"))
        ' CMD.Parameters.AddWithValue("@SensorLocalMaxValue", tbljoin.Rows(0).Item("SensorLocalMaxValue"))
        CMD.Parameters.AddWithValue("@SensorValue", tbljoin.Rows(0).Item("SensorValue"))
        CMD.Parameters.AddWithValue("@IsSendingAlarm", tbljoin.Rows(0).Item("IsSendingAlarm"))
        CMD.Parameters.AddWithValue("@LastValueUpdatedDateTime", tbljoin.Rows(0).Item("LastValueUpdatedDateTime"))
        CMD.Parameters.AddWithValue("@FK_SensorTypeID", tbljoin.Rows(0).Item("FK_SensorTypeID"))
        CMD.Parameters.AddWithValue("@FK_LoggerModuleID", tbljoin.Rows(0).Item("FK_LoggerModuleID"))
        CMD.Parameters.AddWithValue("@MonthNum", System.DateTime.Now.Month)

        CMD.ExecuteNonQuery()
        m_CON.Close()
    End Function
    Public Function tblJoinSensorType_LoggerModuleInsert(ByRef tblSensor As DataTable, ByRef InsertID As Long) As Long
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = m_CON
        m_CON.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandTimeout = 0
        '       For i As Integer = 0 To tblSensor.Rows.Count - 1
        Dim i As Integer = tblSensor.Rows.Count - 1
        CMD.CommandText = "usp_tblJoinSensorType_LoggerModule_Insertcenter"
        CMD.Parameters.AddWithValue("@ChannelName", tblSensor.Rows(i).Item("ChannelName"))
        CMD.Parameters.AddWithValue("@SensorAliasName", tblSensor.Rows(i).Item("SensorAliasName"))
        'CMD.Parameters.AddWithValue("@SensorLocalMinValue", tblSensor.Rows(i).Item("SensorLocalMinValue"))
        'CMD.Parameters.AddWithValue("@SensorLocalMaxValue", tblSensor.Rows(i).Item("SensorLocalMaxValue"))
        CMD.Parameters.AddWithValue("@IsSendingAlarm", False)
        CMD.Parameters.AddWithValue("@LastValueUpdatedDateTime", Convert.ToDateTime(tblSensor.Rows(i).Item("LastValueUpdatedDateTime")))
        CMD.Parameters.AddWithValue("@FK_SensorTypeID", tblSensor.Rows(i).Item("FK_SensorTypeID"))
        CMD.Parameters.AddWithValue("@FK_LoggerModuleID", tblSensor.Rows(i).Item("FK_LoggerModuleID"))
        CMD.Parameters.AddWithValue("@SensorValue", tblSensor.Rows(i).Item("SensorValue"))
        '   CMD.Parameters.AddWithValue("@WMOCode", tblSensor.Rows(i).Item("@WMOCode"))

        ' tblmodule.Rows(i).Item("Module") = CLng(CMD.ExecuteScalar)
        InsertID = CLng(CMD.ExecuteScalar)
        m_CON.Close()
    End Function
    Public Function ISSensorExist(ByVal LoggerModuleID As Long, ByVal SensorAliasName As String, ByRef JoinSensorType_LoggerModuleID As Double) As Boolean
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        con.Open()
        CMD.CommandType = CommandType.StoredProcedure
        ISSensorExist = False
        CMD.CommandText = "usp_DoesStationHaveSensor"
        CMD.Parameters.AddWithValue("@FK_LoggerModuleID", LoggerModuleID)
        CMD.Parameters.AddWithValue("@SensorAliasName", SensorAliasName)
        CMD.CommandTimeout = 0
        '  IIf(IsDBNull(CMD.ExecuteScalar()), ISSensorExit = True, ISSensorExit = False)

        Try
            ISSensorExist = True
            JoinSensorType_LoggerModuleID = CMD.ExecuteScalar
            If JoinSensorType_LoggerModuleID = Nothing Then
                ISSensorExist = False
            Else
                ISSensorExist = True

            End If
        Catch ex As Exception
            ISSensorExist = False
        End Try

    End Function
    'Public Function ISModuleExit(ByVal StationID As Long, ByVal loggerModuleName As String, ByVal ModulID As Long) As Boolean
    '    Dim CMD As New SqlClient.SqlCommand
    '    Dim m_CON As New SqlClient.SqlConnection

    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    CMD.Connection = m_CON
    '    m_CON.Open()
    '    CMD.CommandType = CommandType.StoredProcedure
    '    ISModuleExit = False
    '    CMD.CommandText = "Usp_ISModuleExist"
    '    CMD.Parameters.AddWithValue("@StationID", StationID)
    '    CMD.Parameters.AddWithValue("@loggerModuleName", loggerModuleName)
    '    CMD.CommandTimeout = 0
    '    '   IIf(IsDBNull(CMD.ExecuteScalar()), ISModuleExit = True, ISModuleExit = False)

    '    ISModuleExit = True
    '    Try
    '        ModulID = CMD.ExecuteScalar
    '        If CMD.ExecuteScalar = Nothing Then
    '            ISModuleExit = False
    '        Else
    '            ISModuleExit = True
    '        End If

    '    Catch ex As Exception
    '        ISModuleExit = False
    '    End Try
    '    m_CON.Close()
    'End Function
    Public Function ISModuleExit(ByVal ModuleID As Long) As Boolean
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection

        m_CON.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = m_CON
        m_CON.Open()
        CMD.CommandType = CommandType.StoredProcedure
        ISModuleExit = False
        CMD.CommandText = "Usp_ISModuleExist"
        CMD.Parameters.AddWithValue("@ModuleID", ModuleID)

        CMD.CommandTimeout = 0
        '   IIf(IsDBNull(CMD.ExecuteScalar()), ISModuleExit = True, ISModuleExit = False)

        ISModuleExit = True
        Try
            '     ModulID = CMD.ExecuteScalar
            If CMD.ExecuteScalar = Nothing Then
                ISModuleExit = False
            Else
                ISModuleExit = True
            End If

        Catch ex As Exception
            ISModuleExit = False
        End Try
        m_CON.Close()
    End Function
    Public Function ISReportAuto(ByVal reportid As String, ByRef errorlog As String) As Boolean
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        Dim ds As New DataSet
        Dim da As New SqlClient.SqlDataAdapter
        Try

            m_CON.ConnectionString = m_Connection.ConnectionString
            CMD.Connection = m_CON
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure
            ISReportAuto = False
            CMD.CommandTimeout = 0
            CMD.CommandText = "usp_tblReports_SelectRow"
            CMD.Parameters.AddWithValue("@RepID", reportid)


            da.SelectCommand = CMD

            da.Fill(ds, "tblrep")

            ISReportAuto = CBool(ds.Tables("tblrep").Rows(0).Item("Auto"))
        Catch ex As Exception
            '  dt = System.DateTime.Now.AddDays(-1)
            ISReportAuto = True
            errorlog = ex.Message
        Finally
            m_CON.Close()
        End Try

    End Function

    Public Function updatetblLoggerModules(ByVal tblUpdate As DataTable) As Long

        'usp_tblStations_DeleteRow
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = m_CON
        m_CON.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandTimeout = 0
        CMD.CommandText = "usp_tblLoggerModules_Update"
        '  CMD.Parameters.AddWithValue("@LoggerModuleID", SensorID)
        CMD.Parameters.AddWithValue("@LoggerModuleID", tblUpdate.Rows(0).Item("LoggerModuleID"))
        CMD.Parameters.AddWithValue("@LoggerModuleName", tblUpdate.Rows(0).Item("LoggerModuleName"))
        CMD.Parameters.AddWithValue("@FK_AWSVendorID", tblUpdate.Rows(0).Item("FK_AWSVendorID"))
        CMD.Parameters.AddWithValue("@LoggerModuleModel", tblUpdate.Rows(0).Item("LoggerModuleModel"))
        CMD.Parameters.AddWithValue("@SerialNumber", tblUpdate.Rows(0).Item("SerialNumber"))
        CMD.Parameters.AddWithValue("@HardwareVersion", tblUpdate.Rows(0).Item("HardwareVersion"))
        CMD.Parameters.AddWithValue("@SoftwareVersion", tblUpdate.Rows(0).Item("SoftwareVersion"))
        CMD.Parameters.AddWithValue("@ChannelNumber", tblUpdate.Rows(0).Item("ChannelNumber"))
        '  CMD.Parameters.AddWithValue("@FK_LoggerModuleID", tblUpdate.Rows(0).Item("FK_LoggerModuleID"))
        CMD.Parameters.AddWithValue("@FK_StationID", tblUpdate.Rows(0).Item("FK_StationID"))
        CMD.ExecuteNonQuery()
        m_CON.Close()
    End Function
    Public Function DefineLoggerModules_For_Station(ByRef tblModules As DataTable, ByVal WMOCode As String)
        Dim CMD As New SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        con.ConnectionString = My.Settings.ICSDBConnectionString
        con.Open()
        CMD.Connection = con
        CMD.CommandType = CommandType.StoredProcedure

        Try
            CMD.CommandText = "usp_LoggerModuleDefenition"
            For i As Integer = 0 To tblModules.Rows.Count - 1
                CMD.Parameters.Clear()
                CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
                CMD.Parameters.AddWithValue("@LoggerModuleID", tblModules.Rows(i).Item("LoggerModuleID"))
                CMD.Parameters.AddWithValue("@LoggerModuleName", tblModules.Rows(i).Item("LoggerModuleName"))
                CMD.Parameters.AddWithValue("@LoggerModuleModel", tblModules.Rows(i).Item("LoggerModuleModel"))
                CMD.Parameters.AddWithValue("@VendorName", tblModules.Rows(i).Item("VendorName"))
                CMD.Parameters.AddWithValue("@SerialNumber", tblModules.Rows(i).Item("SerialNumber"))
                CMD.Parameters.AddWithValue("@HardwareVersion", tblModules.Rows(i).Item("HardwareVersion"))
                CMD.Parameters.AddWithValue("@SoftwareVersion", tblModules.Rows(i).Item("SoftwareVersion"))
                CMD.Parameters.AddWithValue("@ChannelNumber", tblModules.Rows(i).Item("ChannelNumber"))
                Try
                    CMD.ExecuteNonQuery()

                Catch ex As Exception
                    WriteLogs("usp_LoggerModuleDefenition" & ex.Message)
                End Try
            Next i
        Catch ex As Exception
            Return 2
            ' RaiseEvent DataBaseErrors("GetSensors_Of_Station " & ex.Message)
        End Try
        Return 0
    End Function

    Public Function InSertInToTblStation(ByVal tblStation As DataTable) As Long
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        Dim CMD As New SqlClient.SqlCommand
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandTimeout = 0
        m_CON.Open()
        CMD.CommandText = "usp_tblStations_Insert"
        CMD.Parameters.AddWithValue("@WMOCODE", tblStation.Rows(0).Item("WMOCODE"))
        CMD.Parameters.AddWithValue("@IKAOCode", tblStation.Rows(0).Item("IKAOCode"))
        CMD.Parameters.AddWithValue("@StationName", tblStation.Rows(0).Item("StationName"))
        CMD.Parameters.AddWithValue("@GPSLongitude", tblStation.Rows(0).Item("GPSLongitude"))
        CMD.Parameters.AddWithValue("@Password", tblStation.Rows(0).Item("Password"))
        CMD.Parameters.AddWithValue("@GPSLatitude", tblStation.Rows(0).Item("GPSLatitude"))
        CMD.Parameters.AddWithValue("@GPSHeight", tblStation.Rows(0).Item("GPSHeight"))
        CMD.Parameters.AddWithValue("@Address", tblStation.Rows(0).Item("Address"))
        CMD.Parameters.AddWithValue("@PhoneNo", tblStation.Rows(0).Item("PhoneNo"))
        CMD.Parameters.AddWithValue("@CellPhone", tblStation.Rows(0).Item("CellPhone"))
        CMD.Parameters.AddWithValue("@IP", tblStation.Rows(0).Item("IP"))
        CMD.Parameters.AddWithValue("@Port", tblStation.Rows(0).Item("Port"))
        CMD.Parameters.AddWithValue("@AdminName", tblStation.Rows(0).Item("AdminName"))
        CMD.Parameters.AddWithValue("@AdminPhone", tblStation.Rows(0).Item("AdminPhone"))
        CMD.Parameters.AddWithValue("@DataFolderAddress", tblStation.Rows(0).Item("DataFolderAddress"))
        CMD.Parameters.AddWithValue("@XCoordination", tblStation.Rows(0).Item("XCoordination"))
        CMD.Parameters.AddWithValue("@YCoordination", tblStation.Rows(0).Item("YCoordination"))
        CMD.Parameters.AddWithValue("@FK_RegionID", tblStation.Rows(0).Item("FK_RegionID"))
        CMD.Parameters.AddWithValue("@FK_TypeOfCommunicationID", tblStation.Rows(0).Item("FK_TypeOfCommunicationID"))
        CMD.Parameters.AddWithValue("@IsSendingStandardReports", tblStation.Rows(0).Item("IsSendingSynopMetar"))
        CMD.Parameters.AddWithValue("@IssendingMetaData", tblStation.Rows(0).Item("IssendingMetaData"))
        CMD.Parameters.AddWithValue("@IsFirstConfigChanged", tblStation.Rows(0).Item("IsFirstConfigChanged"))
        CMD.Parameters.AddWithValue("@Comment", tblStation.Rows(0).Item("Comment"))
        CMD.Parameters.AddWithValue("@PicAdd", tblStation.Rows(0).Item("PicAdd"))
        CMD.Parameters.AddWithValue("@FK_StationTypeID", tblStation.Rows(0).Item("FK_StationTypeID"))
        CMD.Parameters.AddWithValue("@Active", tblStation.Rows(0).Item("active"))
        Return CLng(CMD.ExecuteScalar)
        m_CON.Close()
    End Function
    Public Function usp_tblStations_Update(ByVal tblStation As DataTable, ByVal StationID As Long) As Long

        'usp_tblStations_DeleteRow
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = m_CON
        m_CON.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandTimeout = 0
        CMD.CommandText = "usp_tblStations_UpdateforStation"
        CMD.Parameters.AddWithValue("@stationID", StationID)
        CMD.Parameters.AddWithValue("@WMOCODE", tblStation.Rows(0).Item("WMOCODE"))
        CMD.Parameters.AddWithValue("@IKAOCode", tblStation.Rows(0).Item("IKAOCode"))
        CMD.Parameters.AddWithValue("@StationName", tblStation.Rows(0).Item("StationName"))
        CMD.Parameters.AddWithValue("@GPSLongitude", tblStation.Rows(0).Item("GPSLongitude"))
        CMD.Parameters.AddWithValue("@GPSLatitude", tblStation.Rows(0).Item("GPSLatitude"))
        CMD.Parameters.AddWithValue("@GPSHeight", tblStation.Rows(0).Item("GPSHeight"))
        CMD.Parameters.AddWithValue("@Password", tblStation.Rows(0).Item("Password"))
        CMD.Parameters.AddWithValue("@Address", tblStation.Rows(0).Item("Address"))
        CMD.Parameters.AddWithValue("@PhoneNo", tblStation.Rows(0).Item("PhoneNo"))
        CMD.Parameters.AddWithValue("@CellPhone", tblStation.Rows(0).Item("CellPhone"))
        CMD.Parameters.AddWithValue("@IP", tblStation.Rows(0).Item("IP"))
        CMD.Parameters.AddWithValue("@Port", tblStation.Rows(0).Item("Port"))
        CMD.Parameters.AddWithValue("@AdminName", tblStation.Rows(0).Item("AdminName"))
        CMD.Parameters.AddWithValue("@AdminPhone", tblStation.Rows(0).Item("AdminPhone"))
        CMD.Parameters.AddWithValue("@DataFolderAddress", tblStation.Rows(0).Item("DataFolderAddress"))
        'CMD.Parameters.AddWithValue("@XCoordination", tblStation.Rows(0).Item("XCoordination"))
        'CMD.Parameters.AddWithValue("@YCoordination", tblStation.Rows(0).Item("YCoordination"))
        CMD.Parameters.AddWithValue("@FK_RegionID", tblStation.Rows(0).Item("FK_RegionID"))
        CMD.Parameters.AddWithValue("@FK_TypeOfCommunicationID", tblStation.Rows(0).Item("FK_TypeOfCommunicationID"))
        'CMD.Parameters.AddWithValue("@IsSendingSynopMetar", tblStation.Rows(0).Item("IsSendingSynopMetar"))
        'CMD.Parameters.AddWithValue("@IssendingMetaData", tblStation.Rows(0).Item("IssendingMetaData"))
        CMD.Parameters.AddWithValue("@IsFirstConfigChanged", tblStation.Rows(0).Item("IsFirstConfigChanged"))
        CMD.Parameters.AddWithValue("@Comment", tblStation.Rows(0).Item("Comment"))
        CMD.Parameters.AddWithValue("@PicAdd", tblStation.Rows(0).Item("PicAdd"))
        CMD.Parameters.AddWithValue("@FK_StationTypeID", tblStation.Rows(0).Item("FK_StationTypeID"))
        CMD.Parameters.AddWithValue("@Active", tblStation.Rows(0).Item("Active"))
        CMD.ExecuteNonQuery()
        m_CON.Close()
    End Function
    Public Function GetRegionCodeByName(ByVal RegionName As String) As Integer
        GetRegionCodeByName = 0
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        con.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandText = "usp_tblRegions_SelectRowByName"
        Dim ds As New DataSet
        Dim da As New SqlClient.SqlDataAdapter
        da.SelectCommand = CMD
        CMD.Parameters.AddWithValue("@RegionName", RegionName)
        da.Fill(ds, "tblRegion")
        Try
            GetRegionCodeByName = CLng(ds.Tables(0).Rows(0).Item("RegionID"))
        Catch ex As Exception
            '  dt = System.DateTime.Now.AddDays(-1)
            GetRegionCodeByName = 0
        End Try
    End Function
    Public Function GetGroupIDbyName(ByVal GroupName As String) As Integer
        GetGroupIDbyName = 0
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        con.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandText = "usp_tblGroup_SelectByGroupName"
        Dim ds As New DataSet
        Dim da As New SqlClient.SqlDataAdapter
        da.SelectCommand = CMD
        CMD.Parameters.AddWithValue("@GroupName", GroupName)
        da.Fill(ds, "tblg")
        Try
            GetGroupIDbyName = CLng(ds.Tables(0).Rows(0).Item("GroupID"))
        Catch ex As Exception
            '  dt = System.DateTime.Now.AddDays(-1)
            GetGroupIDbyName = 0
        End Try
    End Function
    'usp_GetStationIDByWMOCODE
    'Public Function GetStationIDByWMOcode(ByVal WMOCODE As String) As Long
    '    Dim CMD As New SqlClient.SqlCommand
    '    Dim m_CON As New SqlClient.SqlConnection
    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    CMD.Connection = m_CON
    '    m_CON.Open()
    '    CMD.CommandType = CommandType.StoredProcedure
    '    GetStationIDByWMOcode = 0
    '    CMD.CommandTimeout = 0
    '    CMD.CommandText = "usp_GetStationNameByWMOCODE"
    '    CMD.Parameters.AddWithValue("@WMOCODE", WMOCODE)


    '    '   IIf(IsDBNull(CMD.ExecuteScalar()), IsExistStationWMOCODE = False, IsExistStationWMOCODE = True)
    '    Try
    '        GetStationIDByWMOcode = CLng(CMD.ExecuteScalar)
    '    Catch ex As Exception

    '    End Try


    '    m_CON.Close()
    'End Function
    Public Function GetStationIDBYWMOCode(ByVal WMOCode As String) As Long
        GetStationIDBYWMOCode = 0
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        con.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandText = "usp_tblStations_SelectRowByWMOCode"
        Dim ds As New DataSet
        Dim da As New SqlClient.SqlDataAdapter
        da.SelectCommand = CMD
        CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
        da.Fill(ds, "tblStation")
        Try
            GetStationIDBYWMOCode = ds.Tables(0).Rows(0).Item("StationID")
        Catch ex As Exception
            '  dt = System.DateTime.Now.AddDays(-1)
            GetStationIDBYWMOCode = 0
        End Try
    End Function
    'Public Function IsExistStationWMOCODE(ByVal WMOCODE As String, ByVal StationName As String) As Boolean
    '    Dim CMD As New SqlClient.SqlCommand
    '    Dim m_CON As New SqlClient.SqlConnection
    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    CMD.Connection = m_CON
    '    m_CON.Open()
    '    CMD.CommandType = CommandType.StoredProcedure
    '    IsExistStationWMOCODE = False
    '    CMD.CommandTimeout = 0
    '    CMD.CommandText = "usp_GetStationNameByWMOCODE"
    '    CMD.Parameters.AddWithValue("@WMOCODE", WMOCODE)
    '    CMD.Parameters.AddWithValue("@stationName", StationName)

    '    '   IIf(IsDBNull(CMD.ExecuteScalar()), IsExistStationWMOCODE = False, IsExistStationWMOCODE = True)
    '    If CMD.ExecuteScalar = Nothing Then
    '        IsExistStationWMOCODE = False
    '    Else
    '        IsExistStationWMOCODE = True
    '    End If
    '    m_CON.Close()
    'End Function
    Public Function IsExistStationWMOCODE(ByVal WMOCODE As String, ByVal StationName As String) As Boolean
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = m_CON
        m_CON.Open()
        CMD.CommandType = CommandType.StoredProcedure
        IsExistStationWMOCODE = False
        CMD.CommandTimeout = 0
        CMD.CommandText = "usp_GetStationNameByWMOCODE"
        CMD.Parameters.AddWithValue("@WMOCODE", WMOCODE)
        CMD.Parameters.AddWithValue("@stationName", StationName)

        '   IIf(IsDBNull(CMD.ExecuteScalar()), IsExistStationWMOCODE = False, IsExistStationWMOCODE = True)
        If CMD.ExecuteScalar = Nothing Then
            IsExistStationWMOCODE = False
        Else
            IsExistStationWMOCODE = True
        End If
        m_CON.Close()
    End Function
    Public Function GetVendorIDBYName(ByVal VendorName As String) As Long
        GetVendorIDBYName = 0
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        con.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandText = "usp_tblAWSVendors_SelectRowByName"
        Dim ds As New DataSet
        Dim da As New SqlClient.SqlDataAdapter
        da.SelectCommand = CMD
        CMD.Parameters.AddWithValue("@VendorName", VendorName)
        da.Fill(ds, "tblVendor")
        Try
            GetVendorIDBYName = CLng(ds.Tables(0).Rows(0).Item("AWSVendorID"))
        Catch ex As Exception
            '  dt = System.DateTime.Now.AddDays(-1)
            GetVendorIDBYName = 0
        End Try
    End Function
    'usp_tblAWSVendors_SelectRowByName
#End Region

    Public Function ConfigChange(ByVal SpName As String, ByVal parameters As List(Of String))
        Try


            Dim con As New SqlClient.SqlConnection
            Dim CMD As New SqlClient.SqlCommand
            con.ConnectionString = My.Settings.ICSDBConnectionString
            CMD.Connection = con
            con.Open()
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = SpName
            CMD.Parameters.Clear()
            For Each p In parameters
                Try


                    Dim arr() = Split(p, "=")
                    If arr(1) = "?" Then
                        arr(1) = ""
                    End If
                    CMD.Parameters.AddWithValue(arr(0), arr(1))
                Catch ex As Exception

                End Try
            Next
            CMD.ExecuteNonQuery()
        Catch ex As Exception
            WriteLogs(SpName & ex.Message)
        End Try


    End Function


    Public Function IsDateExist(ByVal dt As String, ByVal WmoCode As String) As Boolean
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        con.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandText = "usp_tblDaily_IsDayExist"
        CMD.Parameters.AddWithValue("@strDate", dt)
        CMD.Parameters.AddWithValue("@WMOCode", WmoCode)

        Try
            Dim count As Integer
            count = CInt(CMD.ExecuteScalar)
            ' WriteLogs("IsDateExist count" & count)
            '     WriteLogs("dt " & dt)
            If count > 0 Then
                IsDateExist = True
                WriteLogs("IsDateExist" & IsDateExist)
            Else
                IsDateExist = False
                WriteLogs("IsDateExist" & IsDateExist)
            End If

        Catch ex As Exception
            '  dt = System.DateTime.Now.AddDays(-1)
            IsDateExist = False
            WriteLogs("GetlastDateThatIsCompleteIntblDaily" & ex.Message)
        End Try
    End Function
    Public Function InsertINtotblDaily(ByVal tblInsert As DataTable) As Boolean
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Try
            For i = 0 To tblInsert.Rows.Count - 1
                CMD.Parameters.Clear()

                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.ConnectionString = My.Settings.ICSDBConnectionString
                CMD.Connection = con
                con.Open()
                CMD.CommandTimeout = 0
                CMD.CommandType = CommandType.StoredProcedure
                CMD.CommandText = "usp_tblDaily_Insert"
                CMD.Parameters.AddWithValue("@WMOCode", tblInsert.Rows(i).Item("WMOCode"))
                CMD.Parameters.AddWithValue("@DateTimeMilady", tblInsert.Rows(i).Item("DateTimeMilady"))
                CMD.Parameters.AddWithValue("@AirTemperatureMax", tblInsert.Rows(i).Item("AirTemperatureMax"))
                CMD.Parameters.AddWithValue("@AirTemperatureMin", tblInsert.Rows(i).Item("AirTemperatureMin"))
                CMD.Parameters.AddWithValue("@AirTemperatureAvg", tblInsert.Rows(i).Item("AirTemperatureAvg"))
                CMD.Parameters.AddWithValue("@HumidityMin", tblInsert.Rows(i).Item("HumidityMin"))
                CMD.Parameters.AddWithValue("@HumidityMax", tblInsert.Rows(i).Item("HumidityMax"))
                CMD.Parameters.AddWithValue("@HumidityAvg", tblInsert.Rows(i).Item("HumidityAvg"))
                CMD.Parameters.AddWithValue("@RainSum", tblInsert.Rows(i).Item("RainSum"))
                CMD.Parameters.AddWithValue("@Evaporation", tblInsert.Rows(i).Item("Evaporation"))
                CMD.Parameters.AddWithValue("@SunShine", tblInsert.Rows(i).Item("SunShine"))
                CMD.Parameters.AddWithValue("@WSMax", tblInsert.Rows(i).Item("WSMax"))
                CMD.Parameters.AddWithValue("@WDMax", tblInsert.Rows(i).Item("WDMax"))
                CMD.Parameters.AddWithValue("@SurfaceTempMin", tblInsert.Rows(i).Item("SurfaceTempMin"))
                CMD.Parameters.AddWithValue("@SurfaceTempMax", tblInsert.Rows(i).Item("SurfaceTempMax"))
                CMD.Parameters.AddWithValue("@Humidity03", tblInsert.Rows(i).Item("Humidity03"))
                CMD.Parameters.AddWithValue("@Humidity09", tblInsert.Rows(i).Item("Humidity09"))
                CMD.Parameters.AddWithValue("@Humidity15", tblInsert.Rows(i).Item("Humidity15"))
                CMD.Parameters.AddWithValue("@Temperature03", tblInsert.Rows(i).Item("Temperature03"))
                CMD.Parameters.AddWithValue("@Temperature09", tblInsert.Rows(i).Item("Temperature09"))
                CMD.Parameters.AddWithValue("@Temperature15", tblInsert.Rows(i).Item("Temperature15"))
                CMD.Parameters.AddWithValue("@AirPressure03", tblInsert.Rows(i).Item("AirPressure03"))
                CMD.Parameters.AddWithValue("@AirPressure09", tblInsert.Rows(i).Item("AirPressure09"))
                CMD.Parameters.AddWithValue("@AirPressure15", tblInsert.Rows(i).Item("AirPressure15"))
                CMD.Parameters.AddWithValue("@QFF03", tblInsert.Rows(i).Item("QFF03"))
                CMD.Parameters.AddWithValue("@QFF09", tblInsert.Rows(i).Item("QFF09"))
                CMD.Parameters.AddWithValue("@QFF15", tblInsert.Rows(i).Item("QFF15"))
                CMD.Parameters.AddWithValue("@QFFAvg", tblInsert.Rows(i).Item("QFFAvg"))
                CMD.Parameters.AddWithValue("@ISComplete", tblInsert.Rows(i).Item("ISComplete"))
                CMD.Parameters.AddWithValue("@AirPressureAvg", tblInsert.Rows(i).Item("AirPressureAvg"))
                CMD.Parameters.AddWithValue("@WetTemp03", tblInsert.Rows(i).Item("WetTemp03"))
                CMD.Parameters.AddWithValue("@WetTemp09", tblInsert.Rows(i).Item("WetTemp09"))
                CMD.Parameters.AddWithValue("@WetTemp15", tblInsert.Rows(i).Item("WetTemp15"))
                CMD.Parameters.AddWithValue("@DewPoint03", tblInsert.Rows(i).Item("DewPoint03"))
                CMD.Parameters.AddWithValue("@DewPoint09", tblInsert.Rows(i).Item("DewPoint09"))
                CMD.Parameters.AddWithValue("@DewPoint15", tblInsert.Rows(i).Item("DewPoint15"))
                CMD.Parameters.AddWithValue("@SoilTemp5", tblInsert.Rows(i).Item("SoilTemp5"))
                CMD.Parameters.AddWithValue("@SoilTemp10", tblInsert.Rows(i).Item("SoilTemp10"))
                CMD.Parameters.AddWithValue("@SoilTemp20", tblInsert.Rows(i).Item("SoilTemp20"))
                CMD.Parameters.AddWithValue("@SoilTemp50", tblInsert.Rows(i).Item("SoilTemp50"))
                CMD.Parameters.AddWithValue("@SoilTemp100", tblInsert.Rows(i).Item("SoilTemp100"))

                CMD.Parameters.AddWithValue("@SoilTemp30", tblInsert.Rows(i).Item("SoilTemp30"))
                CMD.Parameters.AddWithValue("@WetTempAvg", tblInsert.Rows(i).Item("WetTempAvg"))
                CMD.Parameters.AddWithValue("@DEWPOINTAvg", tblInsert.Rows(i).Item("DEWPOINTAvg"))
                CMD.Parameters.AddWithValue("@Radiation", tblInsert.Rows(i).Item("Radiation"))
                CMD.Parameters.AddWithValue("@WaperPressure", tblInsert.Rows(i).Item("WaperPressure"))
                CMD.Parameters.AddWithValue("@WDPrevailing", tblInsert.Rows(i).Item("WDPrevailing"))
                CMD.Parameters.AddWithValue("@WSPrevailing", tblInsert.Rows(i).Item("WSPrevailing"))
                ' CMD.Parameters.AddWithValue("@LastUpdateTime", tblInsert.Rows(i).Item("LastUpdateTime"))
                '    LastUpdateTime()


                CMD.ExecuteNonQuery()
            Next
            InsertINtotblDaily = True
        Catch ex As Exception
            WriteLogs("InsertINtotblDaily" & ex.Message)
            InsertINtotblDaily = False
        End Try
    End Function
    Public Function Update_tblDaily(ByVal tblInsert As DataTable) As Boolean
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Try
            For i = 0 To tblInsert.Rows.Count - 1
                CMD.Parameters.Clear()

                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.ConnectionString = My.Settings.ICSDBConnectionString
                CMD.Connection = con
                con.Open()
                CMD.CommandType = CommandType.StoredProcedure
                CMD.CommandText = "usp_tblDaily_UpdateTable_Center"
                CMD.Parameters.AddWithValue("@WMOCode", tblInsert.Rows(i).Item("WMOCode"))
                CMD.Parameters.AddWithValue("@DateTimeMilady", tblInsert.Rows(i).Item("DateTimeMilady"))
                CMD.Parameters.AddWithValue("@AirTemperatureMax", tblInsert.Rows(i).Item("AirTemperatureMax"))
                CMD.Parameters.AddWithValue("@AirTemperatureMin", tblInsert.Rows(i).Item("AirTemperatureMin"))
                CMD.Parameters.AddWithValue("@AirTemperatureAvg", tblInsert.Rows(i).Item("AirTemperatureAvg"))
                CMD.Parameters.AddWithValue("@HumidityMin", tblInsert.Rows(i).Item("HumidityMin"))
                CMD.Parameters.AddWithValue("@HumidityMax", tblInsert.Rows(i).Item("HumidityMax"))
                CMD.Parameters.AddWithValue("@HumidityAvg", tblInsert.Rows(i).Item("HumidityAvg"))
                CMD.Parameters.AddWithValue("@RainSum", tblInsert.Rows(i).Item("RainSum"))
                CMD.Parameters.AddWithValue("@Evaporation", tblInsert.Rows(i).Item("Evaporation"))
                CMD.Parameters.AddWithValue("@SunShine", tblInsert.Rows(i).Item("SunShine"))
                CMD.Parameters.AddWithValue("@WSMax", tblInsert.Rows(i).Item("WSMax"))
                CMD.Parameters.AddWithValue("@WDMax", tblInsert.Rows(i).Item("WDMax"))
                CMD.Parameters.AddWithValue("@SurfaceTempMin", tblInsert.Rows(i).Item("SurfaceTempMin"))
                CMD.Parameters.AddWithValue("@SurfaceTempMax", tblInsert.Rows(i).Item("SurfaceTempMax"))
                CMD.Parameters.AddWithValue("@Humidity03", tblInsert.Rows(i).Item("Humidity03"))
                CMD.Parameters.AddWithValue("@Humidity09", tblInsert.Rows(i).Item("Humidity09"))
                CMD.Parameters.AddWithValue("@Humidity15", tblInsert.Rows(i).Item("Humidity15"))
                CMD.Parameters.AddWithValue("@Temperature03", tblInsert.Rows(i).Item("Temperature03"))
                CMD.Parameters.AddWithValue("@Temperature09", tblInsert.Rows(i).Item("Temperature09"))
                CMD.Parameters.AddWithValue("@Temperature15", tblInsert.Rows(i).Item("Temperature15"))
                CMD.Parameters.AddWithValue("@AirPressure03", tblInsert.Rows(i).Item("AirPressure03"))
                CMD.Parameters.AddWithValue("@AirPressure09", tblInsert.Rows(i).Item("AirPressure09"))
                CMD.Parameters.AddWithValue("@AirPressure15", tblInsert.Rows(i).Item("AirPressure15"))
                CMD.Parameters.AddWithValue("@QFF03", tblInsert.Rows(i).Item("QFF03"))
                CMD.Parameters.AddWithValue("@QFF09", tblInsert.Rows(i).Item("QFF09"))
                CMD.Parameters.AddWithValue("@QFF15", tblInsert.Rows(i).Item("QFF15"))
                CMD.Parameters.AddWithValue("@QFFAvg", tblInsert.Rows(i).Item("QFFAvg"))
                CMD.Parameters.AddWithValue("@ISComplete", tblInsert.Rows(i).Item("ISComplete"))
                CMD.Parameters.AddWithValue("@AirPressureAvg", tblInsert.Rows(i).Item("AirPressureAvg"))
                CMD.Parameters.AddWithValue("@WetTemp03", tblInsert.Rows(i).Item("WetTemp03"))
                CMD.Parameters.AddWithValue("@WetTemp09", tblInsert.Rows(i).Item("WetTemp09"))
                CMD.Parameters.AddWithValue("@WetTemp15", tblInsert.Rows(i).Item("WetTemp15"))
                CMD.Parameters.AddWithValue("@DewPoint03", tblInsert.Rows(i).Item("DewPoint03"))
                CMD.Parameters.AddWithValue("@DewPoint09", tblInsert.Rows(i).Item("DewPoint09"))
                CMD.Parameters.AddWithValue("@DewPoint15", tblInsert.Rows(i).Item("DewPoint15"))
                CMD.Parameters.AddWithValue("@SoilTemp5", tblInsert.Rows(i).Item("SoilTemp5"))
                CMD.Parameters.AddWithValue("@SoilTemp10", tblInsert.Rows(i).Item("SoilTemp10"))
                CMD.Parameters.AddWithValue("@SoilTemp20", tblInsert.Rows(i).Item("SoilTemp20"))
                CMD.Parameters.AddWithValue("@SoilTemp50", tblInsert.Rows(i).Item("SoilTemp50"))
                CMD.Parameters.AddWithValue("@SoilTemp100", tblInsert.Rows(i).Item("SoilTemp100"))
                CMD.Parameters.AddWithValue("@SoilTemp30", tblInsert.Rows(i).Item("SoilTemp30"))
                CMD.Parameters.AddWithValue("@WetTempAvg", tblInsert.Rows(i).Item("WetTempAvg"))
                CMD.Parameters.AddWithValue("@DEWPOINTAvg", tblInsert.Rows(i).Item("DEWPOINTAvg"))
                CMD.Parameters.AddWithValue("@Radiation", tblInsert.Rows(i).Item("Radiation"))
                CMD.Parameters.AddWithValue("@WaperPressure", tblInsert.Rows(i).Item("WaperPressure"))
                CMD.Parameters.AddWithValue("@WDPrevailing", tblInsert.Rows(i).Item("WDPrevailing"))
                CMD.Parameters.AddWithValue("@WSPrevailing", tblInsert.Rows(i).Item("WSPrevailing"))
                '   CMD.Parameters.AddWithValue("@LastUpdateTime", tblInsert.Rows(i).Item("LastUpdateTime"))
                '    LastUpdateTime()


                CMD.ExecuteNonQuery()
            Next i
            Update_tblDaily = True
        Catch ex As Exception
            WriteLogs("Update_tblDaily" & ex.Message)
            Update_tblDaily = False
        End Try
    End Function
    Public Function SendMetaDataToSwich(ByVal WMOCode As Long) As Boolean
        'db1
        Dim DA As New SqlClient.SqlDataAdapter
        Dim CMD As New SqlClient.SqlCommand
        Dim DS As New DataSet
        Dim con As New SqlClient.SqlConnection
        ' Dim DT As New DataTable

        con.ConnectionString = My.Settings.ICSDBConnectionString

        Try
            CMD.Connection = con
            CMD.CommandType = CommandType.StoredProcedure
            CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
            DS.Clear()
            CMD.CommandText = "ISSendMetaData"
            DA.SelectCommand = CMD
            DA.Fill(DS, "Response")
            If DS.Tables("Response").Rows.Count > 0 And Not IsDBNull(DS.Tables("Response").Rows(0).Item("IssendingMetaData")) Then
                If DS.Tables("Response").Rows(0).Item("IssendingMetaData") Then
                    SendMetaDataToSwich = True
                Else
                    SendMetaDataToSwich = False
                End If
            End If
        Catch ex As Exception
            SendMetaDataToSwich = False
            WriteLogs("SendMetaDataToSwich" & ex.Message)
        End Try
    End Function
    Public Function GetStationSTSByCode(ByVal WMOCode As Long) As Boolean

        ''db7
        'Dim mycon As New SqlClient.SqlConnection
        'Dim DA As New SqlClient.SqlDataAdapter
        'Dim CMD As New SqlClient.SqlCommand
        'Dim DS As New DataSet
        '' Dim DT As New DataTable
        ''  mycon = CON
        'mycon.ConnectionString = My.Settings.ICSDBConnectionString

        'Try
        '    CMD.Connection = mycon
        '    CMD.CommandType = CommandType.StoredProcedure
        '    DS.Clear()
        '    CMD.CommandText = "ISSendSynopMetar"
        '    CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
        '    DA.SelectCommand = CMD
        '    DA.Fill(DS, "Response")
        '    If DS.Tables("Response").Rows.Count > 0 AndAlso DS.Tables("Response").Rows(0).Item("IsSendingStandardReports") Then
        '        GetStationSTSByCode = True
        '    Else
        GetStationSTSByCode = False
        '    End If
        'Catch ex As Exception
        '    GetStationSTSByCode = False
        '    WriteLogs("GetStationSTSByCode" & ex.Message)
        'End Try
    End Function
    'Public Function GetStationCodeByName(ByVal StationName As String) As String
    '    Dim m_CON As New SqlClient.SqlConnection
    '    Dim CMD As New SqlClient.SqlCommand
    '    Try
    '        m_CON.ConnectionString = m_Connection.ConnectionString
    '        CMD.Connection = m_CON
    '        CMD.CommandType = CommandType.StoredProcedure
    '        GetStationCodeByName = ""
    '        CMD.CommandText = "GetStationCodeByName"
    '        CMD.Parameters.AddWithValue("@Name", StationName)
    '        GetStationCodeByName = CMD.ExecuteScalar()
    '    Catch ex As Exception
    '        WriteLogs("GetStationCodeByName" & ex.Message)
    '    End Try
    'End Function


    Public Function GetStationsWMOcode(ByRef Station As List(Of String)) As String
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try
            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure
            GetStationsWMOcode = ""
            CMD.CommandText = "usp_tblStations_SelectAll"
            Da.SelectCommand = CMD
            Da.Fill(Ds, "tblStations")
            For i = 0 To Ds.Tables("tblStations").Rows.Count - 1
                Station.Add(Ds.Tables(0).Rows(i).Item("WMOCode"))
            Next
        Catch ex As Exception
            WriteLogs("GetStationsWMOcode" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function



    Public Function GetReports(ByRef tblReports As DataTable)
        'db13
        WriteLogs("GetReports start")
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandTimeout = 120000
            CMD.CommandText = "usp_GetSentToSwitchQueueInfo"
            Da.SelectCommand = CMD
            Da.Fill(Ds, "tblReports")
            tblReports = Ds.Tables("tblReports")
            WriteLogs("GetReports end tblReports.count " & tblReports.Rows.Count - 1)
        Catch ex As Exception
            WriteLogs("GetReports" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    Public Function GetBulletins(ByRef tblReports As DataTable)
        'db13
        WriteLogs("GetBulletinS Start ")
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandTimeout = 120000
            CMD.CommandText = "usp_GetBulletinsSentToSwitchQueueInfo"
            Da.SelectCommand = CMD
            Da.Fill(Ds, "tblReports")
            tblReports = Ds.Tables("tblReports")
            WriteLogs("GetBulletinS end Bulletins.count " & tblReports.Rows.Count - 1)
        Catch ex As Exception
            WriteLogs("GetBulletins " & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function


    Public Function GetMeGetProvinceIDByWMOcode(ByRef Wmocode As String) As String
        'db13
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_GetMeGetProvinceIDByWMOcode"
            CMD.Parameters.AddWithValue("@WMOCODE", Wmocode)
            Da.SelectCommand = CMD
            Da.Fill(Ds, "usp_GetMeGetProvinceIDByWMOcode")
            GetMeGetProvinceIDByWMOcode = Ds.Tables("usp_GetMeGetProvinceIDByWMOcode").Rows(0).Item("ProvinceCode")
        Catch ex As Exception
            WriteLogs("usp_GetMeGetProvinceIDByWMOcode" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    'ISSendReportToSwitch
    Public Function ISSendReportToSwitch(ByVal SentoSwitchQueueID As Long) As Boolean
        'db13
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_tblSendToSwitchQueue_ISSend"
            CMD.Parameters.AddWithValue("@SentoSwitchQueueID", SentoSwitchQueueID)
            Da.SelectCommand = CMD
            Da.Fill(Ds, "usp_tblSendToSwitchQueue_ISSend")
            If Ds.Tables("usp_tblSendToSwitchQueue_ISSend").Rows.Count > 0 Then
                ISSendReportToSwitch = True
            Else
                ISSendReportToSwitch = False
            End If
        Catch ex As Exception
            ISSendReportToSwitch = False
            WriteLogs("ISSendReportToSwitch" & ex.Message)
        End Try

    End Function
    Public Function ISSendBulletinToSwitch(ByVal BulletinQueueID As Long) As Boolean
        'db13
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_tblBulletinSendtoSwitchQueue_ISSend"
            CMD.Parameters.AddWithValue("@BulletinQueueID", BulletinQueueID)
            Da.SelectCommand = CMD
            Da.Fill(Ds, "usp_tblSendToSwitchQueue_ISSend")
            If Ds.Tables("usp_tblSendToSwitchQueue_ISSend").Rows.Count > 0 Then
                ISSendBulletinToSwitch = True
            Else
                ISSendBulletinToSwitch = False
            End If
        Catch ex As Exception
            ISSendBulletinToSwitch = False
            WriteLogs("ISSendBulletinToSwitch" & ex.Message)
        End Try

    End Function
    Public Function UpdatetblSentToSwitchQueueByID(ByVal ID As Long, ByVal sent As Byte)
        'db14
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand


        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_tblSendToSwitchQueue_UpdateSendToSwitch"
            CMD.Parameters.AddWithValue("@SentoSwitchQueueID", ID)
            CMD.Parameters.AddWithValue("@Sent", sent)
            CMD.ExecuteNonQuery()

        Catch ex As Exception
            WriteLogs("UpdatetblSentToSwitchQueueByID" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    Public Function UpdatetblBulletinSendtoSwitchQueueByID(ByVal ID As Long, ByVal sent As Byte)
        'db14
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand


        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_tblBulletinSendtoSwitchQueue_UpdateSendToSwitch"
            CMD.Parameters.AddWithValue("@BulletinQueueID", ID)
            CMD.Parameters.AddWithValue("@Sent", sent)
            CMD.ExecuteNonQuery()

        Catch ex As Exception
            WriteLogs("usp_tblBulletinSendtoSwitchQueue_UpdateSendToSwitch" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    Public Function tblStations_UpdateCommetByWMoCode(ByVal Comment As String, ByVal WMOCOde As String)
        'db14
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand


        m_CON.ConnectionString = m_Connection.ConnectionString
        Try

            CMD.Connection = m_CON
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = "usp_tblStations_UpdateCommentbyWmoCode"
            WriteLogs("@Comment:" & Comment)
            WriteLogs("@WMOCode:" & WMOCOde)
            CMD.Parameters.AddWithValue("@Comment", Comment)
            CMD.Parameters.AddWithValue("@WMOCode", WMOCOde)
            CMD.ExecuteNonQuery()
        Catch ex As Exception
            WriteLogs("tblStations_UpdateCommetByWMoCode" & ex.Message)
        End Try

    End Function
    Public Function Report_By_Sensor_with_Pivot_OnlyRealSensor(ByRef tblSensors As System.Data.DataTable, ByVal WMOCode As String, ByVal StartDateTime As String, ByVal EndDateTime As String, ByRef Response_delay As Long) As Long
        Dim t1, t2 As Long
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = m_CON
        CMD.CommandTimeout = 0
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD
        Try
            CMD.CommandText = "usp_Synch_Logs_with_Pivot_OnlyRealSensor"
            CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
            CMD.Parameters.AddWithValue("@StartDateTime", StartDateTime)
            CMD.Parameters.AddWithValue("@EndDateTime", EndDateTime)

            Try
                m_Dataset.Tables("usp_Synch_Logs_with_Pivot_OnlyRealSensor").Clear()
            Catch ex As Exception

            End Try
            t1 = Now.Ticks
            m_DataAdapter.Fill(m_Dataset, "usp_Synch_Logs_with_Pivot_OnlyRealSensor")
            t2 = Now.Ticks
            'WriteLogs("Report_By_Sensor_with_Pivot t2-t1:" & t2 - t1)
            tblSensors = m_Dataset.Tables("usp_Synch_Logs_with_Pivot_OnlyRealSensor")
            Response_delay = (t2 - t1) / Math.Pow(10, 4)
        Catch ex As Exception
            WriteLogs("3-Report_By_Sensor_with_Pivot_OnlyRealSensor:" & ex.Message)
            Report_By_Sensor_with_Pivot_OnlyRealSensor = 3
        End Try
        Report_By_Sensor_with_Pivot_OnlyRealSensor = 0
        Try
            m_CON.Close()
        Catch ex As Exception

        End Try
    End Function
    Public Function GetKnownSensors_Of_Station_WithoutVirtual(ByRef tblSensors As DataTable, ByVal WMOCode As String)
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD
        CMD.CommandTimeout = 0
        Try
            CMD.CommandText = "usp_GetKnownSensorsOfStation_WithoutVirtual"
            CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
            Try
                m_Dataset.Tables("usp_GetKnownSensorsOfStation_WithoutVirtual").Clear()
            Catch ex As Exception
                'WriteLogs("1-GetKnownSensors_Of_Station:" & ex.Message)
                'Return 1
            End Try
            m_DataAdapter.Fill(m_Dataset, "usp_GetKnownSensorsOfStation_WithoutVirtual")
            tblSensors = m_Dataset.Tables("usp_GetKnownSensorsOfStation_WithoutVirtual")
        Catch ex As Exception
            WriteLogs("2-GetKnownSensors_Of_Station_WithoutVirtual:" & ex.Message)
            Return 2
            '  RaiseEvent DataBaseErrors("GetSensors_Of_Station   " & ex.Message)
        End Try
        Return 0
    End Function
    'GetDestinationBYID
    Public Function GetDestinationBYID(ByVal ID As Long, ByRef tblInfo As DataTable) As String
        '   di()
        '   di()
        Try


            Dim con As New SqlClient.SqlConnection
            Dim CMD As New SqlClient.SqlCommand
            Dim Da As New SqlClient.SqlDataAdapter
            Dim Ds As New DataSet
            con.ConnectionString = My.Settings.ICSDBConnectionString
            CMD.Connection = con
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandTimeout = 0
            GetDestinationBYID = ""
            CMD.CommandText = "GetDestinationBYID"
            CMD.Parameters.AddWithValue("@DestinationID", ID)
            Da.SelectCommand = CMD
            CMD.CommandTimeout = 0
            Da.Fill(Ds, "tblInfo")
            tblInfo = Ds.Tables("tblInfo")
        Catch ex As Exception

        End Try
        'For i = 0 To Ds.Tables("tblStations").Rows.Count - 1
        '    StationIKAO.Add(Ds.Tables(0).Rows(i).Item("IKAOCode"))
        '    ListOfWMOCode.Add(Ds.Tables(0).Rows(i).Item("WMOCode"))
        'Next
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    Public Function GetStationsName(ByRef tblInfo As DataTable) As String
        '   di()
        '   di()
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandTimeout = 0
        GetStationsName = ""
        CMD.CommandText = "usp_tblStations_SelectAll_forGenerateFile"
        Da.SelectCommand = CMD
        CMD.CommandTimeout = 0
        Da.Fill(Ds, "tblStations")
        tblInfo = Ds.Tables("tblStations")
        'For i = 0 To Ds.Tables("tblStations").Rows.Count - 1
        '    StationIKAO.Add(Ds.Tables(0).Rows(i).Item("IKAOCode"))
        '    ListOfWMOCode.Add(Ds.Tables(0).Rows(i).Item("WMOCode"))
        'Next
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    Public Function UpdatetblReportsSendToswitchByReportID(ByVal ReportID As Long)
        'db14
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand


        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_tblReports_UpdateSendToSwitch"
            CMD.Parameters.AddWithValue("@ReportID", ReportID)
            CMD.ExecuteNonQuery()

        Catch ex As Exception
            WriteLogs("UpdatetblReportsSendToswitchByReportID" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function

    Public Function GetMetaData(ByRef tblMetaData As DataTable)
        'db15
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Ds As New DataSet
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_tblSendAndReceiveLogs_SelectSendingMetaData"
            Da.SelectCommand = CMD
            Da.Fill(Ds, "tblMetaData")
            tblMetaData = Ds.Tables("tblMetaData")
        Catch ex As Exception
            WriteLogs("GetMetaData" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    Public Function UpdatetblSynopsSendToswitchBySynopID(ByVal SynopID As Long)

        'db18
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand


        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure
            CMD.Parameters.AddWithValue("@SynopID", SynopID)
            CMD.CommandText = "usp_tblSynops_UpdateSendToSwitch"

            CMD.ExecuteNonQuery()

        Catch ex As Exception
            WriteLogs("UpdatetblSynopsSendToswitchBySynopID" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function

    Public Function usp_tblSendAndReceiveLogs_UpdateSentMetaData(ByVal SendAndReceiveLogID As String)

        'db16
        Dim m_CON As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand


        m_CON.ConnectionString = m_Connection.ConnectionString
        Try


            CMD.Connection = m_CON
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure
            CMD.Parameters.AddWithValue("@SendAndReceiveLogID", SendAndReceiveLogID) 'eslah
            CMD.CommandText = "usp_tblSendAndReceiveLogs_UpdateSentMetaData"

            CMD.ExecuteNonQuery()

        Catch ex As Exception
            WriteLogs("usp_tblSendAndReceiveLogs_UpdateSentMetaData" & ex.Message)
        End Try
        '   CMD.Parameters.AddWithValue("@Name", StationName)
        '   GetStationCodeByName = CMD.EndExecuteReader
        '  tol.Se(GetStationCodeByName, "")
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tblSensors"></param>
    ''' <param name="ISLogedError"></param>
    ''' <param name="LogNumber"></param>
    ''' <remarks>
    ''' tblSensors Columns are     1-SensorTypeID  2- SensorTypeName  3-MeasurementUnit 4-FactoryMinValue 5-FactoryMaxValue   
    ''' </remarks>
    Public Sub GetSensorTypeList(ByRef tblSensors As DataTable, ByVal ISLogedError As Boolean, ByVal LogNumber As Long)
        Dim CMD As New SqlClient.SqlCommand
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString

        tblSensors = New DataTable
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD
        Try
            CMD.CommandText = "usp_tblSensorTypes_SelectAll"


            m_DataAdapter.Fill(m_Dataset, "usp_tblSensorTypes_SelectAll")
            tblSensors = m_Dataset.Tables("usp_tblSensorTypes_SelectAll")
        Catch ex As Exception
            'RaiseEvent DataBaseErrors(ex.Message)
            WriteLogs("GetSensorTypeList" & ex.Message)
        Finally
            'm_Dataset.Clear()
            'm_Dataset.Dispose()
        End Try
    End Sub
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="tblSensors"></param>
    ''' <param name="StationName"></param>
    ''' <remarks>
    ''' tblSensors Columns are     1-SensorTypeID  2- VendorName  3-SensorTypeName 4-SensorAliasName 5-SensorValue  6-SensorLocalManValue    
    '''6-SensorLocalMixValue    7-MeasurementUnit
    ''' </remarks>

    Public Function GetSensors_Of_Staion(ByRef tblSensors As DataTable, ByVal WMOcode As String, ByVal clientID As Integer)
        'db3
        Dim con As New SqlClient.SqlConnection
        Dim DataAdapter As New SqlClient.SqlDataAdapter
        Dim Dataset As New DataSet
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = m_Connection.ConnectionString
        Try
            '   WriteLogs("GetSensors_Of_Staion" & WMOcode)
            CMD.Connection = con
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandTimeout = 0
            DataAdapter.SelectCommand = CMD
            Try
                CMD.CommandText = "usp_GetSensorsOfStation"
                CMD.Parameters.AddWithValue("@WMOCode", WMOcode)
                Try
                    Dataset.Tables("usp_GetSensorsOfStation").Clear()
                Catch ex As Exception
                    '     WriteLogs("GetSensors_Of_Staion" & ex.Message)
                End Try
                DataAdapter.Fill(Dataset, "usp_GetSensorsOfStation")
                tblSensors = Dataset.Tables("usp_GetSensorsOfStation")
                'WritetestLogs("tblsenor for " & WMOcode, clientID)
                'For i As Integer = 0 To Dataset.Tables("usp_GetSensorsOfStation").Rows.Count - 1
                '    WritetestLogs(Dataset.Tables("usp_GetSensorsOfStation").Rows(0).Item("JoinSensorType_LoggerModuleID") & "   " & Dataset.Tables("usp_GetSensorsOfStation").Rows(0).Item("SensorValue"), clientID)
                'Next

                CMD.Dispose()
                Dataset.Dispose()
                DataAdapter.Dispose()
                con.Close()
            Catch ex As Exception
                WriteLogs("GetSensors_Of_Staion" & ex.Message)
            End Try
        Catch ex As Exception
        Finally
        End Try
        Return 0
    End Function
    'Public Sub GetLoggerModules_Of_Staion(ByRef tblModules As DataTable, ByVal StationName As String)
    '    Dim m_DataAdapter As New SqlClient.SqlDataAdapter
    '    Dim m_Dataset As New DataSet
    '    Dim m_CON As New SqlClient.SqlConnection
    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    Dim CMD As New SqlClient.SqlCommand
    '    CMD.Connection = m_CON
    '    CMD.CommandType = CommandType.StoredProcedure

    '    m_DataAdapter.SelectCommand = CMD
    '    Try
    '        CMD.CommandText = "usp_GetLoggerModulesByStationName"
    '        CMD.Parameters.AddWithValue("@StationName", StationName)
    '        Try
    '            m_Dataset.Tables("usp_GetLoggerModulesByStationName").Clear()
    '        Catch ex As Exception
    '        End Try
    '        m_DataAdapter.Fill(m_Dataset, "usp_GetLoggerModulesByStationName")
    '        tblModules = m_Dataset.Tables("usp_GetLoggerModulesByStationName")
    '    Catch ex As Exception
    '        WriteLogs("GetLoggerModules_Of_Staion()" & ex.Message)
    '    End Try

    'End Sub

    Public Function DefineLoggerModules_For_Staion(ByRef tblModules As DataTable, ByVal WMOCode As String)
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        Dim i As Long
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure

        Try
            CMD.CommandText = "usp_LoggerModuleDefenition"
            For i = 0 To tblModules.Rows.Count - 1
                CMD.Parameters.Clear()
                CMD.Parameters.AddWithValue("@WMOCode", WMOCode)
                CMD.Parameters.AddWithValue("@LoggerModuleID", tblModules.Rows(i).Item("LoggerModuleID"))
                CMD.Parameters.AddWithValue("@LoggerModuleName", tblModules.Rows(i).Item("LoggerModuleName"))
                CMD.Parameters.AddWithValue("@LoggerModuleModel", tblModules.Rows(i).Item("LoggerModuleModel"))
                CMD.Parameters.AddWithValue("@VendorName", tblModules.Rows(i).Item("VendorName"))
                CMD.Parameters.AddWithValue("@SerialNumber", tblModules.Rows(i).Item("SerialNumber"))
                CMD.Parameters.AddWithValue("@HardwareVersion", tblModules.Rows(i).Item("HardwareVersion"))
                CMD.Parameters.AddWithValue("@SoftwareVersion", tblModules.Rows(i).Item("SoftwareVersion"))
                CMD.Parameters.AddWithValue("@ChannelNumber", tblModules.Rows(i).Item("ChannelNumber"))
                Try
                    CMD.ExecuteNonQuery()
                Catch ex As Exception
                    Return 1
                End Try
            Next i
        Catch ex As Exception
            Return 2
            '  RaiseEvent DataBaseErrors("GetSensors_Of_Station   " & ex.Message)
        End Try
        Return 0
    End Function
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="tblSensors"></param>
    ''' <param name="StationName"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''  ''' tblSensors Columns are     1-SensorTypeName  2- SensorLocalMinValue  3-SensorLocalMaxValue  4-SensorValue  5-SensorAliasName   
    '''6-IsSendingAlarm    7-LastValueUpdatedDateTime
    ''' </remarks>
    Public Function Save_Sensors_Of_Station(ByVal tblSensors As DataTable, ByVal StationName As String) As Long
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        CMD.Connection = m_CON
        CMD.CommandType = CommandType.StoredProcedure

        For i As Integer = 0 To tblSensors.Rows.Count - 1
            With tblSensors.Rows(i)
                Try
                    'Dim SensorTypeID As Integer = GetSensortypeIDByName(.Item("SensorTypeName"))
                    CMD.CommandText = "usp_SensorDefenition"
                    CMD.Parameters.Clear()
                    'CMD.Parameters.AddWithValue("@LoggerModuleName", tblSensors.Rows(i).Item("LoggerModuleName"))
                    'CMD.Parameters.AddWithValue("@SensorLocalMinValue", tblSensors.Rows(i).Item("SensorLocalMinValue"))
                    'CMD.Parameters.AddWithValue("@SensorLocalMaxValue", tblSensors.Rows(i).Item("SensorLocalMaxValue"))
                    CMD.Parameters.AddWithValue("@SensorValue", tblSensors.Rows(i).Item("SensorValue"))
                    CMD.Parameters.AddWithValue("@IsSendingAlarm", tblSensors.Rows(i).Item("IsSendingAlarm"))
                    CMD.Parameters.AddWithValue("@LastValueUpdatedDateTime", tblSensors.Rows(i).Item("LastValueUpdatedDateTime"))
                    CMD.Parameters.AddWithValue("@SensorTypeName", tblSensors.Rows(i).Item("SensorTypeName"))
                    CMD.Parameters.AddWithValue("@FK_LoggerModuleID", tblSensors.Rows(i).Item("FK_LoggerModuleID"))
                    CMD.Parameters.AddWithValue("@ChannelName", tblSensors.Rows(i).Item("ChannelName"))
                    CMD.Parameters.AddWithValue("@SensorAliasName", tblSensors.Rows(i).Item("SensorAliasName"))
                    CMD.Parameters.AddWithValue("@MonthNum", System.DateTime.Now.Month)
                    CMD.ExecuteNonQuery()
                Catch ex As Exception
                    WriteLogs("Save_Sensors_Of_Station" & ex.Message)
                    Return 1
                End Try
            End With
        Next
        Return 0
        '  aDataTable = tblSensors
    End Function
    'Public Function Update_Sensors_Definition(ByVal tblSensors As DataTable) As Long
    '    '@JoinSensorType_LoggerModuleID bigint,
    '    '@SensorAliasName nvarchar(50),
    '    '@SensorLocalMinValue float,
    '    '@SensorLocalMaxValue float,
    '    '@SensorTypeName nvarchar(50)
    '    Dim CMD As New SqlClient.SqlCommand
    '    Dim m_CON As New SqlClient.SqlConnection
    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    CMD.Connection = m_CON
    '    CMD.CommandType = CommandType.StoredProcedure
    '    For i As Integer = 0 To tblSensors.Rows.Count - 1
    '        With tblSensors.Rows(i)
    '            Try
    '                CMD.CommandText = "usp_UpdateSensorsByJoinSensorType_LoggerModuleID"
    '                CMD.Parameters.Clear()
    '                CMD.Parameters.AddWithValue("@JoinSensorType_LoggerModuleID", tblSensors.Rows(i).Item("JoinSensorType_LoggerModuleID"))
    '                CMD.Parameters.AddWithValue("@SensorAliasName", tblSensors.Rows(i).Item("SensorAliasName"))
    '                'CMD.Parameters.AddWithValue("@SensorLocalMinValue", tblSensors.Rows(i).Item("SensorLocalMinValue"))
    '                'CMD.Parameters.AddWithValue("@SensorLocalMaxValue", tblSensors.Rows(i).Item("SensorLocalMaxValue"))
    '                CMD.Parameters.AddWithValue("@SensorTypeName", tblSensors.Rows(i).Item("SensorTypeName"))

    '                CMD.ExecuteNonQuery()
    '            Catch ex As Exception
    '                WriteLogs("UPdate_Sensor_Definition" & ex.Message)
    '                Return 1
    '            End Try
    '        End With
    '    Next
    '    Return 0
    '    '  aDataTable = tblSensors
    'End Function

    'Private Function GetSensortypeIDByName(ByVal SensorName As String) As Integer
    '    Dim CMD As New SqlClient.SqlCommand
    '    Dim m_CON As New SqlClient.SqlConnection
    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    CMD.Connection = m_CON
    '    CMD.CommandType = CommandType.StoredProcedure
    '    GetSensortypeIDByName = 0
    '    CMD.CommandText = "GetStationCodeByName"
    '    CMD.Parameters.AddWithValue("@Name", SensorName)
    '    GetSensortypeIDByName = Convert.ToUInt32(CMD.ExecuteScalar())

    'End Function
    ''' <summary>
    '''     ''' </summary>
    ''' <param name="tblSensors"></param>
    ''' <param name="LastValueUpdatedDateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update_Value_Of_Sensors(ByVal tblSensors As DataTable) As Long
        'db6
        Dim con As New SqlClient.SqlConnection
        con.ConnectionString = m_Connection.ConnectionString
        Try
            Dim DataAdapter As New SqlClient.SqlDataAdapter
            Dim Dataset As New DataSet
            Dim CMD As New SqlClient.SqlCommand
            CMD.Connection = con
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            CMD.CommandType = CommandType.StoredProcedure
            DataAdapter.SelectCommand = CMD
            For i As Integer = 0 To tblSensors.Rows.Count - 1
                With tblSensors.Rows(i)
                    Try
                        CMD.Parameters.Clear()
                        CMD.CommandText = "usp_UpdateValueOfSensorsByJoinSensorType_LoggerModuleID"  'done
                        '   WriteLogs("JoinSensorType_LoggerModuleID:" & .Item("JoinSensorType_LoggerModuleID"))
                        '   WriteLogs("SensorValue:" & .Item("SensorValue"))
                        '  WriteLogs("LastValueUpdatedDateTime:" & .Item("LastValueUpdatedDateTime"))
                        '  WriteLogs("MonthNum:" & CDate(.Item("LastValueUpdatedDateTime")).Month)
                        CMD.Parameters.AddWithValue("@JoinSensorType_LoggerModuleID", .Item("JoinSensorType_LoggerModuleID"))
                        CMD.Parameters.AddWithValue("@SensorValue", .Item("SensorValue"))
                        CMD.Parameters.AddWithValue("@LastValueUpdatedDateTime", .Item("LastValueUpdatedDateTime"))
                        CMD.Parameters.AddWithValue("@MonthNum", CDate(.Item("LastValueUpdatedDateTime")).Month)
                        ' CMD.Parameters.AddWithValue("@SensorLocalMinValue", .Item("SensorLocalMinValue"))
                        CMD.ExecuteNonQuery()
                    Catch ex As Exception

                        WriteLogs("Update_Value_Of_Sensors" & ex.Message)

                        Return 1
                    End Try
                End With
            Next

            Return 0
        Catch ex As Exception

        End Try
    End Function
    ' ''' <summary>
    ' ''' 
    ' ''' </summary>
    ' ''' <param name="StationName"></param>
    ' ''' <returns></returns>
    ' ''' <remarks>    
    ' '''  Clear all Modules of a determinated station to reconfiguration
    ' ''' </remarks>
    'Public Function ClearAllModulesByStation(ByVal StationName As String) As Long
    '    'usp_ClearLoggerModulesByStationName
    '    Dim m_DataAdapter As New SqlClient.SqlDataAdapter
    '    Dim m_Dataset As New DataSet
    '    Dim m_CON As New SqlClient.SqlConnection
    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    Dim CMD As New SqlClient.SqlCommand
    '    CMD.Connection = m_CON
    '    CMD.CommandType = CommandType.StoredProcedure
    '    m_DataAdapter.SelectCommand = CMD


    '    Try
    '        CMD.Parameters.AddWithValue("@StationName", StationName)
    '        CMD.CommandText = "usp_ClearLoggerModulesByStationName"
    '        CMD.ExecuteNonQuery()
    '    Catch ex As Exception
    '        WriteLogs("ClearAllModulesByStation" & ex.Message)
    '        Return 1
    '    End Try
    '    Return 0
    'End Function

#Region "Tabatabaee"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="WMOCODE"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 

    Public Function GetLastReportDateTime() As String
        'db9eve
        ' Dim Con As New SqlClient.SqlConnection
        Dim dr As SqlClient.SqlDataReader
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try
            If m_CON.State = ConnectionState.Closed Then
                m_CON.Open()
            End If

            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure
            ' GetLastSampleIDBYWMOCode = ""
            CMD.CommandText = "usp_tblReports_SelectLast"
            CMD.CommandTimeout = 0
            'CMD.Parameters.AddWithValue("@WMOCode", WMOCODE)
            dr = CMD.ExecuteReader
            If dr.HasRows Then
                dr.Read()
                If dr.Item(0) = Nothing Then
                    GetLastReportDateTime = "2000-01-01 00:00:00"
                Else
                    GetLastReportDateTime = dr.Item("ReportDateTime").ToString
                    '  WriteLogs("GetLastSampleIDBYWMOCode" & dr.Item("SamplingID").ToString)
                End If
                ' IIf(, GetLastSampleIDBYWMOCode = 0, GetLastSampleIDBYWMOCode = dr.Item(0))
            Else
                GetLastReportDateTime = "2000-01-01 00:00:00"
                '   WriteLogs("GetLastSampleIDBYWMOCode" & "nothing")
            End If
            '     MsgBox()
        Catch ex As Exception
            GetLastReportDateTime = "2000-01-01 00:00:00"
            WriteLogs("GetLastReportDateTime" & ex.Message)
        Finally
        
        End Try
        Try
            dr.Close()
            m_CON.Close()
        Catch ex As Exception
        End Try
        '   System.Threading.Thread.Sleep(200)
        '  tol.Se(GetStationCodeByName, "")
    End Function

    Public Function GetLastSampleIDBYWMOCode(ByVal WMOCODE As String) As String
        'db9
        ' Dim Con As New SqlClient.SqlConnection
        Dim dr As SqlClient.SqlDataReader
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try
            If m_CON.State = ConnectionState.Closed Then
                m_CON.Open()
            End If


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure
            ' GetLastSampleIDBYWMOCode = ""
            CMD.CommandText = "usp_GetLastSampleIDbyWMOCode"
            CMD.CommandTimeout = 0
            CMD.Parameters.AddWithValue("@WMOCode", WMOCODE)
            dr = CMD.ExecuteReader
            If dr.HasRows Then
                dr.Read()
                If dr.Item(0) = Nothing Then
                    GetLastSampleIDBYWMOCode = 0
                Else
                    GetLastSampleIDBYWMOCode = dr.Item("SamplingID").ToString
                    '  WriteLogs("GetLastSampleIDBYWMOCode" & dr.Item("SamplingID").ToString)
                End If
                ' IIf(, GetLastSampleIDBYWMOCode = 0, GetLastSampleIDBYWMOCode = dr.Item(0))
            Else
                GetLastSampleIDBYWMOCode = 0
                '   WriteLogs("GetLastSampleIDBYWMOCode" & "nothing")
            End If
            '     MsgBox()
        Catch ex As Exception
            GetLastSampleIDBYWMOCode = 0
            WriteLogs("GetLastSampleIDBYWMOCode" & ex.Message)
        Finally
            dr.Close()
            m_CON.Close()
        End Try
        '   System.Threading.Thread.Sleep(200)
        '  tol.Se(GetStationCodeByName, "")
    End Function
  
    Public Function GetLastDailyTimeBYWMOCode(ByVal WMOCODE As String) As String
        'db10
        ' Dim Con As New SqlClient.SqlConnection
        Dim dr As SqlClient.SqlDataReader
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try
            If m_CON.State = ConnectionState.Closed Then
                m_CON.Open()
            End If

            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure
            ' GetLastSampleIDBYWMOCode = ""
            CMD.CommandText = "usp_GetLastDailyDateTimebyWMOCode"
            '    WriteLogs("GetLastDailyTimeBYWMOCode" & WMOCODE)
            CMD.CommandTimeout = 0
            CMD.Parameters.AddWithValue("@WMOCode", WMOCODE)
            dr = CMD.ExecuteReader
            If dr.HasRows Then
                dr.Read()
                '          WriteLogs(dr.Item(0))
                If dr.Item(0) = Nothing Then
                    GetLastDailyTimeBYWMOCode = ""
                Else
                    GetLastDailyTimeBYWMOCode = Format(dr.Item("DateTimeMilady"), "yyyy/MM/dd HH:mm:ss")
                    '      WriteLogs("GetLastDailyTimeBYWMOCode" & GetLastDailyTimeBYWMOCode)
                    '  WriteLogs("GetLastSampleIDBYWMOCode" & dr.Item("SamplingID").ToString)
                End If
                ' IIf(, GetLastSampleIDBYWMOCode = 0, GetLastSampleIDBYWMOCode = dr.Item(0))
            Else
                GetLastDailyTimeBYWMOCode = ""
                '   WriteLogs("GetLastSampleIDBYWMOCode" & "nothing")
            End If
            '     MsgBox()
        Catch ex As Exception
            GetLastDailyTimeBYWMOCode = ""
            WriteLogs("GetLastDailyTimeBYWMOCode" & ex.Message)
        Finally
            dr.Close()
            m_CON.Close()
        End Try
        '   System.Threading.Thread.Sleep(200)
        '  tol.Se(GetStationCodeByName, "")
    End Function
    'usp_tblReports_InsertBufferCenter

    Public Function InsetIntotblBufferReport(ByVal tblReport As DataTable, ByVal ID As List(Of String), ByVal data As Byte(), ByRef Errorlog As String) As Long

        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim m_CON As New SqlClient.SqlConnection
        Dim flagupdate As Boolean = False
        m_CON.ConnectionString = m_Connection.ConnectionString
        Dim Count As Integer = 0
        Dim CMD As New SqlClient.SqlCommand
        CMD.Connection = m_CON
        If m_CON.State = ConnectionState.Closed Then
            m_CON.Open()
        End If
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD
        CMD.CommandTimeout = 0
        For i As Integer = 0 To tblReport.Rows.Count - 1
            With tblReport.Rows(i)
                Try


                    Try
                        CMD.Parameters.Clear()
                    Catch ex As Exception

                    End Try

                    CMD.CommandText = "usp_tblReports_InsertBufferCenter"
                    CMD.Parameters.AddWithValue("@ReportID", .Item("ReportID"))
                    CMD.Parameters.AddWithValue("@ReportContent", data)
                    CMD.Parameters.AddWithValue("@ReportDateTime", .Item("ReportDateTime"))
                    CMD.Parameters.AddWithValue("@SendingToSwitch", .Item("SendingToSwitch"))
                    CMD.Parameters.AddWithValue("@SentToSwitch", .Item("SentToSwitch"))
                    CMD.Parameters.AddWithValue("@WMOCode", .Item("WMOCode"))
                    CMD.Parameters.AddWithValue("@ErrorConTent", .Item("ErrorConTent"))
                    CMD.Parameters.AddWithValue("@FK_UserID", .Item("FK_UserID"))
                    CMD.Parameters.AddWithValue("@Auto", .Item("Auto"))
                    CMD.Parameters.AddWithValue("@ReportTypeName", .Item("ReportTypeName"))
                    '   CMD.Parameters.AddWithValue("@ReportDateTime", .Item("ReportDateTime"))


                    flagupdate = False
                    Errorlog = Errorlog + "Insert report DB Start  with  id " & .Item("ReportID") & vbCrLf
                    CMD.ExecuteNonQuery()
                    Errorlog = Errorlog + "Insert report DB finish  with  id " & .Item("ReportID") & vbCrLf
                    ID.Add(.Item("ReportID"))
                Catch ex As Exception
                    WriteLogs("InsetIntoReport 1 :" & ex.Message)
                    Errorlog = Errorlog + "Insert report DB error is " & ex.Message & vbCrLf
                    If ex.Message.Contains("Violation of PRIMARY KEY constraint") Then
                        flagupdate = True
                    End If
                End Try
                Try


                    If flagupdate Then
                        CMD.Parameters.Clear()

                        CMD.CommandText = "usp_tblReports_update_BufferCenter"
                        CMD.Parameters.AddWithValue("@ReportID", .Item("ReportID"))
                        CMD.Parameters.AddWithValue("@ReportContent", data)
                        CMD.Parameters.AddWithValue("@ReportDateTime", .Item("ReportDateTime"))
                        CMD.Parameters.AddWithValue("@SendingToSwitch", .Item("SendingToSwitch"))
                        CMD.Parameters.AddWithValue("@SentToSwitch", .Item("SentToSwitch"))
                        CMD.Parameters.AddWithValue("@WMOCode", .Item("WMOCode"))
                        CMD.Parameters.AddWithValue("@ErrorConTent", .Item("ErrorConTent"))
                        CMD.Parameters.AddWithValue("@FK_UserID", .Item("FK_UserID"))
                        CMD.Parameters.AddWithValue("@Auto", .Item("Auto"))
                        CMD.Parameters.AddWithValue("@ReportTypeName", .Item("ReportTypeName"))
                        Errorlog = Errorlog + "update report DB Start  with  id " & .Item("ReportID") & vbCrLf
                        CMD.ExecuteNonQuery()
                        Errorlog = Errorlog + "update report DB finish  with  id " & .Item("ReportID") & vbCrLf
                        ID.Add(.Item("ReportID"))
                    End If
                Catch ex As Exception
                    '   WriteLogs("UpdateIntoReport :" & ex.Message)
                    Errorlog = Errorlog + "update report DB error is " & ex.Message & vbCrLf



                End Try
            End With
        Next
        m_CON.Close()
        Return 0
    End Function
    Public Function InsetIntotblReport(ByVal tblReport As DataTable, ByVal ID As List(Of String), ByRef Errorlog As String, ByVal DuplicateID As List(Of String)) As Long



        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim m_CON As New SqlClient.SqlConnection
        Dim flagupdate As Boolean = False
        m_CON.ConnectionString = m_Connection.ConnectionString
        Dim Count As Integer = 0
        Dim CMD As New SqlClient.SqlCommand
        CMD.Connection = m_CON
        If m_CON.State = ConnectionState.Closed Then
            m_CON.Open()
        End If
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD
        CMD.CommandTimeout = 0
        '  For i As Integer = 0 To tblReport.Rows.Count - 1
        With tblReport.Rows(0)
            Try


                Try
                    CMD.Parameters.Clear()
                Catch ex As Exception

                End Try



                CMD.CommandText = "usp_tblReports_InsertCenter"

                CMD.Parameters.AddWithValue("@ReportID", .Item("ReportID"))
                CMD.Parameters.AddWithValue("@ReportContent", .Item("ReportContent"))
                CMD.Parameters.AddWithValue("@ReportDateTime", .Item("ReportDateTime"))
                CMD.Parameters.AddWithValue("@SendingToSwitch", .Item("SendingToSwitch"))
                CMD.Parameters.AddWithValue("@SentToSwitch", .Item("SentToSwitch"))
                CMD.Parameters.AddWithValue("@WMOCode", .Item("WMOCode"))
                CMD.Parameters.AddWithValue("@ErrorConTent", .Item("ErrorConTent"))
                CMD.Parameters.AddWithValue("@FK_UserID", .Item("FK_UserID"))
                CMD.Parameters.AddWithValue("@Auto", .Item("Auto"))
                CMD.Parameters.AddWithValue("@ReportTypeName", .Item("ReportTypeName"))

                CMD.Parameters.AddWithValue("@MSGNO", .Item("MSGNO"))


                flagupdate = False
                '  WriteLogs("stepM1" & .Item("ReportID"))
                Errorlog = Errorlog + "Insert report start" & vbCrLf
                CMD.ExecuteNonQuery()
                Errorlog = Errorlog + "Insert Report successfully  report id is " & .Item("ReportID") & vbCrLf
                '  WriteLogs("stepM2")

                ID.Add(.Item("ReportID"))
                '   WriteLogs("stepM3" & ID(Count))
            Catch ex As Exception
                WriteLogs("Insert Report error for report id   " & .Item("ReportID") & " is " & ex.Message & vbCrLf)
                Errorlog = Errorlog + "Insert Report error for report id   " & .Item("ReportID") & " is " & ex.Message & vbCrLf
                If ex.Message.Contains("Violation of PRIMARY KEY constraint") Then

                    flagupdate = True
                    '  Else

                End If
            End Try
            Try


                If flagupdate Then
                    'CMD.Parameters.Clear()
                    Dim erlog As String
                    Try


                        If ISReportAuto(.Item("ReportID"), erlog) = False Then
                            Errorlog = Errorlog + "no need to update  report id   " & .Item("ReportID") & vbCrLf
                            DuplicateID.Add(.Item("ReportID"))
                            m_CON.Close()
                            Exit Function

                        End If

                        If erlog <> "" Then
                            Errorlog = Errorlog + "auto status select error " & erlog & vbCrLf
                        End If


                    Catch ex As Exception
                        Errorlog = Errorlog & ex.Message & vbCrLf
                    End Try

                    CMD.Parameters.Clear()

                    CMD.CommandText = "usp_tblReports_update_Center"
                    CMD.Parameters.AddWithValue("@ReportID", .Item("ReportID"))
                    CMD.Parameters.AddWithValue("@ReportContent", .Item("ReportContent"))
                    CMD.Parameters.AddWithValue("@ReportDateTime", .Item("ReportDateTime"))
                    CMD.Parameters.AddWithValue("@SendingToSwitch", .Item("SendingToSwitch"))
                    CMD.Parameters.AddWithValue("@SentToSwitch", .Item("SentToSwitch"))
                    CMD.Parameters.AddWithValue("@WMOCode", .Item("WMOCode"))
                    CMD.Parameters.AddWithValue("@ErrorConTent", .Item("ErrorConTent"))
                    CMD.Parameters.AddWithValue("@FK_UserID", .Item("FK_UserID"))
                    CMD.Parameters.AddWithValue("@Auto", .Item("Auto"))
                    CMD.Parameters.AddWithValue("@ReportTypeName", .Item("ReportTypeName"))
                    CMD.Parameters.AddWithValue("@MSGNO", .Item("MSGNO"))
                    Errorlog = Errorlog + "update report start" & vbCrLf
                    CMD.ExecuteNonQuery()
                    Errorlog = Errorlog + "update Report successfully  report id is " & .Item("ReportID") & vbCrLf
                    ID.Add(.Item("ReportID"))
                    ' WriteLogs(ID.Count)

                End If
            Catch ex As Exception
                WriteLogs("UpdateIntoReport :" & ex.Message)
                Errorlog = Errorlog + "uRpdate report error for report id   " & .Item("ReportID") & " is " & ex.Message & vbCrLf

            End Try
        End With
        '    Next
        m_CON.Close()

    End Function



    Public Function InsetIntotblRatLogs(ByVal tblRatLogs As DataTable, ByRef Errorlog As String) As String
        'db2
        '   WriteLogs("InsetIntotblRatLogs")
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        '    WriteLogs(m_Connection.ConnectionString)
        CMD.Connection = m_CON
        If m_CON.State = ConnectionState.Closed Then
            m_CON.Open()
        End If
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD

        With tblRatLogs.Rows(0)
            Try
                CMD.Parameters.Clear()
                '  WriteLogs("items " & .Item("WMOCODE"))
                'WriteLogs("items " & .Item("FileNameAndAddress"))
                'WriteLogs("items " & .Item("LogDateTime"))
                'WriteLogs("items " & .Item("SendOrReceive"))
                'WriteLogs("items " & .Item("FileSize"))
                'WriteLogs("items " & .Item("HasStoredInDB"))
                'WriteLogs("items " & .Item("StationMetadataSendingStatus"))
                'WriteLogs("items " & .Item("HasSentToSwitch"))
                'WriteLogs("items " & .Item("SendAndReceiveLogID"))
                'WriteLogs("items " & .Item("Fk_FileType"))

                CMD.CommandText = "usp_tblSendAndReceiveLogs_InsertCenter"
                CMD.Parameters.AddWithValue("@WMOCODE", .Item("WMOCODE"))
                CMD.Parameters.AddWithValue("@FileNameAndAddress", .Item("FileNameAndAddress"))
                CMD.Parameters.AddWithValue("@LogDateTime", Convert.ToDateTime(.Item("LogDateTime")))
                CMD.Parameters.AddWithValue("@SendOrReceive", .Item("SendOrReceive"))
                CMD.Parameters.AddWithValue("@FileSize", .Item("FileSize"))
                CMD.Parameters.AddWithValue("@HasStoredInDB", .Item("HasStoredInDB"))
                CMD.Parameters.AddWithValue("@StationMetadataSendingStatus", .Item("StationMetadataSendingStatus"))
                CMD.Parameters.AddWithValue("@HasSentToSwitch", .Item("HasSentToSwitch"))
                CMD.Parameters.AddWithValue("@SendAndReceiveLogID", .Item("SendAndReceiveLogID")) 'eslah
                CMD.Parameters.AddWithValue("@Fk_FileType", .Item("Fk_FileType"))
                InsetIntotblRatLogs = CStr(CMD.ExecuteScalar())
                Errorlog = "Insert done successfully"
            Catch ex As Exception
                WriteLogs("InsetIntotblRatLogs:" & ex.Message)
                Errorlog = "InsetIntotblRatLogs:" & ex.Message
            End Try
        End With
        m_CON.Close()
        If Errorlog.Contains("duplicate key") Then
            InsetIntotblRatLogs = tblRatLogs.Rows(0).Item("SendAndReceiveLogID")
        End If
        '    Return 0
    End Function
    Public Function GetLastUpdateByWMOCode(ByVal WMOCode As String) As DateTime
        'db11
        '   WriteLogs("InsetIntotblRatLogs")
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter

        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString

        CMD.Connection = m_CON
        If m_CON.State = ConnectionState.Closed Then
            m_CON.Open()
        End If
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD


        Try
            CMD.Parameters.Clear()
            CMD.CommandText = "usp_tblJoinSensorType_LoggerModule_GetLastUpdateByWMOCode"
            '  WriteLogs(".Item(FileNameAndAddress)" & .Item("FileNameAndAddress"))

            CMD.Parameters.AddWithValue("@WMOCODE", WMOCode)

            GetLastUpdateByWMOCode = CDate(CMD.ExecuteScalar())
            '  WriteLogs(InsetIntotblRatLogs)
        Catch ex As Exception
            WriteLogs("InsetIntotblRatLogs:" & ex.Message)
        End Try

        m_CON.Close()
        '    Return 0
    End Function
    Public Function usp_tblSendAndReceiveLogs_StoredInDB(ByVal HasStoredInDB As Boolean, ByVal SendAndReceiveLogID As String) As Boolean

        'db5
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim m_CON As New SqlClient.SqlConnection
        Try
            m_CON.ConnectionString = My.Settings.ICSDBConnectionString
            Dim CMD As New SqlClient.SqlCommand
            CMD.Connection = m_CON
            If m_CON.State = ConnectionState.Closed Then
                m_CON.Open()
            End If
            CMD.CommandType = CommandType.StoredProcedure
            m_DataAdapter.SelectCommand = CMD
            CMD.CommandText = "usp_tblSendAndReceiveLogs_StoredInDBForDCS"
            '   CMD.Parameters.AddWithValue("@SamplingID", Convert.ToDecimal(tblSampling.Rows(i).Item("SamplingID")))
            CMD.Parameters.AddWithValue("@HasStoredInDB", HasStoredInDB)
            CMD.Parameters.AddWithValue("@SendAndReceiveLogID", SendAndReceiveLogID) 'eslah
            CMD.ExecuteNonQuery()
        Catch ex As Exception
            WriteLogs("usp_tblSendAndReceiveLogs_StoredInDB  " & ex.Message)
        Finally
            m_CON.Close()
        End Try


    End Function
    ''' <summary>
    ''' tblSamling Field  1-SamplingID 2-SensorValue 3-SampleDateTime 4-QualityControlLevel  5-FK_JoinSensorType_LoggerModuleID 6-FK_SendAndReceiveLogID
    ''' </summary>
    ''' <param name="tblSampling"></param>
    ''' <returns></returns>
    ''' <remarks>tblSamling Field  1-SamplingID 2-SensorValue 3-SampleDateTime 4-QualityControlLevel  5-FK_JoinSensorType_LoggerModuleID 6-FK_SendAndReceiveLogID   </remarks>
    Public Function InsetIntotblSampling1(ByVal tblSampling As DataTable) As Boolean
        'db4
        InsetIntotblSampling1 = True
        Dim tblSensors As New DataTable
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = My.Settings.ICSDBConnectionString
        Dim CMD As New SqlClient.SqlCommand
        CMD.Connection = m_CON
        If m_CON.State = ConnectionState.Closed Then
            m_CON.Open()
        End If
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD
        Dim err As String = ""
        Dim WMOCode As String
        If tblSampling.Rows.Count > 0 Then
            WMOCode = tblSampling.Rows(0).Item("WMOCode")
            For i As Integer = 0 To tblSampling.Rows.Count - 1
                With tblSampling.Rows(i)
                    Try
                        If Not IsDBNull(tblSampling.Rows(i).Item("SensorValue")) Then
                            If tblSampling.Rows(i).Item("SensorValue") <> "" Then
                                CMD.Parameters.Clear()
                                CMD.CommandText = "usp_tblSamplings_Insert_For_Center"
                                '   CMD.Parameters.AddWithValue("@SamplingID", Convert.ToDecimal(tblSampling.Rows(i).Item("SamplingID")))
                                CMD.Parameters.AddWithValue("@SampleDateTime", Convert.ToDateTime(tblSampling.Rows(i).Item("SampleDateTime")))
                                CMD.Parameters.AddWithValue("@Value", Convert.ToDouble(.Item("SensorValue")))
                                '  CMD.Parameters.AddWithValue("@QualityControlLevel", CInt(tblSampling.Rows(i).Item("QualityControlLevel")))
                                CMD.Parameters.AddWithValue("@FK_SendAndReceiveLogID", CStr(tblSampling.Rows(i).Item("FK_SendAndReceiveLogID"))) 'eslah
                                CMD.Parameters.AddWithValue("@FK_JoinSensorType_LoggerModuleID", CStr(tblSampling.Rows(i).Item("FK_JoinSensorType_LoggerModuleID")))
                                CMD.Parameters.AddWithValue("@WMOCode", .Item("WMOCode"))
                                CMD.Parameters.AddWithValue("@MonthNum", System.DateTime.Now.Month)

                                CMD.ExecuteNonQuery()
                                err = ""
                                InsetIntotblSampling1 = True
                            End If
                        End If
                    Catch ex As Exception
                        InsetIntotblSampling1 = False
                        err = ex.Message
                        '   WriteLogs("InsetIntotblSampling1:" & ex.Message)
                        '  WriteLogs(" CLng(tblSampling.Rows(i).Item(""FK_JoinSensorType_LoggerModuleID""):" & CLng(tblSampling.Rows(i).Item("FK_JoinSensorType_LoggerModuleID")))
                    End Try
                    Try
                        If err.Contains("duplicate key") Then
                            err = ""
                            '   MsgBox("")
                            ' WriteLogs("Update")
                            CMD.CommandText = "usp_tblSamplings_Update_Center"
                            CMD.ExecuteNonQuery()
                            ' WriteLogs("update")
                            InsetIntotblSampling1 = True
                        ElseIf err <> "" Then
                            WriteLogs("insertblSampling:" & err)
                            InsetIntotblSampling1 = False
                        End If
                    Catch ex As Exception
                        WriteLogs("updatetblSampling:" & ex.Message)
                        Return False
                    End Try
                End With

            Next



        End If
        m_CON.Close()
        Return True
    End Function
    ''' <summary>
    ''' tblSamling Field  1-SamplingID 2-SensorValue 3-SampleDateTime 4-QualityControlLevel  5-FK_JoinSensorType_LoggerModuleID 6-FK_SendAndReceiveLogID
    ''' </summary>
    ''' <param name="tblSampling"></param>
    ''' <returns></returns>
    ''' <remarks>tblSamling Field  1-SamplingID 2-SensorValue 3-SampleDateTime 4-QualityControlLevel  5-FK_JoinSensorType_LoggerModuleID 6-FK_SendAndReceiveLogID   </remarks>
    'Public Function InsetIntotblSampling(ByVal tblSampling As DataTable) As Long
    '    'db19
    '    Dim m_DataAdapter As New SqlClient.SqlDataAdapter
    '    Dim m_Dataset As New DataSet
    '    Dim m_CON As New SqlClient.SqlConnection
    '    m_CON.ConnectionString = m_Connection.ConnectionString
    '    Dim CMD As New SqlClient.SqlCommand
    '    CMD.Connection = m_CON
    '    If m_CON.State = ConnectionState.Closed Then
    '        m_CON.Open()
    '    End If
    '    CMD.CommandType = CommandType.StoredProcedure
    '    m_DataAdapter.SelectCommand = CMD
    '    For i As Integer = 0 To tblSampling.Rows.Count - 1
    '        With tblSampling.Rows(i)
    '            Try
    '                CMD.Parameters.Clear()
    '                CMD.CommandText = "usp_tblSamplings_Insert"
    '                CMD.Parameters.AddWithValue("@SamplingID", Convert.ToDecimal(tblSampling.Rows(i).Item("SamplingID")))
    '                CMD.Parameters.AddWithValue("@SampleDateTime", Convert.ToDateTime(tblSampling.Rows(i).Item("SampleDateTime")))
    '                CMD.Parameters.AddWithValue("@Value", Convert.ToDouble(.Item("SensorValue")))
    '                CMD.Parameters.AddWithValue("@QualityControlLevel", CInt(tblSampling.Rows(i).Item("QualityControlLevel")))
    '                CMD.Parameters.AddWithValue("@FK_SendAndReceiveLogID", CStr(tblSampling.Rows(i).Item("FK_SendAndReceiveLogID"))) 'eslah
    '                CMD.Parameters.AddWithValue("@FK_JoinSensorType_LoggerModuleID", CLng(tblSampling.Rows(i).Item("FK_JoinSensorType_LoggerModuleID")))
    '                CMD.ExecuteNonQuery()
    '            Catch ex As Exception
    '                WriteLogs("InsetIntotblSampling:" & ex.Message)
    '            End Try
    '        End With
    '    Next
    '    m_CON.Close()
    '    Return 0
    'End Function
#End Region
#End Region
#Region "Property"
    Public Property ConnectionString()
        Get
            Return m_Connection.ConnectionString
        End Get
        Set(ByVal value)

        End Set
    End Property
    Public ReadOnly Property ServerName()
        Get
            Return m_Connection.ServerName
        End Get

    End Property

#End Region

    Public Sub UpDatetblSendandRecieveHasRain(ByVal RATLogID As String, ByVal HasRain As Boolean)
        WriteLogs("UpDatetblSendandRecieve" & " RATLogID " & RATLogID)
        Dim MyCon As New SqlClient.SqlConnection
        Dim comMain As New SqlClient.SqlCommand
        Try
            MyCon.ConnectionString = My.Settings.ICSDBConnectionString
            If MyCon.State = ConnectionState.Closed Then
                MyCon.Open()
            End If
            comMain.Connection = MyCon
            comMain.CommandType = CommandType.StoredProcedure
            comMain.CommandText = "usp_tblSendAndReceiveLogs_Update_RainStatus"
            comMain.Parameters.AddWithValue("@SendAndReceiveLogID", RATLogID)
            comMain.Parameters.AddWithValue("@HasRain", HasRain)

            comMain.ExecuteNonQuery()
        Catch ex As Exception
            WriteLogs("UpDatetblSendandRecieve" & ex.Message)
        Finally
            If MyCon.State = ConnectionState.Open Then
                MyCon.Close()
            End If
        End Try
    End Sub
    Public Function UpDatetblSendandRecieve(ByVal RATLogID As String) As Long
        WriteLogs("UpDatetblSendandRecieve" & " RATLogID " & RATLogID)
        Dim MyCon As New SqlClient.SqlConnection
        Dim comMain As New SqlClient.SqlCommand
        Try
            MyCon.ConnectionString = My.Settings.ICSDBConnectionString
            If MyCon.State = ConnectionState.Closed Then
                MyCon.Open()
            End If
            comMain.Connection = MyCon
            comMain.CommandType = CommandType.StoredProcedure
            comMain.CommandText = "usp_tblSendAndReceiveLogs_StoredInDB"
            comMain.Parameters.AddWithValue("@SendAndReceiveLogID", RATLogID)
            comMain.Parameters.AddWithValue("@HasStoredInDB", 1)
            comMain.ExecuteNonQuery()
            UpDatetblSendandRecieve = 0
        Catch ex As Exception
            UpDatetblSendandRecieve = ex.Source
            WriteLogs("UpDatetblSendandRecieve" & ex.Message)
        Finally
            If MyCon.State = ConnectionState.Open Then
                MyCon.Close()
            End If
        End Try
    End Function
    Public Function updatetblSendAndRecievedSendToSwich(ByVal SendAndReceiveLogID As String, ByVal StationMetadataSendingStatus As Integer, ByVal HasSentToSwitch As Integer) As Boolean
        Dim MyCon As New SqlClient.SqlConnection
        Dim ComUpdateTblStation As New SqlClient.SqlCommand
        MyCon.ConnectionString = My.Settings.ICSDBConnectionString
        If MyCon.State = ConnectionState.Open Then
            MyCon.Close()
        End If
        Try
            'usp_tblSendAndReceiveLogs_UpdateSenttoSwitch
            ComUpdateTblStation.Connection = MyCon
            ComUpdateTblStation.CommandType = CommandType.StoredProcedure
            MyCon.Open()
            ComUpdateTblStation.CommandText = "usp_tblSendAndReceiveLogs_UpdateSentMetaData"
            ComUpdateTblStation.Parameters.AddWithValue("@SendAndReceiveLogID", SendAndReceiveLogID)
            ComUpdateTblStation.Parameters.AddWithValue("@StationMetadataSendingStatus", StationMetadataSendingStatus)
            ComUpdateTblStation.Parameters.AddWithValue("@HasSentToSwitch", StationMetadataSendingStatus)
            ' WriteICSDBlogs(comMain.CommandText)
            ComUpdateTblStation.ExecuteNonQuery()

            'Dim strcommand As String = ComUpdateTblStation.CommandText
            'InserttblSQLCommandQ(MyCon, strcommand)
        Catch ex As Exception
            WriteLogs("updatetblSendAndRecievedSendToSwich" & ex.Message)
        Finally
            If MyCon.State = ConnectionState.Open Then
                MyCon.Close()
            End If
        End Try
    End Function
    Public Function UpDatetblJoinSensorType_LoggerModuleValue(ByVal WMOCode As Long, ByVal Value As Single, ByVal JoinSensorType_LoggerModuleID As Long, ByVal strTime As String, ByVal strDate As String) As Long
        Dim SysDateNow As String = Format(System.DateTime.Now, "yyyy/MM/dd")
        Dim SysTimeNow As String = Microsoft.VisualBasic.Format(System.DateTime.Now, "HH:mm")
        Dim comMain As New SqlClient.SqlCommand
        Dim conMain As New SqlClient.SqlConnection
        Try
            conMain.ConnectionString = My.Settings.ICSDBConnectionString
            If conMain.State = ConnectionState.Closed Then
                conMain.Open()
            End If
            comMain.Connection = conMain
            '  comMain.CommandType = CommandType.Text   'usp_tblJoinSensorType_LoggerModule_UpdateValues
            comMain.CommandType = CommandType.StoredProcedure
            comMain.CommandText = "usp_UpdateValueOfSensorsByJoinSensorType_LoggerModuleID"
            comMain.Parameters.AddWithValue("@SensorValue", Value)
            comMain.Parameters.AddWithValue("@LastValueUpdatedDateTime", CDate(strDate & " " & strTime))
            comMain.Parameters.AddWithValue("@JoinSensorType_LoggerModuleID", JoinSensorType_LoggerModuleID)
            comMain.Parameters.AddWithValue("@MonthNum", CDate(strDate & " " & strTime).Month)
            comMain.ExecuteNonQuery()
        Catch ex As Exception
            WriteLogs("UpDatetblJoinSensorType_LoggerModuleValue" & ex.Message)
        Finally
        End Try
    End Function
    Public Sub usp_tblAlarms_Insert(alarms As clsLogger.ALARM, WMOCODE As String)
        'usp_tblAlarms_Insert

        Dim comMain As New SqlClient.SqlCommand
        Dim conMain As New SqlClient.SqlConnection
        Try
            conMain.ConnectionString = My.Settings.ICSDBConnectionString
            If conMain.State = ConnectionState.Closed Then
                conMain.Open()
            End If
            comMain.Connection = conMain
            '  comMain.CommandType = CommandType.Text   'usp_tblJoinSensorType_LoggerModule_UpdateValues
            comMain.CommandType = CommandType.StoredProcedure
            comMain.CommandText = "usp_tblAlarms_Insert"
            comMain.Parameters.AddWithValue("@AlarmDateTime", CDate(alarms.A_DateTime))
            comMain.Parameters.AddWithValue("@AlamName", alarms.A_Name)
            comMain.Parameters.AddWithValue("@AlarmValue", alarms.A_Value)
            comMain.Parameters.AddWithValue("@interval", CStr(alarms.A_Interval))
            comMain.Parameters.AddWithValue("@WMOCode", WMOCODE)
            comMain.ExecuteNonQuery()
        Catch ex As Exception
            WriteLogs("usp_tblAlarms_Insert" & ex.Message)
        Finally
        End Try
    End Sub


    Public Function InsetIntotblSamplingwithQuality(ByVal tblSampling As DataTable, ByVal QualityControlLevel As Integer) As Boolean
        'db4
        InsetIntotblSamplingwithQuality = False
        Dim tblSensors As New DataTable
        Dim m_DataAdapter As New SqlClient.SqlDataAdapter
        Dim m_Dataset As New DataSet
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = My.Settings.ICSDBConnectionString
        Dim CMD As New SqlClient.SqlCommand
        CMD.Connection = m_CON
        If m_CON.State = ConnectionState.Closed Then
            m_CON.Open()
        End If
        CMD.CommandType = CommandType.StoredProcedure
        m_DataAdapter.SelectCommand = CMD
        Dim err As String = ""
        Dim WMOCode As String

        If tblSampling.Rows.Count > 0 Then
            WMOCode = tblSampling.Rows(0).Item("WMOCode")
            For i As Integer = 0 To tblSampling.Rows.Count - 1
                With tblSampling.Rows(i)
                    Try
                        If Not IsDBNull(tblSampling.Rows(i).Item("Value")) Then
                            If tblSampling.Rows(i).Item("Value") <> "" Then


                                CMD.Parameters.Clear()




                                CMD.CommandText = "usp_tblSamplings_Insert_For_Center_RTUWithQuality"
                                '   CMD.Parameters.AddWithValue("@SamplingID", Convert.ToDecimal(tblSampling.Rows(i).Item("SamplingID")))
                                CMD.Parameters.AddWithValue("@SampleDateTime", Convert.ToDateTime(tblSampling.Rows(i).Item("SampleDateTime")))
                                CMD.Parameters.AddWithValue("@Value", Convert.ToDouble(.Item("Value")))
                                '  CMD.Parameters.AddWithValue("@QualityControlLevel", CInt(tblSampling.Rows(i).Item("QualityControlLevel")))
                                CMD.Parameters.AddWithValue("@FK_SendAndReceiveLogID", CLng(tblSampling.Rows(i).Item("FK_SendAndReceiveLogID")))
                                CMD.Parameters.AddWithValue("@FK_JoinSensorType_LoggerModuleID", CLng(tblSampling.Rows(i).Item("FK_JoinSensorType_LoggerModuleID")))
                                CMD.Parameters.AddWithValue("@WMOCode", .Item("WMOCode"))
                                '    CMD.Parameters.AddWithValue("@MonthNum", Convert.ToDateTime(tblSampling.Rows(i).Item("SampleDateTime")).Month)
                                CMD.Parameters.AddWithValue("@QualityControlLevel", QualityControlLevel)
                                CMD.ExecuteNonQuery()
                                err = ""
                                ' InsetIntotblSampling1 = True
                            End If
                        End If
                    Catch ex As Exception
                        err = ex.Message
                        WriteLogs("InsetIntotblSampling1:" & ex.Message)
                        WriteLogs(" CLng(tblSampling.Rows(i).Item(""FK_JoinSensorType_LoggerModuleID""):" & CLng(tblSampling.Rows(i).Item("FK_JoinSensorType_LoggerModuleID")))
                    End Try
                    Try
                        If err.Contains("duplicate key") Then
                            err = ""
                            '   MsgBox("")
                            '   WriteLogs("Update")
                            CMD.CommandText = "usp_tblSamplings_Update_Center"
                            CMD.ExecuteNonQuery()
                            '  WriteLogs("update")
                            InsetIntotblSamplingwithQuality = True
                        ElseIf err <> "" Then
                            Return False
                        End If
                    Catch ex As Exception
                        WriteLogs("updatetblSampling:" & ex.Message)
                        Return False
                    End Try
                End With

            Next



        End If
        m_CON.Close()
        Return True
    End Function

    Public Sub checkDirectory()
        Dim dirPath As String = Format(System.DateTime.Now, "yyyyMMdd")
        Try

            If Not Directory.Exists(ErrorLogADD & "\Logs\" & dirPath) Then
                Directory.CreateDirectory(ErrorLogADD & "\Logs\" & dirPath)
            End If
            Dim dirold As String = Format(System.DateTime.Now.AddDays(-7), "yyyyMMdd")
            If Directory.Exists(ErrorLogADD & "\Logs\" & dirold) Then
                Dim files()
                files = Directory.GetFiles(ErrorLogADD & "\Logs\" & dirold)
                For Each File In files
                    System.IO.File.Delete(File)
                Next
                Directory.Delete(ErrorLogADD & "\Logs\" & dirold)
            End If
        Catch ex As Exception

        End Try
    End Sub


#Region "GPRS"
    Public Function GetLastSampleDateTimeBYWMOCode(ByVal WMOCODE As String) As String

        Dim dr As SqlClient.SqlDataReader
        Dim CMD As New SqlClient.SqlCommand
        Dim m_CON As New SqlClient.SqlConnection
        m_CON.ConnectionString = m_Connection.ConnectionString
        Try
            If m_CON.State = ConnectionState.Closed Then
                m_CON.Open()
            End If


            CMD.Connection = m_CON
            CMD.CommandType = CommandType.StoredProcedure

            CMD.CommandText = "usp_GetLastSampleDateTimebyWMOCode"
            CMD.CommandTimeout = 0
            CMD.Parameters.AddWithValue("@WMOCode", WMOCODE)
            dr = CMD.ExecuteReader
            If dr.HasRows Then
                dr.Read()
                If dr.Item(0) = Nothing Then
                    GetLastSampleDateTimeBYWMOCode = Format(System.DateTime.Now.AddDays(-1), "yyyy/MM/dd HH:mm")
                Else
                    GetLastSampleDateTimeBYWMOCode = Format(dr.Item("SampleDateTime"), "yyyy/MM/dd HH:mm")

                End If

            Else
                GetLastSampleDateTimeBYWMOCode = Format(System.DateTime.Now.AddDays(-1), "yyyy/MM/dd HH:mm")

            End If
            '     MsgBox()
        Catch ex As Exception
            GetLastSampleDateTimeBYWMOCode = Format(System.DateTime.Now.AddDays(-1), "yyyy/MM/dd HH:mm")
            '  WriteDBError("GetLastSampleDateTimeBYWMOCode" & ex.Message)

        Finally
            dr.Close()
            m_CON.Close()
        End Try
        '  WriteDBError("GetLastSampleDateTimeBYWMOCode " & WMOCODE & GetLastSampleDateTimeBYWMOCode)
        '   System.Threading.Thread.Sleep(200)
        '  tol.Se(GetStationCodeByName, "")
    End Function
    Public Function usp_tblJoinSensorType_LoggerModule_GetValue(ByVal StationCode As Long, ByVal Sensor As String) As Long

        Dim SenName As String = ""
        Dim MyCon As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Dr As SqlClient.SqlDataReader
        Try
            MyCon.ConnectionString = My.Settings.ICSDBConnectionString
            If MyCon.State = ConnectionState.Closed Then
                MyCon.Open()
            End If
            CMD.Connection = MyCon
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = "usp_tblJoinSensorType_LoggerModule_GetValue"
            SenName = Sensor

            If MyCon.State = ConnectionState.Closed Then
                MyCon.Open()
            End If
            CMD.Parameters.AddWithValue("@SensorAliasName", SenName)
            CMD.Parameters.AddWithValue("@WMocode", StationCode)

            Dr = CMD.ExecuteReader()
            If Not Dr.HasRows Then
                usp_tblJoinSensorType_LoggerModule_GetValue = 0
            Else
                Dr.Read()
                usp_tblJoinSensorType_LoggerModule_GetValue = Dr.Item("SensorValue")
            End If

            MyCon.Close()
        Catch ex As Exception
            '     WriteDBError("usp_tblJoinSensorType_LoggerModule_GetID" & ex.Message)
            Dim a As Long = 3
        End Try
    End Function
    Public Function GetSenTypeIDByStation_AND_SenNameINICSDB(ByVal StationCode As Long, ByVal Sensor As String, ByRef Min As Single, ByRef Max As Single) As Long

        Dim SenName As String = ""
        Dim MyCon As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        Dim Dr As SqlClient.SqlDataReader
        Try
            MyCon.ConnectionString = My.Settings.ICSDBConnectionString
            If MyCon.State = ConnectionState.Closed Then
                MyCon.Open()
            End If
            CMD.Connection = MyCon
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = "usp_tblJoinSensorType_LoggerModule_GetID"
            SenName = Sensor

            If MyCon.State = ConnectionState.Closed Then
                MyCon.Open()
            End If
            CMD.Parameters.AddWithValue("@SensorAliasName", SenName)
            CMD.Parameters.AddWithValue("@WMocode", StationCode)

            Dr = CMD.ExecuteReader()
            If Not Dr.HasRows Then
                GetSenTypeIDByStation_AND_SenNameINICSDB = 0
            Else
                Dr.Read()
                GetSenTypeIDByStation_AND_SenNameINICSDB = Dr.Item("JoinSensorType_LoggerModuleID")
            End If

            MyCon.Close()
        Catch ex As Exception
            '    WriteDBError("GetSenTypeIDByStation_AND_SenNameINICSDB" & ex.Message)
            GetSenTypeIDByStation_AND_SenNameINICSDB = 0
        End Try
    End Function
    Public Function GetStationCodeByName(ByVal StationName As String) As Long
        'p27
        Dim DA As New SqlClient.SqlDataAdapter
        Dim con As New SqlClient.SqlConnection
        Dim com As New SqlClient.SqlCommand
        Dim DS As New DataSet
        ' Dim DT As New DataTable
        Try
            con.ConnectionString = My.Settings.ICSDBConnectionString
            com.Connection = con
            com.CommandType = CommandType.StoredProcedure
            com.CommandTimeout = 0
            DS.Clear()
            '  WriteErrorLogRTU("usp_GetStationWMOCODEByNameForRTU" & StationName)
            com.CommandText = "usp_GetStationWMOCODEByNameForRTU"
            com.Parameters.AddWithValue("@stationName", StationName)
            DA.SelectCommand = com
            DA.Fill(DS, "GetStationCodeByName")
            If DS.Tables("GetStationCodeByName").Rows.Count > 0 Then
                GetStationCodeByName = DS.Tables("GetStationCodeByName").Rows(0).Item("WMOCode")

            End If
        Catch ex As Exception
            '  WriteDBError("GetStationCodeByName" & ex.Message)
        End Try

    End Function
    'ups_GetStation_LoggerType
    Public Function GetStationIDByCode(ByVal WMOCode As Long, ByRef lngError As Long) As String
        Dim comMain As New SqlClient.SqlCommand
        Dim Mycon As New SqlClient.SqlConnection
        Dim rdr As SqlClient.SqlDataReader
        GetStationIDByCode = ""
        Try
            Mycon.ConnectionString = My.Settings.ICSDBConnectionString
            If Mycon.State = ConnectionState.Closed Then
                Mycon.Open()
            End If
            comMain.Connection = Mycon
            comMain.CommandType = CommandType.StoredProcedure
            comMain.CommandText = "usp_tblStations_SelectRowByWMOCode"
            comMain.Parameters.AddWithValue("@WMOCode", WMOCode)
            rdr = comMain.ExecuteReader
            If Not rdr.HasRows Then
                lngError = -1
            Else
                rdr.Read()
                GetStationIDByCode = rdr.Item("IKAOCode").ToString
                lngError = 0
            End If
        Catch ex As Exception
            '    WriteDBError("GetStationIDByCode" & ex.Message)
            lngError = -1
        End Try
    End Function
    Public Function InsertReceivedFromStationINICDB(ByVal strFile As String, ByVal WMOcode As Long, ByVal strDate As String, ByVal strTime As String, ByVal SendAndReceiveLogID As String, ByVal FKFileType As Integer) As String
        Dim MyCon As New SqlClient.SqlConnection
        Dim comMain As New SqlClient.SqlCommand
        Dim DataAdaptMain As New SqlClient.SqlDataAdapter
        Dim DSMain As New DataSet
        InsertReceivedFromStationINICDB = 0
        Try
            MyCon.ConnectionString = My.Settings.ICSDBConnectionString
            If MyCon.State = ConnectionState.Closed Then
                MyCon.Open()
            End If
            comMain.Connection = MyCon
            comMain.CommandType = CommandType.StoredProcedure




            comMain.CommandText = "usp_tblSendAndReceiveLogs_InsertCenter"
            '   "INSERT INTO tblSendAndReceiveLogs (FileNameAndAddress,LogDateTime,SendOrReceive,FileSize,FK_StationID,HasStoredInDB) VALUES ('" & strFile & "','" & CDate(strDate & "  " & strTime) & "','1','" & Microsoft.VisualBasic.FileLen(strFile) & "','" & lngStationCode & "','0')" & _
            ' "   select SCOPE_IDENTITY() "
            comMain.Parameters.AddWithValue("@FileNameAndAddress", strFile)
            comMain.Parameters.AddWithValue("@LogDateTime", CDate(strDate & "  " & strTime))
            comMain.Parameters.AddWithValue("@SendOrReceive", 1)
            comMain.Parameters.AddWithValue("@FileSize", Microsoft.VisualBasic.FileLen(strFile))
            comMain.Parameters.AddWithValue("@WMOCODE", WMOcode)
            comMain.Parameters.AddWithValue("@HasStoredInDB", 0)
            comMain.Parameters.AddWithValue("@StationMetadataSendingStatus", 0)
            comMain.Parameters.AddWithValue("@HasSentToSwitch", 0)

            comMain.Parameters.AddWithValue("@SendAndReceiveLogID", SendAndReceiveLogID) 'eslah
            comMain.Parameters.AddWithValue("@Fk_FileType", FKFileType)
            ' InsertReceivedFromStationINICDB = comMain.ExecuteScalar()
            InsertReceivedFromStationINICDB = CStr(comMain.ExecuteScalar())
        Catch ex As Exception
            InsertReceivedFromStationINICDB = ""
            '  WriteDBError("InsertReceivedFromStationINICDB" & ex.Message)
        Finally
            If MyCon.State = ConnectionState.Open Then
                MyCon.Close()
            End If
        End Try
    End Function
    Public Function usp_GetStation_LoggerType(ByVal StationName As String) As String
        Dim con As New SqlClient.SqlConnection
        Dim CMD As New SqlClient.SqlCommand
        con.ConnectionString = My.Settings.ICSDBConnectionString
        CMD.Connection = con
        con.Open()
        CMD.CommandType = CommandType.StoredProcedure
        CMD.CommandText = "usp_GetStation_LoggerType"

        CMD.Parameters.AddWithValue("@STationName", StationName)

        Try

            usp_GetStation_LoggerType = CMD.ExecuteScalar
            ' WriteLogs("IsDateExist count" & count)
            '     WriteLogs("dt " & dt)


        Catch ex As Exception
            '  dt = System.DateTime.Now.AddDays(-1)
            WriteLogs("usp_GetStation_LoggerType:" & ex.Message)
        End Try
    End Function
    Public Sub GETStationInfoByName(ByRef tblInfo As DataTable, ByVal StationName As String)
        Try

            Dim CMD As New SqlClient.SqlCommand
            Dim m_DataAdapter As New SqlClient.SqlDataAdapter
            Dim m_CON As New SqlClient.SqlConnection
            Dim m_Dataset As New DataSet
            CMD.Connection = m_CON
            m_CON.ConnectionString = m_Connection.ConnectionString
            m_CON.Open()
            CMD.CommandType = CommandType.StoredProcedure
            CMD.CommandText = "usp_tblStations_SelectRowByName"
            '@DirectoryPath,@DirectorySize,@DrectoryModify
            m_DataAdapter.SelectCommand = CMD
            CMD.Parameters.AddWithValue("@StationName", StationName)
            m_DataAdapter.Fill(m_Dataset, "tblInfo")
            tblInfo = m_Dataset.Tables("tblInfo")
        Catch ex As Exception
            WriteLogs("GETStationInfoByName" & ex.Message)
        End Try
        '     System.DateTime.Now.Month
    End Sub

    'usp_tblStations_SelectRowByName
    Public Sub WriteLogs(ByVal log As String) '
        '    WriteFileLogError("WriteLogs create path  ")
        Dim dirPath As String = Format(System.DateTime.Now, "yyyyMMdd")
        checkDirectory()

        Try





            Dim sw As New StreamWriter(ApplicationAdd & "\Logs\" & dirPath & "\DBError_" & Format(System.DateTime.Now, "yyyyMMdd") & ".txt", True)
            sw.WriteLine(System.DateTime.Now.ToString & ":" & log)
            sw.Close()


        Catch ex As Exception
            'WriteFileLogError("WriteLogs  " & ex.Message)
            'WriteFileLogError("Log is  " & log)
        End Try

    End Sub
#End Region
End Class
