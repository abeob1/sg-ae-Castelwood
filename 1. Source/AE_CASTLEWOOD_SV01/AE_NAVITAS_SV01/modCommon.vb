Imports System.Configuration
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.OleDb
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.Odbc


Module modCommon

    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing information about the initialing variables
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSqlstr As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            ''  Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.sLicenseServer = String.Empty
            oCompDef.iServerLanguage = 3
            'oCompDef.iServerType = 7
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty
            oCompDef.sSeries = String.Empty


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenseServer")) Then
                oCompDef.sLicenseServer = ConfigurationManager.AppSettings("LicenseServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("InboxDir")) Then
                oCompDef.sInboxDir = ConfigurationManager.AppSettings("InboxDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FailDir")) Then
                oCompDef.sFailDir = ConfigurationManager.AppSettings("FailDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = ConfigurationManager.AppSettings("LogPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailFrom")) Then
                oCompDef.sEmailFrom = ConfigurationManager.AppSettings("EmailFrom")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo")) Then
                oCompDef.sEmailTo = ConfigurationManager.AppSettings("EmailTo")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailSubject")) Then
                oCompDef.sEmailSubject = ConfigurationManager.AppSettings("EmailSubject")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPServer")) Then
                oCompDef.sSMTPServer = ConfigurationManager.AppSettings("SMTPServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPort")) Then
                oCompDef.sSMTPPort = ConfigurationManager.AppSettings("SMTPPort")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPUser")) Then
                oCompDef.sSMTPUser = ConfigurationManager.AppSettings("SMTPUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPassword")) Then
                oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("SMTPPassword")
            End If


            '' Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Public Function ExecuteSQLQuery_DT(ByVal sQuery As String, ByRef sErrDesc As String) As DataTable

        Dim sConnString As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_sSAPEntityName
        Dim oHanaOdbcConnection As New OdbcConnection(sConnString)
        Dim oHanaOdbcCommand As New OdbcCommand()
        Dim oDataset As New DataSet()


        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "HANAtoDatatable()"

            If p_iDebugMode = DEBUG_ON Then
                WriteToLogFile_Debug("Starting Function ", sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then
                WriteToLogFile_Debug("Query : " + sQuery, sFuncName)
            End If

            If oHanaOdbcConnection.State = ConnectionState.Closed Then
                oHanaOdbcConnection.Open()
            End If

            oHanaOdbcCommand.CommandType = CommandType.Text
            oHanaOdbcCommand.CommandText = sQuery
            oHanaOdbcCommand.Connection = oHanaOdbcConnection
            oHanaOdbcCommand.CommandTimeout = 0
            Dim oHanaDA As New OdbcDataAdapter(oHanaOdbcCommand)
            oHanaDA.Fill(oDataset)
            oHanaDA.Dispose()

            If p_iDebugMode = DEBUG_ON Then
                WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)
            End If
            sErrDesc = String.Empty

            Return oDataset.Tables(0)
        Catch Ex As Exception

            sErrDesc = Ex.Message.ToString()
            If p_iDebugMode = DEBUG_ON Then
                WriteToLogFile(sErrDesc, sFuncName)
            End If
            If p_iDebugMode = DEBUG_ON Then
                WriteToLogFile_Debug("Completed With ERROR  ", sFuncName)
            End If
            Return Nothing
            Throw Ex
        Finally
            oHanaOdbcCommand.Dispose()
            oHanaOdbcConnection.Close()

            oHanaOdbcConnection.Dispose()
        End Try
        '' Return oDataset.Tables(0)
    End Function

    Public Function ExecuteQueryReturnDataTable_HANA(ByVal sQueryString As String, ByVal sCompanyDB As String) As DataTable

        Dim sFuncName As String = "ExecuteQueryReturnDataTable_HANA"
        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sCompanyDB

        Dim oCmd As New System.Data.Odbc.OdbcCommand
        Dim oDS As DataSet = New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()
        Dim dtDetail As DataTable = New DataTable


        Try
            Con.ConnectionString = sConstr
            Con.Open()

            oCmd.CommandText = CommandType.Text
            oCmd.CommandText = sQueryString
            oCmd.Connection = Con
            oCmd.CommandTimeout = 0

            Dim da As New System.Data.Odbc.OdbcDataAdapter(oCmd)
            da.Fill(dtDetail)
            dtDetail.TableName = "Data"

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try

        ExecuteQueryReturnDataTable_HANA = dtDetail

    End Function


    Public Function IdentifyExcelFile_JournalEntry(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   IdentifyExcelFile_JournalEntry()
        '   Purpose     :   This function will identify the Excel file of Journal Entry
        '                    Upload the file into Dataview and provide the information to post transaction in SAP.
        '                     Transaction Success : Move the Excel file to SUCESS folder
        '                     Transaction Fail :    Move the Excel file to FAIL folder and send Error notification to concern person
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************


        Dim sSqlstr As String = String.Empty
        Dim bJEFileExist As Boolean
        Dim sFileType As String = String.Empty
        Dim oDTDistinct As DataTable = Nothing
        Dim oDTRowFilter As DataTable = Nothing
        Dim oDSJE As DataSet = Nothing
        Dim oDICompany As SAPbobsCOM.Company = Nothing

        Dim sFuncName As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim oDVLineTable As DataView = Nothing
        Dim oDTHeader As DataTable = Nothing
        Dim sEntity As String = String.Empty
        Dim sDTError As DataTable = Nothing
        Dim sBatchno As String = String.Empty
        Dim sTransno As String = String.Empty
        Dim sSplit() As String
        Dim sSeries As String = String.Empty
        Dim sDocEntry As String = String.Empty


        Try
            sFuncName = "IdentifyExcelFile_JournalEntry()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)


            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("*.xls")

            For Each File As System.IO.FileInfo In files
                bJEFileExist = True
                Console.WriteLine("Reading File - " & File.Name, sFuncName)

                'sSplit = File.Name.ToString.Split("_")
                'sSeries = sSplit(1)
                'sSeries = Right(Left(sSeries, 8), 4)
                'sSeries = "XN-" & Right(sSeries, 2) & Left(sSeries, 2)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting File Name - " & File.Name, sFuncName)
                'sFileType = Replace(File.Name, ".txt", "").Trim
                'upload the CSV to Dataview

                ''  Console.WriteLine("Calling GetDataViewFromExcel() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetDataViewFromExcel() ", sFuncName)
                oDSJE = GetDataViewFromExcel(File.FullName, File.Name, File.Extension, sDTError)

                If sDTError.Rows.Count > 0 Then
                    Write_TextFile(sDTError, p_oCompDef.sLogPath, sErrDesc)
                    FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
                    Throw New ArgumentException("GL Accounts are not defined in SAP")
                End If

                '' oDVLineTable = oDSJE.Tables(1).DefaultView
                ''oDTHeader = oDSJE.Tables(0)
                Console.WriteLine("Connecting to the Entity " & p_sSAPEntityName, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                oDICompany = New SAPbobsCOM.Company
                If ConnectToTargetCompany(oDICompany, p_sSAPEntityName, sErrDesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrDesc)
                End If
                Console.WriteLine("Connected Successfully " & p_sSAPEntityName, sFuncName)
                oRS = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sBatchno = oDSJE.Tables(0).Rows(0).Item(2).ToString.Trim
                oRS.DoQuery("SELECT T0.""BatchNum"" , T0.""TransId"" FROM OJDT T0 WHERE T0.""Ref1""  = '" & sBatchno & "'")
                If oRS.RecordCount > 0 Then
                    sTransno = oRS.Fields.Item("TransId").Value
                    Throw New ArgumentException("File already Upload, Journal Entry No. " & sTransno)
                End If

                Console.WriteLine("Processing Journal Entry " & File.FullName, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function JournalEntry_Posting() ", sFuncName)
                sDocEntry = String.Empty
                If JournalEntry_Posting(oDSJE, oDICompany, File.Name, sDocEntry, sErrDesc) <> RTN_SUCCESS Then
                    Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    'AddDataToTable(p_oDtError, File.Name, "Error", sErrDesc)
                    FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
                    AddDataToTable(p_oDtError, File.Name, "Error", sErrDesc)
                Else

                    Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    FileMoveToArchive(File, File.FullName, RTN_SUCCESS, "")
                    AddDataToTable(p_oDtSuccess, File.Name, sDocEntry, "Success")
                End If

            Next

            If bJEFileExist = False Then
                Console.WriteLine("No input file found  ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No input file found ", sFuncName)
            End If

            'Send Error Email if Datable has rows.
            If p_oDtError.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
                EmailTemplate_Error()
            End If
            p_oDtError.Rows.Clear()

            'Send Success Email if Datable has rows..
            If p_oDtSuccess.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Success()", sFuncName)
                EmailTemplate_Success()
            End If
            p_oDtSuccess.Rows.Clear()




            Console.WriteLine("Completed With SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            IdentifyExcelFile_JournalEntry = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed With ERROR", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            IdentifyExcelFile_JournalEntry = RTN_ERROR

        End Try

    End Function

    Public Function SendEmailNotification(ByVal CurrFileToUpload As String, ByVal sCompanyCode As String, _
                                          ByVal sCompanyName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oSmtpServer As New SmtpClient()
        Dim oMail As New MailMessage
        Dim p_SyncDateTime As String = String.Empty

        Try
            sFuncName = "SendEmailNotification()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            Console.WriteLine("Sending Mail To : ")


            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
            '--------- Message Content in HTML tags
            Dim sBody As String = String.Empty

            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & " Dear Sir/Madam,<br /><br />"
            sBody = sBody & p_SyncDateTime & " <br /><br />"
            sBody = sBody & " " & "Please find the attached FAILED document in SAP and followed by the ERROR.<br /><br /> "
            sBody = sBody & " " & " Company Code : " & sCompanyCode & "<br /> "
            sBody = sBody & " " & " Company Name : " & sCompanyName & " <br /> "
            sBody = sBody & "<br /> <font color=""red""> Error Message : " & sErrDesc & "</font><br />"
            sBody = sBody & "<br /><br />"
            sBody = sBody & " Please do not reply to this email. <div/>"


            ''<font size="3" color="red">This is some text!</font>

            Dim attachment As System.Net.Mail.Attachment
            attachment = New System.Net.Mail.Attachment(CurrFileToUpload)
            oMail.Attachments.Add(attachment)


            oSmtpServer.Credentials = New Net.NetworkCredential(p_oCompDef.sSMTPUser, p_oCompDef.sSMTPPassword)
            oSmtpServer.Port = p_oCompDef.sSMTPPort '587
            oSmtpServer.Host = p_oCompDef.sSMTPServer '"smtp.gmail.com"
            oSmtpServer.EnableSsl = True
            oMail.From = New MailAddress(p_oCompDef.sEmailFrom) '("sapb1.abeoelectra@gmail.com")
            oMail.To.Add(p_oCompDef.sEmailTo)
            ' oMail.Attachments.Add(New Attachment(sfileName192.168.1.4
            oMail.Subject = "Reg., Error While Uploading Journal Entry. "
            oMail.Body = sBody
            oMail.IsBodyHtml = True

            oSmtpServer.Send(oMail)
            oMail.Dispose()
            Console.WriteLine("Sending Mail Completed Successfully to this EmailID : " & p_oCompDef.sEmailTo)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Notification Sent to " & p_oCompDef.sEmailTo, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SendEmailNotification = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            oMail.Dispose()
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            SendEmailNotification = RTN_ERROR
        Finally
            oMail.Dispose()

        End Try

    End Function

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                          ByVal sEntity As String, _
                                          ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectToTargetCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2013 21
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDT As New DataTable
        Dim sSAPUser As String = String.Empty
        Dim sSAPPWd As String = String.Empty

        Dim sTrgtDBName As String = String.Empty
        Try
            sFuncName = "ConnectToTargetCompany()"
            ''  Console.WriteLine("Starting function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)


            If String.IsNullOrEmpty(p_sSAPEntityName) Then
                sErrDesc = "No Database login information found in COMPANYDATA Table. Please check"
                Console.WriteLine("No Database login information found in Template. Please check ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Database login information found in COMPANYDATA Table. Please check", sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sSQL = "SELECT * FROM ""@AI_TB01_COMPANYDATA""  WHERE ""Code"" = '" & sEntity & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)

                oDT = ExecuteSQLQuery_DT(sSQL, sErrDesc)

                If oDT.Rows.Count > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
                    '' Console.WriteLine("Initializing the Company Object ", sFuncName)
                    oCompany = New SAPbobsCOM.Company

                    sTrgtDBName = oDT.Rows(0).Item("Name").ToString
                    sSAPUser = oDT.Rows(0).Item("U_SAPUSER").ToString
                    sSAPPWd = oDT.Rows(0).Item("U_SAPPASSWORD").ToString

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)
                    ''  Console.WriteLine("Assigning the representing database name ", sFuncName)
                    oCompany.Server = p_oCompDef.sServer
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                    oCompany.LicenseServer = p_oCompDef.sLicenseServer
                    oCompany.CompanyDB = sTrgtDBName
                    oCompany.UserName = sSAPUser
                    oCompany.Password = sSAPPWd

                    oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

                    oCompany.UseTrusted = False

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
                    Console.WriteLine("Connecting to the Company Database. ", sFuncName)
                    iRetValue = oCompany.Connect()

                    If iRetValue <> 0 Then
                        oCompany.GetLastError(iErrCode, sErrDesc)

                        sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                            oCompany.CompanyDB, System.Environment.NewLine, _
                                        vbTab, sErrDesc)

                        Throw New ArgumentException(sErrDesc)
                    End If
                Else
                    sErrDesc = "No Database login information found in COMPANYDATA Table. Please check"
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            '' Console.WriteLine("Completed with SUCCESS ", sFuncName)
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

    Public Sub FileMoveToArchive(ByVal oFile As System.IO.FileInfo, ByVal CurrFileToUpload As String, ByVal iStatus As Integer, ByVal sErrDesc As String)

        'Event      :   FileMoveToArchive
        'Purpose    :   For Renaming the file with current time stamp & moving to archive folder
        'Author     :   JOHN 
        'Date       :   21 MAY 2014

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            'Dim RenameCurrFileToUpload = Replace(CurrFileToUpload.ToUpper, ".CSV", "") & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim RenameCurrFileToUpload As String = Mid(oFile.Name, 1, oFile.Name.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & ".xls"

            If iStatus = RTN_SUCCESS Then
                Console.WriteLine("Moving CSV file to success folder ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to success folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sSuccessDir & "\" & RenameCurrFileToUpload)
            Else
                Console.WriteLine("Moving CSV file to Fail folder ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to Fail folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in renaming/copying/moving ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Function Del_schema(ByVal csvFileFolder As String) As Long

        ' ***********************************************************************************
        '   Function   :    Del_schema()
        '   Purpose    :    This function is handles - Delete the Schema file
        '   Parameters :    ByVal csvFileFolder As String
        '                       csvFileFolder = Passing file name
        '   Author     :    JOHN
        '   Date       :    26/06/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Del_schema()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            Dim FileToDelete As String
            FileToDelete = csvFileFolder & "\\schema.ini"
            If System.IO.File.Exists(FileToDelete) = True Then
                System.IO.File.Delete(FileToDelete)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            Del_schema = RTN_SUCCESS
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            Del_schema = RTN_ERROR
        End Try
    End Function

    Public Function Create_schema(ByVal csvFileFolder As String, ByVal FileName As String) As Long

        ' ***********************************************************************************
        '   Function   :    Create_schema()
        '   Purpose    :    This function is handles - Create the Schema file
        '   Parameters :    ByVal csvFileFolder As String
        '                       csvFileFolder = Passing file name
        '   Author     :    JOHN
        '   Date       :    26/06/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Create_schema()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            Dim csvFileName As String = FileName
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            'Dim s1, s2, s3, s4, s5 As String

            srOutput.WriteLine("[" & csvFileName & "]")
            srOutput.WriteLine("ColNameHeader=False")
            srOutput.WriteLine("Format=CSVDelimited")
            srOutput.WriteLine("Col1=F1 Text")
            srOutput.WriteLine("Col2=F2 Text")
            srOutput.WriteLine("Col3=F3 Text")
            srOutput.WriteLine("Col4=F4 Text")
            srOutput.WriteLine("Col5=F5 Text")
            srOutput.WriteLine("Col6=F6 Text")
            srOutput.WriteLine("Col7=F7 Text")
            srOutput.WriteLine("Col8=F8 Text")
            srOutput.WriteLine("Col9=F9 Text")
            srOutput.WriteLine("Col10=F10 Double")
            srOutput.WriteLine("Col11=F11 Text")
            srOutput.WriteLine("Col12=F12 Double")
            srOutput.WriteLine("Col13=F13 Text")
            srOutput.WriteLine("Col14=F14 Text")
            srOutput.WriteLine("Col15=F15 Text")
            srOutput.WriteLine("MaxScanRows=0")
            srOutput.WriteLine("CharacterSet=OEM")
            'srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf)
            srOutput.Close()
            fsOutput.Close()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            Create_schema = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            Create_schema = RTN_ERROR
        End Try

    End Function

    Public Function GetDataViewFromExcel(ByVal CurrFileToUpload As String, ByVal Filename As String, ByVal sExtension As String, ByRef oDTError As DataTable) As DataSet

        ' **********************************************************************************
        '   Function    :   GetDataViewFromExcel()
        '   Purpose     :   This function will upload the data from Excel file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim oDTHeader As New DataTable
        Dim oDTRows As New DataTable
        Dim oDSResult As New DataSet


        Dim sDataBase As String = String.Empty
        Dim sEntityName As String = String.Empty
        'Dim sBatchNo As String = String.Empty
        Dim sBatchNo As Double = 0
        'Dim dtTransactionDate As String = String.Empty
        Dim dtTransactionDate As New DateTime
        Dim sXNDocFrom As String = String.Empty
        Dim sXNDocTO As String = String.Empty

        Dim sAcctCode As String = String.Empty
        Dim dDebit As Double = 0
        Dim dCredit As Double = 0
        Dim sCostCenter As String = String.Empty
        Dim sRemarks As String = String.Empty
        Dim sQuery As String = String.Empty

        Dim sBUCode As String = String.Empty
        Dim sLOS As String = String.Empty
        Dim sEntity As String = String.Empty
        Dim iCount As Integer = 1
        Dim sGLAccount As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sSplit() As String
        Dim conStr As String = String.Empty
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)


        Dim cmdExcel As New OleDbCommand()
        Dim oda As New OleDbDataAdapter()
        Dim dt As New DataTable()
        Dim dt_B As New DataTable()
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Object Initializing  ", sFuncName)
        oDTHeader.Columns.Add("EntityName", GetType(String))
        oDTHeader.Columns.Add("Entity", GetType(String))
        oDTHeader.Columns.Add("BatchNo", GetType(Double))
        oDTHeader.Columns.Add("TransactionDate", GetType(Date))
        oDTHeader.Columns.Add("XNDocFrom", GetType(String))
        oDTHeader.Columns.Add("XNDocTo", GetType(String))

        oDTRows.Columns.Add("GLAccount", GetType(String))
        oDTRows.Columns.Add("Debit", GetType(Double))
        oDTRows.Columns.Add("Credit", GetType(Double))
        oDTRows.Columns.Add("CostCenter", GetType(String))
        oDTRows.Columns.Add("Remarks", GetType(String))
        oDTError = New DataTable
        oDTError.Columns.Add("GLAccount", GetType(String))
        oDTError.Columns.Add("Error", GetType(String))

        sFuncName = "GetDataViewFromExcel"

        Try
            ''  Console.WriteLine("Starting Function ", sFuncName)
         


            Select Case sExtension
                Case ".xls"
                    'Excel 97-03
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'"
                    Exit Select
                Case ".xlsx"
                    'Excel 07
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1'"
                    Exit Select
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection String  " & conStr, sFuncName)
            Dim connExcel As New OleDbConnection(conStr)
            cmdExcel.Connection = connExcel

            'Get the name of First Sheet
            connExcel.Open()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection Open  ", sFuncName)
            Dim dtExcelSchema As DataTable
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
            connExcel.Close()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection Close  ", sFuncName)
            'Read Data from First Sheet
            connExcel.Open()
            '  cmdExcel.CommandText = "SELECT * From [" & SheetName & "]" [w1$A10:B10]
            cmdExcel.CommandText = "SELECT * From [" & SheetName & "A3:B3]"
            oda.SelectCommand = cmdExcel
            dt_B = New DataTable("Data")
            oda.Fill(dt_B)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Data Filled   ", sFuncName)
            cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
            'cmdExcel.CommandText = "SELECT * From [" & SheetName & "A3:B3]"
            oda.SelectCommand = cmdExcel
            dt = New DataTable("Data")
            oda.Fill(dt)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Remaining Data filled  ", sFuncName)
            connExcel.Close()

            sEntityName = dt.Rows(0)(1).ToString()
            sDataBase = dt.Rows(1)(1).ToString()
            p_sSAPEntityName = String.Empty
            p_sSAPEntityName = sDataBase
            sBatchNo = dt_B.Rows(0)(1).ToString()
            'sSplit = sBatchNo.ToString.Split(".")
            'sBatchNo = sSplit(0)

            dtTransactionDate = DateTime.ParseExact(dt.Rows(3)(1).ToString(), "d/M/yyyy", Nothing)

            ''sQuery = "SELECT T0.""Code"", T0.""Name"", T0.""U_GLCode"" FROM ""@AE_XNGLMAPPING""  T0"
            ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", "GetDataViewFromExcel")
            ''p_oGLMapping = ExecuteSQLQuery_DT(sQuery, sErrDesc)
            ''If sErrDesc.Length > 0 Then
            ''    Dim sEmailError As String = String.Empty
            ''    EmailTemplate_GeneralError(sErrDesc, sEmailError)
            ''    If sEmailError.Length > 0 Then
            ''        WriteToLogFile(sEmailError, sFuncName)
            ''    End If
            ''    Throw New ArgumentException(sErrDesc)
            ''End If

            sQuery = "SELECT T0.""AcctCode"", T0.""AcctName"" FROM OACT T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", "GetDataViewFromExcel")
            p_oOACT = ExecuteSQLQuery_DT(sQuery, sErrDesc)
            If sErrDesc.Length > 0 Then
                Dim sEmailError As String = String.Empty
                EmailTemplate_GeneralError(sErrDesc, sEmailError)
                If sEmailError.Length > 0 Then
                    WriteToLogFile(sEmailError, sFuncName)
                End If
                Throw New ArgumentException(sErrDesc)
            End If



            sXNDocFrom = dt.Rows(4)(1).ToString()
            sXNDocTO = dt.Rows(5)(1).ToString()

            oDTHeader.Rows.Add(sEntityName, sDataBase, sBatchNo, dtTransactionDate, sXNDocFrom, sXNDocTO)


            Dim dvOACT As DataView = New DataView(p_oOACT)
            ''     Dim dvGLAcccountMapping As DataView = New DataView(p_oGLMapping)

            For imjs As Integer = 8 To dt.Rows.Count - 1
                If Not String.IsNullOrEmpty(dt.Rows(imjs)(0).ToString()) Then
                    sAcctCode = dt.Rows(imjs)(0).ToString()
                    dDebit = dt.Rows(imjs)(1).ToString()
                    dCredit = dt.Rows(imjs)(2).ToString()
                    sCostCenter = dt.Rows(imjs)(3).ToString()
                    sRemarks = dt.Rows(imjs)(4).ToString()

                    dvOACT.RowFilter = "AcctCode='" & sAcctCode & "'"
                    If dvOACT.Count = 0 Then
                        oDTError.Rows.Add(sAcctCode, "GL Codes are not defined in the Chart of Accounts (SAP)")
                        'dvGLAcccountMapping.RowFilter = "Code='" & sAcctCode & "'"
                        'If dvGLAcccountMapping.Count = 0 Then
                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No GL Account Found  " & sAcctCode & "  Line No " & iCount, sFuncName)
                        '    sGLAccount = ""
                        '    oDTError.Rows.Add(sAcctCode, "GL Codes are not defined in the Chart of Accounts (SAP)")
                        'Else
                        '    sGLAccount = dvGLAcccountMapping.Item(0)(2).ToString
                        'End If

                    Else
                        sGLAccount = sAcctCode
                    End If

                    oDTRows.Rows.Add(sGLAccount, dDebit, dCredit, sCostCenter, sRemarks)

                End If
            Next

            oDSResult.Tables.Add(oDTHeader)
            oDSResult.Tables.Add(oDTRows)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", "GetDataViewFromExcel")
            Return oDSResult

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while reading content of " & ex.Message, sFuncName)
            Call WriteToLogFile_Debug(ex.Message, sFuncName)
            Return Nothing
        Finally
           

        End Try
    End Function

    Public Function GetDataViewFromExcel_OLD(ByVal CurrFileToUpload As String, ByVal Filename As String, ByRef oDTError As DataTable) As DataSet

        ' **********************************************************************************
        '   Function    :   GetDataViewFromExcel()
        '   Purpose     :   This function will upload the data from Excel file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim oDTHeader As New DataTable
        Dim oDTRows As New DataTable
        Dim oDSResult As New DataSet


        Dim sDataBase As String = String.Empty
        Dim sEntityName As String = String.Empty
        'Dim sBatchNo As String = String.Empty
        Dim sBatchNo As Double = 0
        'Dim dtTransactionDate As String = String.Empty
        Dim dtTransactionDate As New DateTime
        Dim sXNDocFrom As String = String.Empty
        Dim sXNDocTO As String = String.Empty

        Dim sAcctCode As String = String.Empty
        Dim dDebit As Double = 0
        Dim dCredit As Double = 0
        Dim sCostCenter As String = String.Empty
        Dim sRemarks As String = String.Empty
        Dim sQuery As String = String.Empty

        Dim sBUCode As String = String.Empty
        Dim sLOS As String = String.Empty
        Dim sEntity As String = String.Empty
        Dim iCount As Integer = 1
        Dim sGLAccount As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sSplit() As String


        oDTError = New DataTable

        sFuncName = "GetDataViewFromExcel"

        ''  Console.WriteLine("Starting Function ", sFuncName)
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Object Initializing  ", sFuncName)
        Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
        Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim excelRng As Microsoft.Office.Interop.Excel.Range
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Object Initialized  ", sFuncName)
        Try
            ExcelWorkbook = ExcelApp.Workbooks.Open(CurrFileToUpload)
            ExcelWorkSheet = ExcelWorkbook.ActiveSheet
            excelRng = ExcelWorkSheet.Range("A1")
            Dim RowIndex As Integer = 15

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Excel file Uploaded  ", sFuncName)

            oDTHeader.Columns.Add("EntityName", GetType(String))
            oDTHeader.Columns.Add("Entity", GetType(String))
            oDTHeader.Columns.Add("BatchNo", GetType(Double))
            oDTHeader.Columns.Add("TransactionDate", GetType(Date))
            oDTHeader.Columns.Add("XNDocFrom", GetType(String))
            oDTHeader.Columns.Add("XNDocTo", GetType(String))

            oDTRows.Columns.Add("GLAccount", GetType(String))
            oDTRows.Columns.Add("Debit", GetType(Double))
            oDTRows.Columns.Add("Credit", GetType(Double))
            oDTRows.Columns.Add("CostCenter", GetType(String))
            oDTRows.Columns.Add("Remarks", GetType(String))

            oDTError.Columns.Add("GLAccount", GetType(String))
            oDTError.Columns.Add("Error", GetType(String))

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Work Sheet Defined ", sFuncName)

            While excelRng.Range("A" & RowIndex & "").Text <> "" 'And excelRng.Range("C" & RowIndex & "").Text <> ""
                RowIndex = RowIndex + 1
            End While

            sEntityName = excelRng.Range("B1").Text
            sDataBase = excelRng.Range("B2").Text
            p_sSAPEntityName = String.Empty
            p_sSAPEntityName = sDataBase
            sBatchNo = excelRng.Range("B3").Value
            sSplit = sBatchNo.ToString.Split(".")
            sBatchNo = sSplit(0)

            dtTransactionDate = DateTime.ParseExact(excelRng.Range("B4").Text, "dd/MM/yyyy", Nothing)
            ' dtTransactionDate = Convert.ToDateTime(excelRng.Range("B4").Text.ToString())
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Work Sheet Range Calculated ", sFuncName)

            sQuery = "SELECT T0.""Code"", T0.""Name"", T0.""U_GLCode"" FROM ""@AE_XNGLMAPPING""  T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", "GetDataViewFromExcel")
            p_oGLMapping = ExecuteSQLQuery_DT(sQuery, sErrDesc)
            If sErrDesc.Length > 0 Then
                Dim sEmailError As String = String.Empty
                EmailTemplate_GeneralError(sErrDesc, sEmailError)
                If sEmailError.Length > 0 Then
                    WriteToLogFile(sEmailError, sFuncName)
                End If
                Throw New ArgumentException(sErrDesc)
            End If

            sQuery = "SELECT T0.""AcctCode"", T0.""AcctName"" FROM OACT T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", "GetDataViewFromExcel")
            p_oOACT = ExecuteSQLQuery_DT(sQuery, sErrDesc)
            If sErrDesc.Length > 0 Then
                Dim sEmailError As String = String.Empty
                EmailTemplate_GeneralError(sErrDesc, sEmailError)
                If sEmailError.Length > 0 Then
                    WriteToLogFile(sEmailError, sFuncName)
                End If
                Throw New ArgumentException(sErrDesc)
            End If



            sXNDocFrom = excelRng.Range("B5").Text
            sXNDocTO = excelRng.Range("B6").Text

            oDTHeader.Rows.Add(sEntityName, sDataBase, sBatchNo, dtTransactionDate, sXNDocFrom, sXNDocTO)


            Dim dvOACT As DataView = New DataView(p_oOACT)
            Dim dvGLAcccountMapping As DataView = New DataView(p_oGLMapping)

            Dim i As Integer = 1
            For i = 9 To RowIndex - 1

                sAcctCode = excelRng.Range("A" & i & "").Text
                dDebit = excelRng.Range("B" & i & "").Text
                dCredit = excelRng.Range("C" & i & "").Text
                sCostCenter = excelRng.Range("D" & i & "").Text
                sRemarks = excelRng.Range("E" & i & "").Text

                dvOACT.RowFilter = "AcctCode='" & sAcctCode & "'"
                If dvOACT.Count = 0 Then
                    dvGLAcccountMapping.RowFilter = "Code='" & sAcctCode & "'"
                    If dvGLAcccountMapping.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No GL Account Found  " & sAcctCode & "  Line No " & iCount, sFuncName)
                        sGLAccount = ""
                        oDTError.Rows.Add(sAcctCode, "GL Codes are not defined in the Chart of Accounts (SAP)")
                    Else
                        sGLAccount = dvGLAcccountMapping.Item(0)(2).ToString
                    End If

                Else
                    sGLAccount = sAcctCode
                End If

                oDTRows.Rows.Add(sGLAccount, dDebit, dCredit, sCostCenter, sRemarks)

            Next

            oDSResult.Tables.Add(oDTHeader)
            oDSResult.Tables.Add(oDTRows)

            Return oDSResult

        Catch ex As Exception
            Return Nothing
        Finally
            ExcelWorkbook.Close()
            ExcelWorkbook = Nothing
            ExcelApp.Quit()
            ExcelApp = Nothing
            ExcelWorkSheet = Nothing
            excelRng = Nothing

        End Try
    End Function

    Public Function Write_TextFile(ByVal oDT_FinalResult As DataTable, ByVal sPAth As String, ByRef sErrDesc As String) As Long
        Try
            Dim sFuncName As String = String.Empty
            Dim irow As Integer
            Dim sFileName As String = "\SyncError.txt"
            Dim sbuffer As String = String.Empty
            Dim sline As String = "="
            sFuncName = "Write_TextFile()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            If File.Exists(sPAth & sFileName) Then
                Try
                    File.Delete(sPAth & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPAth & sFileName)
            ' Add some text to the file.

            sw.WriteLine("      ")
            sw.WriteLine("      ")
            sw.WriteLine("GL Code                 " & "Error Msg")
            sw.WriteLine(sline.PadRight(150, "="c))
            sw.WriteLine("      ")

            For imjs = 0 To oDT_FinalResult.Rows.Count - 1

                sw.WriteLine(oDT_FinalResult.Rows(imjs).Item(0).ToString & "     " & oDT_FinalResult.Rows(imjs).Item(1).ToString)

            Next imjs
            sw.Close()
            Process.Start(sPAth & sFileName)

            Write_TextFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

        Catch ex As Exception
            Write_TextFile = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, "Write_TextFile")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", "Write_TextFile")
        End Try

    End Function

    Public Function CreateDataTable(ByVal ParamArray oColumnName() As String) As DataTable
        Dim oDataTable As DataTable = New DataTable()

        Dim oDataColumn As DataColumn

        For i As Integer = LBound(oColumnName) To UBound(oColumnName)
            oDataColumn = New DataColumn()
            oDataColumn.DataType = Type.GetType("System.String")
            oDataColumn.ColumnName = oColumnName(i).ToString
            oDataTable.Columns.Add(oDataColumn)
        Next

        Return oDataTable

    End Function

    Public Sub AddDataToTable(ByVal oDt As DataTable, ByVal ParamArray sColumnValue() As String)
        Dim oRow As DataRow = Nothing
        oRow = oDt.NewRow()
        For i As Integer = LBound(sColumnValue) To UBound(sColumnValue)
            oRow(i) = sColumnValue(i).ToString
        Next
        oDt.Rows.Add(oRow)
    End Sub

End Module