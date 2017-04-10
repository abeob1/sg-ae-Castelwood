Module modMain

#Region "Variables"

    ' Company Default Structure

    Public Structure CompanyDefault

        Public sServer As String
        Public sLicenseServer As String
        Public sDBName As String
        Public sServerType As String
        Public iServerLanguage As Integer
        Public sSAPUser As String
        Public sSAPPwd As String
        Public sSAPDBName As String
        Public sDBUser As String
        Public sDBPwd As String

        Public sInboxDir As String
        Public sSuccessDir As String
        Public sFailDir As String
        Public sLogPath As String
        Public sDebug As String

        'Email Credentials

        Public sEmailFrom As String
        Public sEmailTo As String
        Public sEmailSubject As String
        Public sSMTPServer As String
        Public sSMTPPort As String
        Public sSMTPUser As String
        Public sSMTPPassword As String


        Public sSeries As String



    End Structure

    ' Return Value Variable Control
    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16 = DEBUG_ON
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault
    Public p_oCompany As SAPbobsCOM.Company
    Public sGJDBName As String = String.Empty
    Public sGJCostCenter As String = String.Empty
    Public p_sPatientType As String = String.Empty

    Public p_oDtSuccess As DataTable
    Public p_oDtError As DataTable
    Public p_oDtReport As DataTable

    Public p_sSAPEntityName As String = String.Empty
    Public p_sSAPUName As String = String.Empty
    Public p_sSAPUPass As String = String.Empty

    Public p_SyncDateTime As String

    'Global DataTable

    Public p_oGLMapping As DataTable
    Public p_oOACT As DataTable

#End Region

#Region "Main Method"

    Sub Main()
        Dim strFunctName As String = String.Empty
        Dim strErrDesc As String = String.Empty

        Try
            strFunctName = "Main Method"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Main Method Starting", strFunctName)
            Console.WriteLine("Starting Integration Module")
            Console.WriteLine("Reading Excel Values")
            Start()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with success", strFunctName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", strFunctName)
            End
        End Try

    End Sub

#End Region

End Module
