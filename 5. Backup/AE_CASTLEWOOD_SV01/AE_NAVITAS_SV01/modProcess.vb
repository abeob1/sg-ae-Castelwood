Module modProcess

#Region "Start"
    Public Sub Start()
        Dim sFuncName As String = "Start()"
        Dim sErrDesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("calling UploadFiles()", sFuncName)



            UploadFiles(sErrDesc)

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


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try

    End Sub
#End Region

#Region "Read Excel Files"

    Private Function UploadFiles(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadFiles"
        Dim bIsFileExists As Boolean = False
        Dim oDVData As DataView = New DataView

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Upload funciton", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            p_oDtSuccess = CreateDataTable("FileName", "DocEntry", "Status")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")
            p_oDtReport = CreateDataTable("Type", "DocEntry", "BPCode", "Owner")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IdentifyTXTFile_JournalEntry() ", sFuncName)
            If IdentifyExcelFile_JournalEntry(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


        
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)
            UploadFiles = RTN_SUCCESS

        Catch ex As Exception
            UploadFiles = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uplodiang AR file.", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function
#End Region

End Module
