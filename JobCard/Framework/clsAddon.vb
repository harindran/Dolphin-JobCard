Imports System.IO

Public Class clsAddOn
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0
    Public SOMenuID As String = "0"
   
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim MenuCount As Integer = 0

   
    Public objJobCard As clsJobCard
    Public objSO As clsSO
    Public objSQ As clsSQ
    Public HWKEY() As String = New String() {"L1552968038", "Q0198611247", "T0264302252", "M0090876837", "H0922924113", "K1679825911", "F0123559701"}
    Private Sub CheckLicense()

    End Sub
    Function isValidLicense() As Boolean
        Try
            objApplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = objApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            objApplication.Forms.ActiveForm.Close()

            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next
            MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
            Return False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return True
    End Function
    Public Sub Intialize()
        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try
            createObjects()
            loadMenu()
            createTables()
            createUDOs()
            addJobCardReporttype()
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message)
            End
        End Try
        If isValidLicense() Then
            objApplication.SetStatusBarMessage("Addon connected  successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Else
            objApplication.SetStatusBarMessage("Failed To Connect, Please Check The License Configuration", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objCompany.Disconnect()
            objApplication = Nothing
            objCompany = Nothing
            End
        End If
    End Sub
    Private Sub createUDOs()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        Dim ct1(3) As String
        

        ct1(0) = "MI_JOBCRD1"
        ct1(1) = "MI_JOBCRD2"
        ct1(2) = "MI_JOBCRD3"

       

        objUDFEngine.createUDO("MI_JOBCRD", "MI_JOBCRD", "Job Card", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
      
    End Sub
    Private Sub createObjects()
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)

      
        objJobCard = New clsJobCard
        objSO = New clsSO
        objSQ = New clsSQ
    End Sub
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                Case clsJobCard.FormType
                    objJobCard.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "139" 'sales order
                    objSO.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "149" 'sales Quotation
                    objSQ.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        Try
          
        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Sub
    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application)
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        If pVal.BeforeAction Then
            Select Case pVal.MenuUID
               
            End Select
        Else
            Try
                Select Case pVal.MenuUID
                    Case clsJobCard.FormType
                        objJobCard.LoadScreen()
                    Case "1282", "1290", "1288", "1289", "1291"
                        If objApplication.Forms.ActiveForm.UniqueID.Contains(clsJobCard.FormType) Then
                            objJobCard.LoadSeries(objApplication.Forms.ActiveForm.UniqueID)
                            objJobCard.LoadSalesEmp(objApplication.Forms.ActiveForm.UniqueID)
                        End If

                    Case "2050" ' Sales Order
                        objSO.LoadScreen(objApplication.Forms.ActiveForm.UniqueID)
                    Case "2049" ' Sales Order
                        objSQ.LoadScreen(objApplication.Forms.ActiveForm.UniqueID)
                    Case clsJCAnalysis.FormType

                End Select



            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
    Private Sub loadMenu()
        If objApplication.Menus.Item("43520").SubMenus.Exists("MIPLJC") Then Return
        MenuCount = objApplication.Menus.Item("43520").SubMenus.Count

        CreateMenu(Application.StartupPath + "\jc1.png", MenuCount + 1, "Job Card Management", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLJC", objApplication.Menus.Item("43520"))
        CreateMenu("", 1, "Job Card", SAPbouiCOM.BoMenuType.mt_STRING, clsJobCard.FormType, objApplication.Menus.Item("MIPLJC"))
        CreateMenu("", 2, "JC Analysis", SAPbouiCOM.BoMenuType.mt_STRING, clsJCAnalysis.FormType, objApplication.Menus.Item("MIPLJC"))
        
    End Sub
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function
    Private Sub createTables()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        objAddOn.objApplication.SetStatusBarMessage("Creating Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        ' WriteSMSLog("0")
        '-----SalesOrder--- ---------------------------- 
        objUDFEngine.AddAlphaField("OCRG", "grpname", "Group Name", 100)
        objUDFEngine.AddAlphaField("ORDR", "jcentry", "Job Card Entry", 10)
        objUDFEngine.AddAlphaField("ORDR", "jcno", "Job Card No", 20)
        objUDFEngine.AddAlphaField("ORDR", "cashcust", "Ref Name", 50)
        objUDFEngine.AddAlphaField("ORDR", "addr", "Address", 250)
        objUDFEngine.AddAlphaField("ORDR", "custdesc", "Customer Desc", 100)
        '---------------Job Card ---------------
        objUDFEngine.CreateTable("MI_JOBCRD", "JobCard Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "cardcode", "Card Code", 10)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "cardname", "Card Name", 50)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "contact", "Contact Person", 50)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "poref", "LPO reference", 30)
        objUDFEngine.AddDateField("MI_JOBCRD", "docdate", "Doc date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.addField("MI_JOBCRD", "type", "Job Card Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "SP,SR", "Supply,Service", "SR")
        objUDFEngine.addField("MI_JOBCRD", "status", "Job Card Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 25, SAPbobsCOM.BoFldSubTypes.st_None, "O,P,H,W,R,C,N", "Open,Pending,Hold,WIP,Ready to Dispatch,Closed,Cancelled", "O")
        objUDFEngine.AddAlphaField("MI_JOBCRD", "saleemp", "Sales Employee", 20)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "owner", "Owner", 120)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "remarks", "Remarks", 250)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "addr", "Address", 250)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "cashcust", "Reference Name", 50)
        objUDFEngine.AddNumericField("MI_JOBCRD", "soentry", "SO Entry", 10)
        objUDFEngine.AddAlphaField("MI_JOBCRD", "sonum", "SO Entry", 10)


        '--------------------JobCard Content ----------------------------------------

        objUDFEngine.CreateTable("MI_JOBCRD1", "JobCard Content", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("MI_JOBCRD1", "itemcode", "Item Code", 25)
        objUDFEngine.AddAlphaField("MI_JOBCRD1", "itemname", "Item Name", 100)
        objUDFEngine.AddAlphaField("MI_JOBCRD1", "custdesc", "Customer description", 100)
        objUDFEngine.AddNumericField("MI_JOBCRD1", "qty", "Quantity", 10)
        objUDFEngine.AddAlphaField("MI_JOBCRD1", "uom", "UOM", 10)
        objUDFEngine.AddDateField("MI_JOBCRD1", "cldate", "Closing Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("MI_JOBCRD1", "whscode", "Warehouse", 20)

        '-----------------JobCard Stages --------------------
        objUDFEngine.CreateTable("MI_JOBCRD2", "JobCard Stages", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddDateField("MI_JOBCRD2", "stgdate", "Stage Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("MI_JOBCRD2", "stage", "Stage", 25)
        objUDFEngine.AddAlphaField("MI_JOBCRD2", "updateby", "Updated By", 25)
        objUDFEngine.AddAlphaField("MI_JOBCRD2", "remarks", "Remarks", 200)
        '--------------------JobCard Attachments -----------------------------------
        objUDFEngine.CreateTable("MI_JOBCRD3", "JobCard Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("MI_JOBCRD3", "trgtpath", "Target Path", 100)
        objUDFEngine.AddAlphaField("MI_JOBCRD3", "filename", "File Name", 100)
        objUDFEngine.AddDateField("MI_JOBCRD3", "attdate", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("MI_JOBCRD3", "remarks", "Remarks", 100)
      


        '*******************  Table ******************* START********************************* END
    End Sub
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        If eventInfo.BeforeAction Then
        Else
            If eventInfo.FormUID.Contains("MI_JOBCRD") And (eventInfo.ItemUID = "23") And eventInfo.Row > 0 Then

                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try

                    If objAddOn.objApplication.Menus.Exists("ditem") Then
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    End If
                Catch ex As Exception

                End Try
                Try

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280").SubMenus.Item("ditem")
                    ZB_row = eventInfo.Row
                Catch ex As Exception
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "ditem"
                    oCreationPackage.String = "Delete Row"
                    oCreationPackage.Enabled = True

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)
                    ZB_row = eventInfo.Row
                End Try
                If eventInfo.ItemUID <> "45" Then
                    '   Dim oMenuItem As SAPbouiCOM.MenuItem
                    '  Dim oMenus As SAPbouiCOM.Menus
                    Try
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    Catch ex As Exception
                        ' MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
            End If
    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                ' objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub

    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)




        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)


    End Sub
    Public Sub WriteSMSLog(ByVal Str As String)
        Dim fs As FileStream
        Dim chatlog As String = Application.StartupPath & "\Log_" & Today.ToString("yyyyMMdd") & ".txt"
        If File.Exists(chatlog) Then
        Else
            fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
            fs.Close()
        End If
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub
    Private Sub addJobCardReporttype()
        Dim rptTypeService As SAPbobsCOM.ReportTypesService
        Dim newType As SAPbobsCOM.ReportType
        Dim newtypeParam As SAPbobsCOM.ReportTypeParams
        Dim newReportParam As SAPbobsCOM.ReportLayoutParams
        Dim ReportExists As Boolean = False
        Try


            Dim newtypesParam As SAPbobsCOM.ReportTypesParams
            rptTypeService = objAddOn.objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
            newtypesParam = rptTypeService.GetReportTypeList

            Dim i As Integer
            For i = 0 To newtypesParam.Count - 1
                If newtypesParam.Item(i).TypeName = clsJobCard.FormType And newtypesParam.Item(i).MenuID = clsJobCard.FormType Then
                    ReportExists = True
                    Exit For
                End If
            Next i

            If Not ReportExists Then
                rptTypeService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
                newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)


                newType.TypeName = clsJobCard.FormType
                newType.AddonName = "JC2Addon"
                newType.AddonFormType = clsJobCard.FormType
                newType.MenuID = clsJobCard.FormType
                newtypeParam = rptTypeService.AddReportType(newType)

                Dim rptService As SAPbobsCOM.ReportLayoutsService
                Dim newReport As SAPbobsCOM.ReportLayout
                rptService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
                newReport = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
                newReport.Author = objCompany.UserName
                newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
                newReport.Name = clsJobCard.FormType
                newReport.TypeCode = newtypeParam.TypeCode

                newReportParam = rptService.AddReportLayout(newReport)

                newType = rptTypeService.GetReportType(newtypeParam)
                newType.DefaultReportLayout = newReportParam.LayoutCode
                rptTypeService.UpdateReportType(newType)

                Dim oBlobParams As SAPbobsCOM.BlobParams
                oBlobParams = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
                oBlobParams.Table = "RDOC"
                oBlobParams.Field = "Template"
                Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
                oKeySegment = oBlobParams.BlobTableKeySegments.Add
                oKeySegment.Name = "DocCode"
                oKeySegment.Value = newReportParam.LayoutCode

                Dim oFile As FileStream
                oFile = New FileStream(Application.StartupPath + "\JobCard.rpt", FileMode.Open)
                Dim fileSize As Integer
                fileSize = oFile.Length
                Dim buf(fileSize) As Byte
                oFile.Read(buf, 0, fileSize)
                oFile.Dispose()

                Dim oBlob As SAPbobsCOM.Blob
                oBlob = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob)
                oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
                objCompany.GetCompanyService.SetBlob(oBlobParams, oBlob)
            End If
        Catch ex As Exception
            objApplication.MessageBox(ex.Message)
        End Try

    End Sub

    Private Sub objApplication_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objApplication.LayoutKeyEvent

        'BubbleEvent = True
        If eventInfo.BeforeAction = True Then
            If eventInfo.FormUID.Contains(clsJobCard.FormType) Then
                objJobCard.LayoutKeyEvent(eventInfo, BubbleEvent)
            End If
        End If
    End Sub
End Class


