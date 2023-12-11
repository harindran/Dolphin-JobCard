Public Class clsJobCard
    Public Const FormType As String = "JobCard"
    Private objForm As SAPbouiCOM.Form
    Private objMatrix As SAPbouiCOM.Matrix
    Private objCombo As SAPbouiCOM.ComboBox
    Private objRecordSet As SAPbobsCOM.Recordset
    Private StrSQL As String
    Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
    Dim objDT As SAPbouiCOM.DataTable
    Public Sub LoadScreen(Optional ByVal Source As String = "")
        objForm = objAddOn.objUIXml.LoadScreenXML("JobCard.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, FormType)
        setReport(objForm.UniqueID)
        objForm.PaneLevel = 1
        objForm.Visible = True
        LoadSeries(objForm.UniqueID)
        If Source = "" Then SetCurrentDate(objForm.UniqueID)
        LoadSalesEmp(objForm.UniqueID)

    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Not Validate(FormUID) Then
                            BubbleEvent = False
                        End If
                        RemoveEmptyRows(FormUID)
                    End If
                    '  If pVal.ItemUID = "1" And pVal.ActionSuccess Then objAddOn.objApplication.Menus.Item("1289").Activate()
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "20" Then
                        objForm.PaneLevel = 1
                    ElseIf pVal.ItemUID = "21" Then
                        objForm.PaneLevel = 2
                    ElseIf pVal.ItemUID = "22" Then
                        objForm.PaneLevel = 3
                    
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    objCFLEvent = pVal
                    objDT = objCFLEvent.SelectedObjects
                    If pVal.ItemUID = "23" And pVal.ColUID = "1" Then ' Item master
                        ChooseItems(FormUID, objCFLEvent)
                    ElseIf pVal.ItemUID = "4" Then
                        CustomerCFL(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "15" Then
                        LoadDocNo(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "23" And pVal.ColUID = "3" Then 'And pVal.CharPressed = 9 Then
                        addItems(FormUID, pVal.Row, pVal.ItemUID)
                    End If
                    If pVal.ItemUID = "30" And pVal.ColUID = "3" Then ' And pVal.CharPressed = 9 Then
                        addItems(FormUID, pVal.Row, pVal.ItemUID)
                    End If

            End Select
        End If
    End Sub
    Private Sub RemoveEmptyRows(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("23").Specific
        For i As Integer = objMatrix.RowCount To 1 Step -1
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                Exit For
            Else
                objMatrix.DeleteRow(i)
            End If
        Next
    End Sub

    Public Sub LoadSeries(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("15").Specific
        If objCombo.ValidValues.Count = 0 Then
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            StrSQL = "select Series, SeriesName from NNM1 where objectcode='MI_JOBCRD'"
            objRecordSet.DoQuery(StrSQL)
            While Not objRecordSet.EoF
                objCombo.ValidValues.Add(objRecordSet.Fields.Item("Series").Value, objRecordSet.Fields.Item("SeriesName").Value)
                objRecordSet.MoveNext()
            End While
            objRecordSet = Nothing
        End If
        '  LoadDocNo(FormUID)
        objMatrix = objForm.Items.Item("23").Specific
        If objMatrix.RowCount = 0 Then
            objMatrix.AddRow()
        End If
        objMatrix = objForm.Items.Item("30").Specific
        If objMatrix.RowCount = 0 Then
            objMatrix.AddRow()
        End If
       
    End Sub
    Private Sub SetCurrentDate(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If objForm.Items.Item("14").Specific.string = "" Then
            objForm.Items.Item("14").Specific.string = objAddOn.objApplication.Company.ServerDate
            'objForm.Items.Item("15").Specific.click(0, 0)

        End If
    End Sub
    Public Sub LoadSalesEmp(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("25").Specific
        If objCombo.ValidValues.Count = 0 Then
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            StrSQL = "select slpcode, slpname from OSLP"
            objRecordSet.DoQuery(StrSQL)
            While Not objRecordSet.EoF
                objCombo.ValidValues.Add(objRecordSet.Fields.Item("slpcode").Value, objRecordSet.Fields.Item("slpname").Value)
                objRecordSet.MoveNext()
            End While
            objRecordSet = Nothing
        End If
    End Sub
    Private Sub ChooseItems(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ChooseFromListEvent)

        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("23").Specific
        objMatrix.GetLineData(pVal.Row)
        objForm.DataSources.DBDataSources.Item("@MI_JOBCRD1").SetValue("U_itemCode", 0, objDT.GetValue("ItemCode", 0))
        objForm.DataSources.DBDataSources.Item("@MI_JOBCRD1").SetValue("U_itemname", 0, objDT.GetValue("ItemName", 0))
        objMatrix.SetLineData(pVal.Row)
    End Sub

    Public Sub LoadDocNo(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objCombo = objForm.Items.Item("15").Specific
            If objCombo.Selected.Description.ToUpper = "MANUAL" Then
                objForm.Items.Item("12").Enabled = True
                objForm.Items.Item("12").Specific.value = " "
            Else
                objForm.Items.Item("12").Enabled = False
                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                StrSQL = ""
                StrSQL = "select nextnumber from nnm1 where series=" & objCombo.Selected.Value & " and objectcode='MI_JOBCRD'"
                objRecordSet.DoQuery(StrSQL)
                objForm.Items.Item("12").Specific.value = objRecordSet.Fields.Item("nextnumber").Value
            End If
        Catch ex As Exception
            '            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub CustomerCFL(ByVal FormUID As String)
        Try
            If objDT Is Nothing Then
            Else
                objForm.Items.Item("4").Specific.value = objDT.GetValue("CardCode", 0).ToString
            End If
        Catch ex As Exception
        End Try
        objForm.Items.Item("6").Specific.value = objDT.GetValue("CardName", 0).ToString

        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        StrSQL = "select cntctcode, name  from ocpr where CardCode ='" & objForm.Items.Item("4").Specific.value & "'"
        objRecordSet.DoQuery(StrSQL)
        While Not objRecordSet.EoF
            objCombo = objForm.Items.Item("8").Specific
            objCombo.ValidValues.Add(objRecordSet.Fields.Item("cntctcode").Value, objRecordSet.Fields.Item("name").Value)
            objRecordSet.MoveNext()
        End While
        objRecordSet = Nothing
    End Sub
    Private Function Validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("23").Specific
        If objForm.Items.Item("34").Specific.ToString = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Please enter Address")
            Return False
        ElseIf objForm.Items.Item("36").Specific.ToString = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Please enter Reference Name")
            Return False
        ElseIf objForm.Items.Item("14").Specific.ToString = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Please enter Document Date")
            Return False
        ElseIf objForm.Items.Item("4").Specific.ToString = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Customer Code is Mandatory")
            Return False
        ElseIf objForm.Items.Item("12").Specific.ToString = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Document Number is Mandatory")
            Return False
        ElseIf objMatrix.Columns.Item("3").Cells.Item(1).Specific.String = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Atleast One Item has to be recorded to Save Job Card")
            Return False
        End If

        Return True
    End Function
    Private Sub addItems(ByVal FormUID As String, ByVal RowID As Integer, ByVal MatrixID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item(MatrixID).Specific
        Select Case MatrixID
            Case "23"
                If objMatrix.Columns.Item("3").Cells.Item(RowID).Specific.string <> "" Then
                    If RowID = objMatrix.RowCount Then
                        objForm.DataSources.DBDataSources.Item("@MI_JOBCRD1").Clear()
                        objMatrix.AddRow()
                    End If
                End If
            Case "30"
                If objMatrix.Columns.Item("3").Cells.Item(RowID).Specific.string <> "" Then
                    If RowID = objMatrix.RowCount Then
                        objForm.DataSources.DBDataSources.Item("@MI_JOBCRD2").Clear()
                        objMatrix.AddRow()
                    End If
                End If

        End Select
            'objMatrix.Columns.Item("0").Cells.Item(objMatrix.RowCount).Specific.string = CStr(objMatrix.RowCount)
            ' objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            ' objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    End Sub
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.ActiveForm

        If pVal.MenuUID = "ditem" Then
            objMatrix = objForm.Items.Item("23").Specific
            objMatrix.DeleteRow(objAddOn.ZB_row)
        End If
    End Sub
    Public Sub LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.Item(eventInfo.FormUID)
        eventInfo.LayoutKey = objForm.Items.Item("12").Specific.string
        'eventInfo.LayoutKey = objForm.Items.Item("12A").Specific.string '--- for PLD changed on 02-Jan-2018
    End Sub

    Private Sub setReport(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim rptTypeService As SAPbobsCOM.ReportTypesService
        Dim newType As SAPbobsCOM.ReportType
        Dim newtypesParam As SAPbobsCOM.ReportTypesParams
        rptTypeService = objAddOn.objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        newtypesParam = rptTypeService.GetReportTypeList

        Dim i As Integer
        For i = 0 To newtypesParam.Count - 1

            If newtypesParam.Item(i).TypeName = FormType And newtypesParam.Item(i).MenuID = FormType Then
                objForm.ReportType = newtypesParam.Item(i).TypeCode

                Exit For
            End If
        Next i

    End Sub
End Class
