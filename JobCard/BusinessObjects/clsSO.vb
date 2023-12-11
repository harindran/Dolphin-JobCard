Public Class clsSO
    Private objForm As SAPbouiCOM.Form
    Private objMatrix As SAPbouiCOM.Matrix
    Private objCombo As SAPbouiCOM.ComboBox
    Private objRecordSet As SAPbobsCOM.Recordset
    Private StrSQL As String
    Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
    Dim objDT As SAPbouiCOM.DataTable
    Dim objButtonCombo As SAPbouiCOM.ButtonCombo
    Public Sub LoadScreen(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objForm.Freeze(True)
        CreateButtons(FormUID)
        objForm.Freeze(False)
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "CFJC" Then
                        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objAddOn.objApplication.SetStatusBarMessage("Document should be in Add Mode", SAPbouiCOM.BoMessageTime.bmt_Short)
                            Exit Sub
                        End If

                        CopyFromJobCard(FormUID)
                    ElseIf pVal.ItemUID = "CTJC" Then
                        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objAddOn.objApplication.SetStatusBarMessage("Document should be in OK Mode", SAPbouiCOM.BoMessageTime.bmt_Short)
                            Exit Sub
                        End If
                        CopyToJobCard(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objAddOn.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "JC" Then

                        If objButtonCombo.Selected.Value = "CF" Then
                            objButtonCombo.Caption = "Job Card"
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            CopyFromJobCard(FormUID)
                        ElseIf objButtonCombo.Selected.Value = "CT" Then
                            objButtonCombo.Caption = "Job Card"
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            CopyToJobCard(FormUID)
                        End If
                    End If

            End Select
        End If
    End Sub
    Private Sub CreateButtons(ByVal FormUID As String)
        Dim objButton As SAPbouiCOM.Button
        Dim objItem As SAPbouiCOM.Item
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try
            objButton = objForm.Items.Item("CFJC").Specific
        Catch ex As Exception
            objItem = objForm.Items.Add("CFJC", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
            objItem.Top = objForm.Items.Item("2").Top
            objItem.Width = objForm.Items.Item("10000330").Width
            objItem.Height = objForm.Items.Item("2").Height
            objButton = objItem.Specific
            objItem.Visible = True
            objButton.Caption = "Copy From JC"
        End Try
        Try
            objButton = objForm.Items.Item("CTJC").Specific
        Catch ex As Exception
            objItem = objForm.Items.Add("CTJC", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objForm.Items.Item("CFJC").Left + objForm.Items.Item("CFJC").Width + 5
            objItem.Top = objForm.Items.Item("2").Top
            objItem.Width = objForm.Items.Item("10000330").Width
            objItem.Height = objForm.Items.Item("2").Height
            objButton = objItem.Specific
            objItem.Visible = True
            objButton.Caption = "Copy To JC"
        End Try

        Try
            objbuttoncombo = objForm.Items.Item("JC").Specific

        Catch ex As Exception
            objItem = objForm.Items.Add("JC", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
            objItem.Width = objForm.Items.Item("10000330").Width
            objItem.Height = objForm.Items.Item("2").Height
            objItem.Top = objForm.Items.Item("2").Top
            objItem.Left = objForm.Items.Item("10000330").Left - (objForm.Items.Item("10000330").Width + 5)
            objItem.DisplayDesc = True
            objButtonCombo = objItem.Specific
            objItem.Visible = False
            objButtonCombo.Caption = "Job Card"
            objButtonCombo.ValidValues.Add("CF", "From JC")
            objButtonCombo.ValidValues.Add("CT", "To JC")
            objButtonCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


        End Try
        
       
    End Sub
    Private Function SOAlreadyExists(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        StrSQL = "Select docnum from ordr where U_jcentry ='" & Trim(objForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_jcentry", 0)) & "'"
        objRecordSet.DoQuery(StrSQL)
        If objRecordSet.RecordCount > 0 Then
            objAddOn.objApplication.SetStatusBarMessage("SO already created; Please refer SO Number : " & CStr(objRecordSet.Fields.Item("docnum").Value), SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        End If

        Return False

    End Function
    Private Function JCAlreadyExists(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        StrSQL = "Select docnum from [@MI_JOBCRD] where U_soentry ='" & Trim(objForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0)) & "'"
        objRecordSet.DoQuery(StrSQL)
        If objRecordSet.RecordCount > 0 Then
            objAddOn.objApplication.SetStatusBarMessage("SO already created; Please refer JC Number : " & CStr(objRecordSet.Fields.Item("docnum").Value), SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        End If

        Return False
    End Function
    Private Sub CopyFromJobCard(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

        If Not SOAlreadyExists(FormUID) Then
            StrSQL = "  select LineId,U_cardcode, U_docdate,isnull(U_contact,'') U_contact,U_custdesc ,U_poref, isnull(U_cashcust,'') U_cashcust ,isnull(U_addr,'') U_addr, U_itemcode ,isnull(U_uom,'') U_uom,isnull(U_whscode,'') U_whscode ,U_qty , U_cldate ,isnull(U_saleemp,'-1') U_saleemp ,U_owner,U_remarks   from [@MI_JOBCRD] T0 join [@MI_JOBCRD1] T1 on T1.DocEntry = T0.DocEntry   where T0.docentry='" & Trim(objForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_jcentry", 0)) & "'"
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim objUDFForm As SAPbouiCOM.Form
            objUDFForm = objAddOn.objApplication.Forms.Item(objForm.UDFFormUID)
            objRecordSet.DoQuery(StrSQL)
            If Not objRecordSet.EoF Then
                ' objForm.Freeze(True)
                objForm.Items.Item("4").Specific.String = objRecordSet.Fields.Item("U_cardcode").Value

                objForm.Items.Item("10").Specific.String = objAddOn.objGenFunc.GetSBODateString(objRecordSet.Fields.Item("U_docdate").Value)
                If objRecordSet.Fields.Item("U_contact").Value <> "" Then
                    objCombo = objForm.Items.Item("85").Specific
                    objCombo.Select(objRecordSet.Fields.Item("U_contact").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                ' objForm.Items.Item("85").Specific.String = objRecordSet.Fields.Item("U_contact").Value
                objForm.Items.Item("14").Specific.String = objRecordSet.Fields.Item("U_poref").Value
               

                objUDFForm.Items.Item("U_cashcust").Specific.String = objRecordSet.Fields.Item("U_cashcust").Value
                objUDFForm.Items.Item("U_addr").Specific.String = objRecordSet.Fields.Item("U_addr").Value

                objForm.Items.Item("12").Specific.String = objAddOn.objGenFunc.GetSBODateString(objRecordSet.Fields.Item("U_cldate").Value)
                If objRecordSet.Fields.Item("U_saleemp").Value <> "" Then
                    objCombo = objForm.Items.Item("20").Specific
                    objCombo.Select(objRecordSet.Fields.Item("U_saleemp").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                objForm.Items.Item("222").Specific.String = objRecordSet.Fields.Item("U_owner").Value
                objForm.Items.Item("16").Specific.String = objRecordSet.Fields.Item("U_remarks").Value
             

                objMatrix = objForm.Items.Item("38").Specific
                Dim count As Integer
                count = 0

                While Not objRecordSet.EoF

                    objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Specific.String = objRecordSet.Fields.Item("U_itemcode").Value
                    objMatrix.Columns.Item("U_custdesc").Cells.Item(objMatrix.RowCount - 1).Specific.String = objRecordSet.Fields.Item("U_custdesc").Value
                    objMatrix.Columns.Item("U_JcLineNum").Cells.Item(objMatrix.RowCount - 1).Specific.String = objRecordSet.Fields.Item("LineId").Value
                    objMatrix.Columns.Item("11").Cells.Item(objMatrix.RowCount - 1).Specific.String = objRecordSet.Fields.Item("U_qty").Value
                    objMatrix.Columns.Item("24").Cells.Item(objMatrix.RowCount - 1).Specific.String = objRecordSet.Fields.Item("U_whscode").Value
                    objMatrix.Columns.Item("25").Cells.Item(objMatrix.RowCount - 1).Specific.String = objAddOn.objGenFunc.GetSBODateString(objRecordSet.Fields.Item("U_cldate").Value)
                    'objMatrix.Columns.Item("212").Cells.Item(objMatrix.RowCount - 1).Specific.String = objRecordSet.Fields.Item("U_uom").Value

                    count = count + 1

                    objRecordSet.MoveNext()
                End While

                ' objForm.Freeze(False)

                objAddOn.objApplication.MessageBox("Please fill up the price")

            End If
        End If
    End Sub

    Private Sub CopyToJobCard(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objJCForm As SAPbouiCOM.Form
        Dim objSORS As SAPbobsCOM.Recordset
        Dim SOEntry As Long
        If JCAlreadyExists(FormUID) Then Exit Sub

        Try

            SOEntry = objForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0)
            'If Not SOAlreadyExists(FormUID) Then

            StrSQL = "select CardCode, T0.DocDate,isnull(cntctcode,0) CntctCode,isnull(NumAtCard,'') NumAtCard , isnull(U_cashcust,'') cashcust ,isnull(U_addr,'') addr, ItemCode,unitMsr,T1.U_custdesc ,Quantity,WhsCode ,ShipDate, DocDueDate ,isnull(T0.slpcode,'-1') SlpCode , T2.LastName +', '+T2.FirstName as OwnerCode, Comments   from [ORDR] T0 join [RDR1] T1 on T1.DocEntry = T0.DocEntry  left outer join OHEM T2 on T2.empid=T0.ownercode where T0.docentry='" & SOEntry & "'"
            objSORS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim objUDFForm As SAPbouiCOM.Form
            objUDFForm = objAddOn.objApplication.Forms.Item(objForm.UDFFormUID)
            objSORS.DoQuery(StrSQL)
            objAddOn.objJobCard.LoadScreen("FromSO")
            objJCForm = objAddOn.objApplication.Forms.GetForm("JobCard", 1)
            If Not objSORS.EoF Then
                objJCForm.Items.Item("4").Specific.value = objSORS.Fields.Item("CardCode").Value
                objJCForm.Items.Item("14").Specific.String = objAddOn.objGenFunc.GetSBODateString(objSORS.Fields.Item("DocDate").Value)
                objJCForm.Items.Item("10").Specific.String = objSORS.Fields.Item("NumAtCard").Value
                objJCForm.Items.Item("36").Specific.String = objSORS.Fields.Item("cashcust").Value
                objJCForm.Items.Item("34").Specific.String = objSORS.Fields.Item("addr").Value
                objJCForm.Items.Item("27").Specific.String = objSORS.Fields.Item("OwnerCode").Value
                objJCForm.Items.Item("29").Specific.String = objSORS.Fields.Item("Comments").Value
                objJCForm.Items.Item("39").Specific.value = SOEntry
                'Dim cntctcode As Integer = objSORS.Fields.Item("CntctCode").Value

                'Dim slpcode As Integer = objSORS.Fields.Item("SlpCode").Value

                objMatrix = objJCForm.Items.Item("23").Specific
                Dim count As Integer
                count = 0
                Try
                    objCombo = objJCForm.Items.Item("25").Specific

                    'For i As Integer = 0 To objCombo.ValidValues.Count - 1
                    '    If objCombo.ValidValues.Item(i).Value = objSORS.Fields.Item("SlpCode").Value Then
                    '        objCombo.SelectExclusive(i, SAPbouiCOM.BoSearchKey.psk_Index)
                    '        Exit For
                    '    End If
                    'Next
                    objCombo.SelectExclusive(CStr(objSORS.Fields.Item("SlpCode").Value), SAPbouiCOM.BoSearchKey.psk_ByValue)


                    objCombo = objJCForm.Items.Item("8").Specific
                    If objSORS.Fields.Item("CntctCode").Value <> 0 Then
                        objCombo.SelectExclusive(CStr(objSORS.Fields.Item("CntctCode").Value), SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If



                Catch ex As Exception

                End Try
                While Not objSORS.EoF
                    objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Specific.String = objSORS.Fields.Item("ItemCode").Value
                    objMatrix.Columns.Item("3").Cells.Item(objMatrix.RowCount).Specific.String = objSORS.Fields.Item("U_custdesc").Value
                    objMatrix.Columns.Item("4").Cells.Item(objMatrix.RowCount).Specific.String = objSORS.Fields.Item("Quantity").Value
                    objMatrix.Columns.Item("5").Cells.Item(objMatrix.RowCount - 1).Specific.String = objSORS.Fields.Item("unitMsr").Value
                    objMatrix.Columns.Item("5A").Cells.Item(objMatrix.RowCount - 1).Specific.String = objSORS.Fields.Item("WhsCode").Value
                    objMatrix.Columns.Item("6").Cells.Item(objMatrix.RowCount - 1).Specific.String = objAddOn.objGenFunc.GetSBODateString(objSORS.Fields.Item("ShipDate").Value)
                    'objJCForm.DataSources.DBDataSources.Item("@MI_JOBCRD1").SetValue("U_itemcode", count, objSORS.Fields.Item("ItemCode").Value)
                    'objJCForm.DataSources.DBDataSources.Item("@MI_JOBCRD1").SetValue("U_custdesc", count, objSORS.Fields.Item("U_desc").Value)
                    'objJCForm.DataSources.DBDataSources.Item("@MI_JOBCRD1").SetValue("U_qty", count, objSORS.Fields.Item("Quantity").Value)
                    ' objMatrix.AddRow()
                    '  objJCForm.DataSources.DBDataSources.Item("MI_JOBCRD1").Clear()
                    count = count + 1

                    objSORS.MoveNext()
                End While


            End If
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message)
        End Try


    End Sub
End Class
