Public Class clsSQ
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
        


    End Sub
    Private Function SQAlreadyExists(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        StrSQL = "Select docnum from OQUT where U_jcentry ='" & Trim(objForm.DataSources.DBDataSources.Item("OQUT").GetValue("U_jcentry", 0)) & "'"
        objRecordSet.DoQuery(StrSQL)
        If objRecordSet.RecordCount > 0 Then
            objAddOn.objApplication.SetStatusBarMessage("SQ already created; Please refer SQ Number : " & CStr(objRecordSet.Fields.Item("docnum").Value), SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        End If

        Return False

    End Function
    Private Sub CopyFromJobCard(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

        If Not SQAlreadyExists(FormUID) Then
            StrSQL = "  select U_cardcode, U_docdate,isnull(U_contact,'') U_contact,U_custdesc ,U_poref, isnull(U_cashcust,'') U_cashcust ,isnull(U_addr,'') U_addr, U_itemcode ,isnull(U_uom,'') U_uom,isnull(U_whscode,'') U_whscode ,U_qty , U_cldate ,isnull(U_saleemp,'-1') U_saleemp ,U_owner,U_remarks   from [@MI_JOBCRD] T0 join [@MI_JOBCRD1] T1 on T1.DocEntry = T0.DocEntry   where T0.docentry='" & Trim(objForm.DataSources.DBDataSources.Item("OQUT").GetValue("U_jcentry", 0)) & "'"
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

   
End Class
