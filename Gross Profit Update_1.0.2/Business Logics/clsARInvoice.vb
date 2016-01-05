Public Class clsARInvoice
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private oDT_GrossProfit As DataTable
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub populateGrossprofitBase(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim frmSource As SAPbouiCOM.Form
            Dim oSourcematrix As SAPbouiCOM.Matrix
            Dim strBaseType, strBaseEntry, strBaseLine As String
            Dim intSOEntry, intSOLine, intPOEntry, intPOLine, intPIEntry, intPILine As Integer
            Dim dblsalesprice, dblPurchaseprice As Double
            frmSource = aform
            oSourcematrix = frmSource.Items.Item("38").Specific

            Dim oRs As SAPbobsCOM.Recordset
            oRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oDT_GrossProfit = New DataTable()
            oDT_GrossProfit.Columns.Add("ARLineNum")
            oDT_GrossProfit.Columns.Add("ARItemCode")
            oDT_GrossProfit.Columns.Add("SaleBaseType")
            oDT_GrossProfit.Columns.Add("SaleBaseEntry")
            oDT_GrossProfit.Columns.Add("SaleBaseLine")
            oDT_GrossProfit.Columns.Add("GPBasePrice")

            For intRow As Integer = 1 To oSourcematrix.VisualRowCount
                If oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString() = "" Then
                Else
                    If oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString() = "17" Then
                        Dim dr As DataRow
                        Dim oStr As String = "Select isnull(T0.DropShip,'N') from  OWHS T0  Where T0.WhsCode='" & oSourcematrix.Columns.Item("24").Cells.Item(intRow).Specific.Value.ToString() & "'"
                        Dim oTest As SAPbobsCOM.Recordset
                        Dim intBaseEntry, intBaseRow, intBaseType As Double
                        Dim strItemCode As String
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTest.DoQuery(oStr)
                        If 1 = 1 Then 'oTest.Fields.Item(0).Value = "Y" Then
                            intBaseEntry = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oSourcematrix, "45", intRow))
                            intBaseRow = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oSourcematrix, "46", intRow))
                            intBaseType = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oSourcematrix, "43", intRow))
                            strItemCode = oApplication.Utilities.getMatrixValues(oSourcematrix, "1", intRow)

                            Dim strQry As String = "Select * from POR1 T1 join OPOR T0 on T1.DocEntry=T0.docEntry Where T1.TargetType=18 and T1.BaseType='17' and T1.BaseEntry='" & oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString() & "' and T1.BaseLine='" & oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString() & "'"

                            strQry = "Select  avg(T1.U_Cost)FROM SRI1 T0 inner join  OSRN T1 on T1.SysNumber=T0.SysSerial "
                            strQry = strQry & "  where T0.BaseEntry = " & CInt(intBaseEntry) & " And T0.BaseType = " & CInt(intBaseType) & " And T0.BaseLinNum = " & CInt(intBaseRow)
                            oRs.DoQuery(strQry)
                            If oRs.RecordCount > 0 Then
                                strQry = "Select  avg(T1.U_Cost)'GPPrice' FROM SRI1 T0 inner join  OSRN T1 on T1.SysNumber=T0.SysSerial "
                                strQry = strQry & "  where  T0.BaseEntry = " & CInt(intBaseEntry) & " And T0.BaseType = " & CInt(intBaseType) & " And T0.BaseLinNum = " & CInt(intBaseRow)

                                strQry = " SELECT T1.MdAbsEntry   FROM [dbo].[OITL]  T0 INNER JOIN [dbo].[ITL1]  T1 ON T0.LogEntry = T1.LogEntry where T0.ItemCode='" & strItemCode & "' and  T0.DocEntry=" & CInt(intBaseEntry) & " and T0.DocLine=" & CInt(intBaseRow) & " and T0.DocType=" & CInt(intBaseType)
                                strQry = "Select Avg(T1.U_Cost) 'GPPrice' from OSRN T1 where T1.AbsEntry in (" & strQry & ")"


                                oRs.DoQuery(strQry)
                                If oRs.RecordCount > 0 Then
                                    Dim BasePrice As Double = oRs.Fields.Item("GPPrice").Value
                                    dr = oDT_GrossProfit.NewRow()
                                    dr("ARLineNum") = intRow
                                    dr("ARItemCode") = oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseType") = oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseEntry") = oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseLine") = oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("GPBasePrice") = BasePrice
                                    oDT_GrossProfit.Rows.Add(dr)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If Not IsNothing(oDT_GrossProfit) And oDT_GrossProfit.Rows.Count > 0 Then
                Try
                    oApplication.SBO_Application.ActivateMenuItem(mnu_Grossprofit)
                Catch ex As Exception
                End Try

                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_Grossprofit Then
                    Dim oMatrix_GP = oForm.Items.Item("3").Specific
                    Dim dr4 As DataRow
                    For intRow As Integer = 1 To oDT_GrossProfit.Rows.Count
                        dr4 = oDT_GrossProfit.Rows(intRow - 1)
                        oMatrix_GP.Columns.Item("3").Cells.Item(Convert.ToInt32(dr4("ARLineNum").ToString())).Specific.Value = dr4("GPBasePrice").ToString()
                    Next
                End If
                Try

                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Catch ex As Exception
                    aform.Freeze(False)
                End Try
                aform.Freeze(False)
            End If
            aform.Freeze(False)
            
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub populateGrossprofitBase_menuClick(ByVal aform As SAPbouiCOM.Form, ByVal aSource As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim frmSource As SAPbouiCOM.Form
            Dim oSourcematrix As SAPbouiCOM.Matrix
            Dim strBaseType, strBaseEntry, strBaseLine As String
            Dim intSOEntry, intSOLine, intPOEntry, intPOLine, intPIEntry, intPILine As Integer
            Dim dblsalesprice, dblPurchaseprice As Double
            frmSource = aform
            oSourcematrix = frmSource.Items.Item("38").Specific

            Dim oRs As SAPbobsCOM.Recordset
            oRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oDT_GrossProfit = New DataTable()
            oDT_GrossProfit.Columns.Add("ARLineNum")
            oDT_GrossProfit.Columns.Add("ARItemCode")
            oDT_GrossProfit.Columns.Add("SaleBaseType")
            oDT_GrossProfit.Columns.Add("SaleBaseEntry")
            oDT_GrossProfit.Columns.Add("SaleBaseLine")
            oDT_GrossProfit.Columns.Add("GPBasePrice")

            For intRow As Integer = 1 To oSourcematrix.VisualRowCount
                If oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString() = "" Then
                Else
                    If oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString() = "17" Then
                        Dim dr As DataRow
                        Dim oStr As String = "Select isnull(T0.DropShip,'N') from  OWHS T0  Where T0.WhsCode='" & oSourcematrix.Columns.Item("24").Cells.Item(intRow).Specific.Value.ToString() & "'"
                        Dim oTest As SAPbobsCOM.Recordset
                        Dim intBaseEntry, intBaseRow, intBaseType As Double
                        Dim stritemcode As String
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTest.DoQuery(oStr)
                        If 1 = 1 Then 'oTest.Fields.Item(0).Value = "Y" Then
                            intBaseEntry = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oSourcematrix, "45", intRow))
                            intBaseRow = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oSourcematrix, "46", intRow))
                            intBaseType = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oSourcematrix, "43", intRow))
                            stritemcode = oApplication.Utilities.getMatrixValues(oSourcematrix, "1", intRow)

                            Dim strQry As String = "Select * from POR1 T1 join OPOR T0 on T1.DocEntry=T0.docEntry Where T1.TargetType=18 and T1.BaseType='17' and T1.BaseEntry='" & oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString() & "' and T1.BaseLine='" & oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString() & "'"
                            strQry = "Select  avg(T1.U_Cost)FROM SRI1 T0 inner join  OSRN T1 on T1.SysNumber=T0.SysSerial "
                            strQry = strQry & "  where T0.BaseEntry = " & CInt(intBaseEntry) & " And T0.BaseType = " & CInt(intBaseType) & " And T0.BaseLinNum = " & CInt(intBaseRow)
                            oRs.DoQuery(strQry)
                            If oRs.RecordCount > 0 Then
                                strQry = "Select  avg(T1.U_Cost)'GPPrice' FROM SRI1 T0 inner join  OSRN T1 on T1.SysNumber=T0.SysSerial "
                                strQry = strQry & "  where T0.BaseEntry = " & CInt(intBaseEntry) & " And T0.BaseType = " & CInt(intBaseType) & " And T0.BaseLinNum = " & CInt(intBaseRow)

                                strQry = " SELECT T1.MdAbsEntry   FROM [dbo].[OITL]  T0 INNER JOIN [dbo].[ITL1]  T1 ON T0.LogEntry = T1.LogEntry where T0.ItemCode='" & stritemcode & "' and  T0.DocEntry=" & CInt(intBaseEntry) & " and T0.DocLine=" & CInt(intBaseRow) & " and T0.DocType=" & CInt(intBaseType)
                                strQry = "Select Avg(T1.U_Cost) 'GPPrice' from OSRN T1 where T1.AbsEntry in (" & strQry & ")"


                                oRs.DoQuery(strQry)
                                If oRs.RecordCount > 0 Then
                                    Dim BasePrice As Double = oRs.Fields.Item("GPPrice").Value
                                    dr = oDT_GrossProfit.NewRow()
                                    dr("ARLineNum") = intRow
                                    dr("ARItemCode") = oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseType") = oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseEntry") = oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseLine") = oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("GPBasePrice") = BasePrice
                                    oDT_GrossProfit.Rows.Add(dr)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If Not IsNothing(oDT_GrossProfit) And oDT_GrossProfit.Rows.Count > 0 Then
                'oForm = oApplication.SBO_Application.Forms.ActiveForm()
                oForm = aSource
                If oForm.TypeEx = frm_Grossprofit Then
                    Dim oMatrix_GP = oForm.Items.Item("3").Specific
                    Dim dr4 As DataRow
                    For intRow As Integer = 1 To oDT_GrossProfit.Rows.Count
                        dr4 = oDT_GrossProfit.Rows(intRow - 1)
                        oMatrix_GP.Columns.Item("3").Cells.Item(Convert.ToInt32(dr4("ARLineNum").ToString())).Specific.Value = dr4("GPBasePrice").ToString()
                    Next
                End If
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                aform.Freeze(False)
            End If
            aform.Freeze(False)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Invoice Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    populateGrossprofitBase(oForm)
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oForm.Freeze(False)
                                End If


                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Grossprofit
                    If pVal.BeforeAction = True Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        frmSourceForm = oForm
                    Else
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If frmSourceForm.TypeEx = frm_Invoice Then
                            If 1 = 1 Then 'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                '      populateGrossprofitBase_menuClick(frmSourceForm, oForm)
                            End If
                        End If
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Form Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If BusinessObjectInfo.FormTypeEx = frm_GoodsReceipt Then
                    Dim intBaseType, intBaseEntry, intBaseLine As Integer
                    Dim oDoc As SAPbobsCOM.Documents
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                    If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        For intRow As Integer = 0 To oDoc.Lines.Count - 1
                            oDoc.Lines.SetCurrentLine(intRow)
                            intBaseEntry = oDoc.Lines.DocEntry
                            intBaseLine = oDoc.Lines.LineNum
                            intBaseType = 20
                            Dim strSysnumbers As String
                            strSysnumbers = "SELECT  T0.[SysSerial] FROM SRI1 T0 inner join  OSRN T1 on T1.SysNumber=T0.SysSerial "
                            strSysnumbers = strSysnumbers & " where T0.BaseEntry = " & intBaseEntry & " And T0.BaseType = 20 And T0.BaseLinNum = " & intBaseLine
                            Dim orec As SAPbobsCOM.Recordset
                            orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strSysnumbers = "Update OSRN set U_Cost ='" & oDoc.Lines.Price & "' where SysNumber in (" & strSysnumbers & ")"
                            orec.DoQuery(strSysnumbers)
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

    
End Class
