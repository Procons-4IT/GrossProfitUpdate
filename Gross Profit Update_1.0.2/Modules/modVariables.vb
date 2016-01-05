Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frmSourceFormUD As String
    Public frmSourceForm As SAPbouiCOM.Form
    Public frmSourcePMForm As SAPbouiCOM.Form
    Public frmSourceQCOR As SAPbouiCOM.Form

    Public LoalDB As String
    Public intCurrentRow As Integer = 10000


    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public strDocEntry As String
    Public strImportErrorLog As String = ""
    Public companyStorekey As String = ""

    Public intSelectedMatrixrow As Integer = 0
    Public strSourceformEmpID As String = ""
    Public strApprovalType As String = ""


    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum
    Public Const frm_GoodsReceipt As String = "143"
    Public Const frm_GRReceipt As String = "frm_GoodsReceiptCost"
    Public Const xml_GRREceipt As String = "frm_GoodsReceiptCost.xml"

    Public Const frm_Invoice As String = "133"
    Public Const mnu_Grossprofit As String = "5891"
    Public Const frm_Grossprofit As String = "241"






  
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_DuplicateRow As String = "1294"

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
   

End Module
