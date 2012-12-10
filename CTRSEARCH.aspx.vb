Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic

Public Class CTRSEARCH
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Report As System.Web.UI.WebControls.Table
    Protected WithEvents TabReport As System.Web.UI.HtmlControls.HtmlTable

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        Dim dbConnSqlServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSqlServer.Open()

        If Session("OS") <> Nothing Or User.IsInRole("LCLMTL\LCL_APT") Or _
             User.IsInRole("LCLMTL\LCL_AE") Or User.IsInRole("LCLMTL\LCL_SA") Or User.IsInRole("LCLMTL\LCL_SIE") Then
            ' If Not (Request.QueryString("OrderNo") = Nothing And Request.QueryString("Description") = Nothing _
            '   And Request.QueryString("OsNo") = Nothing And Request.QueryString("BillName") = Nothing _
            '   And Request.QueryString("CustNo") = Nothing And Request.QueryString("List") = Nothing _
            '  And Request.QueryString("ShipCity") = Nothing And Request.QueryString("Orig") = Nothing _
            '   And Request.QueryString("CustPONo") = Nothing And Request.QueryString("PC") = Nothing _
            '  And Request.QueryString("QuoteNo") = Nothing And Request.QueryString("CustInvItem") = Nothing _
            '   And Request.QueryString("") = Nothing) Then
            Dim strReq As String
            If Request.QueryString("SerialNumber") <> "" Then
                strReq = "Select ORDERNO AS [Order #], DISCOUNT_RATE as [Discount Rate], QUOTENO AS [Quote #], OrderDate as [Order date], LineItemNo as [Line #], " & _
                               "Qty, Inventory_Number as [Inventory Number],DESCRIPTION1 as [Model Number], DESCRIPTION2 as [Customer Inventory No.], List AS [Unit Price], FIRSTREQDATE AS [1stReqS], OSNO as [Sls #], " & _
                               "BillName AS [Customer Name], CustNo AS [Cust #], ShipAddr AS [Shipping ADDRESS], " & _
                               "ShipCity AS [Shipping City], Orig, CUSTPONO AS [CUT PO #], ProductCodeCategory,PC, Currency, CONVERT(money ,(Select TOP 1 CUREXCHANGE FROM CURRENCY WHERE CTR1.OrderDate >= CURDATE AND CURCODE='US$' ORDER BY CURDATE DESC)) AS [Exchange], ORDERTYPE AS [Order Type], " & _
                               "CASE WHEN ORDERTYPE = 99 THEN -(List) ELSE List END*QTy AS [Extension], Convert(money , Round(CASE WHEN ORDERTYPE = 99 THEN -(List) ELSE List END*QTy * (Select TOP 1 CUREXCHANGE FROM CURRENCY WHERE CTR1.OrderDate >= CURDATE AND CURCODE=CTR1.CURRENCY ORDER BY CURDATE DESC),2),1) AS [CD$], Serial_Number as [Serial Number] from CTR CTR1 WHERE " & Where() & _
                               " Order By OrderDate, ORDERNO, LineItemNo "
            Else
                strReq = strReq & " SELECT  ORDERNO AS [ORDER #] , DISCOUNT_RATE AS [DISCOUNT RATE] , QUOTENO AS [QUOTE #] , ORDERDATE AS [ORDER DATE] , LINEITEMNO AS [LINE #] , QTY , Inventory_Number as [Inventory Number],DESCRIPTION1 AS [MODEL NUMBER] , DESCRIPTION2 AS [CUSTOMER INVENTORY NO.] , LIST AS [UNIT PRICE] , FIRSTREQDATE AS [1STREQS] , OSNO AS [SLS #] , BILLNAME AS [CUSTOMER NAME] , CUSTNO AS [CUST #] , SHIPADDR AS [SHIPPING ADDRESS] , SHIPCITY AS [SHIPPING CITY] , ORIG , CUSTPONO AS [CUT PO #] , ProductCodeCategory as [Product Code Category],PC , CURRENCY , (CONVERT(money ,(SELECT TOP 1 CUREXCHANGE "
                strReq = strReq & "         FROM    CURRENCY                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         "
                strReq = strReq & "         WHERE   CTR1.OrderDate >= CURDATE AND CURCODE ='US$'                                                                                                                                                                                                                                                                                                                                                                                                                                                     "
                strReq = strReq & "         ORDER BY CURDATE DESC                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    "
                strReq = strReq & "         ))) AS [Exchange] , ORDERTYPE AS [ORDER TYPE], CASE WHEN ORDERTYPE = 99 THEN -(LIST) ELSE LIST END * QTY AS [EXTENSION], (CONVERT(money , ROUND( CASE WHEN ORDERTYPE = 99 THEN -(List) ELSE List END*QTy *                                                                                                                                                                                                                                                                                               "
                strReq = strReq & "         (SELECT TOP 1 CUREXCHANGE                                                                                                                                                                                                                                                                                                                                                                                                                                                                                "
                strReq = strReq & "         FROM    CURRENCY                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         "
                strReq = strReq & "         WHERE   CTR1.OrderDate >= CURDATE AND CURCODE =CTR1.CURRENCY ORDER BY CURDATE DESC                                                                                                                                                                                                                                                                                                                                                                                                                       "
                strReq = strReq & "         ) ,2),1)) AS [CD$]                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       "
                strReq = strReq & " FROM    CTR AS CTR1                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              "
                strReq = strReq & " WHERE " & Where()
                strReq = strReq & " GROUP BY ORDERNO  , DISCOUNT_RATE , QUOTENO , ORDERDATE , LINEITEMNO , QTY , DESCRIPTION1 , DESCRIPTION2 , LIST , FIRSTREQDATE , OSNO , BILLNAME , CUSTNO , SHIPADDR , SHIPCITY , ORIG , CUSTPONO , ProductCodeCategory,PC , CURRENCY , ORDERTYPE,Inventory_Number                                                                                                                                                                                                                                                                                    "
                strReq = strReq & " ORDER BY ORDERDATE, ORDERNO , LINEITEMNO"
            End If
            Dim dtTable As New DataTable
            Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
            cmdTable.Fill(dtTable)

            FillCTRTable(dtTable)
        Else
            Response.Redirect("Denied.html")
        End If


        dbConnSqlServer.Close()

    End Sub

    Sub FillCTRTable(ByVal dtTable As DataTable)
        Dim myRow As DataRow
        Dim myCol As DataColumn

        Dim unTableau As Table = FindControl("Report")
        unTableau.CssClass = "EspaceTableau"
        Dim uneLigne As TableRow
        Dim uneCellule As TableCell

        'Affiche l'entête
        uneLigne = New TableRow
        For Each myCol In dtTable.Columns
            uneCellule = New TableCell
            uneCellule.BorderStyle = BorderStyle.Solid
            uneCellule.CssClass = "BordureTableau texte EspacesCTR Gras"
            uneCellule.Text = myCol.ToString
            uneLigne.Cells.Add(uneCellule)
        Next

        unTableau.Controls.Add(uneLigne)


        For Each myRow In dtTable.Rows
            uneLigne = New TableRow
            For Each myCol In dtTable.Columns
                uneCellule = New TableCell
                uneCellule.BorderStyle = BorderStyle.Solid
                uneCellule.CssClass = "BordureTableau texte EspacesCTR"
                If Not IsDBNull(myRow(myCol)) Then
                    uneCellule.Text = myRow(myCol)
                Else
                    uneCellule.Text = ""
                End If
                uneLigne.Cells.Add(uneCellule)
            Next
            unTableau.Controls.Add(uneLigne)
        Next


    End Sub

    Function Where() As String
        Dim strWhere As String = ""

        strWhere += "ORDERNO like '%" & Request.QueryString("OrderNo") & "%' AND "
        If Request.QueryString("QuoteNo") <> Nothing Then
            strWhere += "QUOTENO like '%" & Request.QueryString("QuoteNo") & "%' AND "
        End If
        If Request.QueryString("OrderDateStart") <> Nothing And Request.QueryString("OrderDateEnd") <> Nothing Then
            strWhere += "ORDERDATE BETWEEN " & Request.QueryString("OrderDateStart") & " AND " & Request.QueryString("OrderDateEnd") & " AND "
        ElseIf Request.QueryString("OrderDateStart") <> Nothing Then
            strWhere += "ORDERDATE >= " & Request.QueryString("OrderDateStart") & " AND "
        ElseIf Request.QueryString("OrderDateEnd") <> Nothing Then
            strWhere += "ORDERDATE <= " & Request.QueryString("OrderDateEnd") & " AND "
        End If

        strWhere += "DESCRIPTION1 like '%" & Request.QueryString("Desc") & "%' AND  "
        strWhere += "DESCRIPTION2 like '%" & Request.QueryString("CustInvItem") & "%' AND  "

        If Request.QueryString("OsNo") <> "" Then
            strWhere += "OSNO = '" & Request.QueryString("OsNo") & "' AND  "
        End If

        strWhere += "BILLNAME like '%" & Request.QueryString("BillName") & "%' AND  "
        strWhere += IIf(Request.QueryString("CustNo") <> Nothing, "CUSTNO =" & Request.QueryString("CustNo") & " AND  ", "")
        strWhere += IIf(Request.QueryString("list") <> Nothing, "LIST = " & Request.QueryString("list") & " AND  ", "")
        strWhere += "SHIPCITY like '%" & Request.QueryString("ShipCity") & "%' AND "
        strWhere += "ORIG like '%" & Request.QueryString("Orig") & "%' AND "
        strWhere += "CUSTPONO like '%" & Request.QueryString("CustPoNo") & "%' AND "
        strWhere += "PC like '%" & Request.QueryString("PC") & "%' "
        If Request.QueryString("DiscountRate") <> "" Then
            strWhere += " AND (Discount_Rate like '%" & Request.QueryString("DiscountRate") & "%')  "
        End If
        If Request.QueryString("InventoryNumber") <> "" Then
            strWhere += " AND (inventory_number like '%" & Request.QueryString("InventoryNumber") & "%')"
        End If
        If Request.QueryString("ProductCodeCategory") <> "" Then
            strWhere += " AND (ProductCodeCategory like '%" & Request.QueryString("ProductCodeCategory") & "%')"
        End If
        If Request.QueryString("SerialNumber") <> "" Then
            strWhere += " AND (Serial_Number like '%" & Request.QueryString("SerialNumber") & "%')"
        End If
        Where = strWhere

    End Function
    '
    '|----------------------------------------------------------------------------------------------|
    '| EtablitConnexionSQLServer												                    |
    '|----------------------------------------------------------------------------------------------|
    Function EtablitConnexionSQLServer() As OleDbConnection
        Dim strConn As String = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=APT;" & _
                                "User ID=APT;Password=APT"
        EtablitConnexionSQLServer = New OleDbConnection(strConn)
    End Function

    Sub ExportToExcel(ByVal sender As Object, ByVal e As EventArgs)
        Dim Tableau As HtmlTable = FindControl("TabReport")
        Tableau.Width = "100%"

        Response.ContentType = "application/vnd.ms-excel"
        ' Remove the charset from the Content-Type header.
        Response.Charset = ""
        ' Turn off the view state.
        Me.EnableViewState = False
        Dim tw As New System.IO.StringWriter
        Dim hw As New HtmlTextWriter(tw)
        ' Get the HTML for the control.
        Tableau.RenderControl(hw)
        ' Write the HTML back to the browser.
        Response.Write(tw.ToString())
        ' End the response.
        Response.End()
    End Sub

End Class
