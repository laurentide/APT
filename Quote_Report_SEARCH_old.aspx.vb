Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic

Public Class Quote_Report_SEARCH
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

        If Session("OS") <> Nothing Or User.IsInRole("LCLMTL\LCL_Apt") Or _
        User.IsInRole("LCLMTL\LCL_AE") Or User.IsInRole("LCLMTL\LCL_SA") Then
            If Not (Request.QueryString("Status") = Nothing _
                And Request.QueryString("OsNo") = Nothing And Request.QueryString("BillName") = Nothing _
                And Request.QueryString("CustNo") = Nothing And Request.QueryString("QuotedBy") = Nothing _
                And Request.QueryString("City") = Nothing And Request.QueryString("WFollowUp") = Nothing _
                And Request.QueryString("PC") = Nothing And Request.QueryString("StartDate") = Nothing _
                And Request.QueryString("QuoteNo") = Nothing And Request.QueryString("EndDate") = Nothing _
                And Request.QueryString("ModelNumber") = Nothing And Request.QueryString("FUStartDate") = Nothing _
                And Request.QueryString("NetPrice") = Nothing And Request.QueryString("FUEndDate") = Nothing _
                And Request.QueryString("LastName") = Nothing _
                And Request.QueryString("") = Nothing) Then

                Dim dtTable As New DataTable
                Dim cmdTable As OleDbDataAdapter

                Dim strReq As String = "SELECT QuoteId AS [Quote #], Revision, CONVERT(varchar, [Date], 101) AS [Quote Date mm/dd/yyyy], CustomerID AS [Customer #], " & _
                                    "Name AS Customer, City, FirstName AS [First Name], LastName AS [Last Name], Telephone, Extension AS [Ext.], " & _
                                    "Item AS [Item #], ModelNumber AS [Model Number], ProductCode AS PC, Status, " & _
                                    "StatusReason AS [Status Reason], Ref1 AS Reference, Qty, USListPrice AS [US List Price], Discount, " & _
                                    "NetPrice AS [Net Price], Currency, ExchRate AS Rate, CDNList AS [CDN List], Total, " & _
                                    "Os1 AS [Os 1], Os2 AS [Os 2], QuotedBy AS [Quoted By], FollowUpDate AS [F-U Date mm/dd/yyyy], WhodoFollowup AS [Who to F-U]" & _
                                    "FROM OSREPORT WHERE " & Where() & " ORDER BY QuoteId, [Date], Item"


                cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
                cmdTable.Fill(dtTable)

                FillCTRTable(dtTable)
            End If
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

        strWhere += "QuoteId like '%" & Request.QueryString("QuoteNo") & "%' AND "

        If Request.QueryString("StartDate") <> Nothing And Request.QueryString("EndDate") <> Nothing Then
            strWhere += "[Date] >= '" & Request.QueryString("StartDate") & "' AND [Date] <= '" & Request.QueryString("EndDate") & "' AND "
        ElseIf Request.QueryString("StartDate") <> Nothing Then
            strWhere += "[Date] >= '" & Request.QueryString("StartDate") & "' AND "
        ElseIf Request.QueryString("EndDate") <> Nothing Then
            strWhere += "[Date] <= '" & Request.QueryString("EndDate") & "' AND "
        End If

        If Request.QueryString("FUStartDate") <> Nothing And Request.QueryString("FUEndDate") <> Nothing Then
            strWhere += "FollowUpDate >= '" & Request.QueryString("FUStartDate") & "' AND FollowUpDate <= '" & Request.QueryString("FUEndDate") & "' AND "
        ElseIf Request.QueryString("FUStartDate") <> Nothing Then
            strWhere += "FollowUpDate >= '" & Request.QueryString("FUStartDate") & "' AND "
        ElseIf Request.QueryString("FUEndDate") <> Nothing Then
            strWhere += "FollowUpDate <= '" & Request.QueryString("FUEndDate") & "' AND "
        End If

        strWhere += "Status like '%" & Request.QueryString("Status") & "%' AND  "
        strWhere += "City like '%" & Request.QueryString("City") & "%' AND  "
        If Request.QueryString("OsNo") <> "" Then
            strWhere += "OSNO = '" & Request.QueryString("OsNo") & "' AND  "
        End If
        strWhere += "Name like '%" & Request.QueryString("BillName") & "%' AND  "
        strWhere += IIf(Request.QueryString("CustNo") <> Nothing, "CustomerID =" & Request.QueryString("CustNo") & " AND  ", "")
        strWhere += IIf(Request.QueryString("NetPrice") <> Nothing, "NetPrice = " & Request.QueryString("NetPrice") & " AND  ", "")
        strWhere += "ModelNumber like '%" & Request.QueryString("ModelNumber") & "%' AND "
        strWhere += "ProductCode like '%" & Request.QueryString("PC") & "%' AND "
        If Request.QueryString("QuotedBy") <> Nothing Then
            strWhere += "QuotedBy = '" & Request.QueryString("QuotedBy") & "' AND "
        End If
        If Request.QueryString("WFollowUp") <> Nothing Then
            strWhere += "WhodoFollowup = '" & Request.QueryString("WFollowUp") & "' AND "
        End If
        strWhere += "LastName like '%" & Request.QueryString("LastName") & "%'"

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
