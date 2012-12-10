Imports System.Data
Imports System.Data.OleDb

Public Class Bookings_Results
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Report As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents TabReport As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Table1 As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents year As System.Web.UI.WebControls.Label

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
        year.Text = Request.QueryString("year1")
        CreateReport()
    End Sub

    Sub CreateReport()
        Dim TableauRapport As PlaceHolder = FindControl("Report")
        Dim dtTableDivisions As New DataTable
        Dim dtTableCustomersRegions(1) As DataTable
        Dim dtTableRegions As New DataTable
        Dim dtTableDonnees(1) As DataTable
        Dim dvCustomers As DataView
        Dim myRow As DataRow
        Dim intNbElements As Integer
        Dim TotalColonnes() As Integer

        Dim unTableau As New Table
        unTableau.CssClass = "EspaceTableau"

        Dim dbConnSQLServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSQLServer.Open()

        Dim strReq As String

        '1
        strReq = "SELECT PC, PrimaryCat, ProductDesc FROM PCNOMIS where " & WhereDivision() & " Group by PC, PrimaryCat, ProductDesc order by PrimaryCat, PC"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSQLServer)
        cmdTable.Fill(dtTableDivisions)

        '2
        strReq = "Select Customers.CustomerNo, CustomerName from NOMIS, Customers where Customers.CustomerNo = NOMIS.CustomerNo AND AB='A' " & WhereOs("Customers") & " Group by Customers.CustomerNo, CustomerName order by CustomerName"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(0) = New DataTable
        dtTableCustomersRegions(0).TableName = "CustomersA"
        cmdTable.Fill(dtTableCustomersRegions(0))

        strReq = "Select RegionNo, Region From Regions where 1=1 " & WhereOs("Regions") & " Group by RegionNo, Region Order by Region"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(1) = New DataTable
        dtTableCustomersRegions(1).TableName = "CustomersB"
        cmdTable.Fill(dtTableCustomersRegions(1))

        'Réserve le nombre de colonnes pour les totaux en bas de page
        intNbElements = (dtTableCustomersRegions(0).Rows.Count + dtTableCustomersRegions(1).Rows.Count) * 4 - 1 + 4
        ReDim TotalColonnes(intNbElements)

        'Ligne contenant les nos et les noms des customers
        unTableau.Controls.Add(AfficheNoCustomers(dtTableCustomersRegions))
        unTableau.Controls.Add(AfficheNomCustomers(dtTableCustomersRegions))

        'Affichage l'entête entre les noms des customers et les données
        unTableau.Controls.Add(AfficheEntete(dtTableCustomersRegions))

        'Remplit les données selon le customer et l'initiative
       strReq = "SELECT dbo.vPCFYCUstomers.CustomerNo AS [NO], dbo.vPCFYCUstomers.PC, dbo.vForecastA.Forecast, ROUND(vBookingsA_1.BOOKINGS, 0) " & _
                      "AS BOOKINGSLY, ROUND(dbo.vBookingsA.BOOKINGS, 0) AS BOOKINGSAC, ROUND(dbo.vQuoteA.Total, 0) AS QuoteTotal " & _
                      "FROM dbo.vPCFYCUstomers LEFT OUTER JOIN " & _
                      "dbo.vBookingsA vBookingsA_1 ON dbo.vPCFYCUstomers.CustomerNo = vBookingsA_1.CUSTOMERNO AND " & _
                      "dbo.vPCFYCUstomers.OSNo = vBookingsA_1.OSNO AND dbo.vPCFYCUstomers.PC = vBookingsA_1.PC AND " & _
                      "dbo.vPCFYCUstomers.[FY-1] = vBookingsA_1.FY LEFT OUTER JOIN " & _
                      "dbo.vQuoteA ON dbo.vPCFYCUstomers.CustomerNo = dbo.vQuoteA.Customer AND dbo.vPCFYCUstomers.PC = dbo.vQuoteA.PC LEFT OUTER JOIN " & _
                      "dbo.vBookingsA ON dbo.vPCFYCUstomers.CustomerNo = dbo.vBookingsA.CUSTOMERNO AND " & _
                      "dbo.vPCFYCUstomers.OSNo = dbo.vBookingsA.OSNO AND dbo.vPCFYCUstomers.FY = dbo.vBookingsA.FY AND " & _
                      "dbo.vPCFYCUstomers.PC = dbo.vBookingsA.PC LEFT OUTER JOIN " & _
                      "dbo.vForecastA ON dbo.vPCFYCUstomers.PC = dbo.vForecastA.PC AND dbo.vPCFYCUstomers.CustomerNo = dbo.vForecastA.CustomerNo AND " & _
                      "dbo.vPCFYCUstomers.OSNo = dbo.vForecastA.OSNo And dbo.vPCFYCUstomers.FY = dbo.vForecastA.FY " & _
                      "WHERE (dbo.vPCFYCUstomers.OSNo = '" & Request.QueryString("OS") & "') AND (dbo.vPCFYCUstomers.FY = '" & Request.QueryString("year1") & "')"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableDonnees(0) = New DataTable

        cmdTable.Fill(dtTableDonnees(0))

        strReq = "SELECT     dbo.vPCFYREGION.RegionNo AS [NO], dbo.vPCFYREGION.PC, dbo.vForecastB.Forecast, ROUND(vBookingsB_1.BOOKINGS, 0) AS BOOKINGSLY, " & _
                      "ROUND(dbo.vBookingsB.BOOKINGS, 0) AS BOOKINGSAC, ROUND(dbo.vQuoteB.QuoteTotal, 0) AS QuoteTotal " & _
                      "FROM dbo.vPCFYREGION LEFT OUTER JOIN " & _
                      "dbo.vBookingsB vBookingsB_1 ON dbo.vPCFYREGION.PC = vBookingsB_1.PC AND dbo.vPCFYREGION.RegionNo = vBookingsB_1.RegionNo AND " & _
                      "dbo.vPCFYREGION.[FY-1] = vBookingsB_1.FY LEFT OUTER JOIN " & _
                      "dbo.vQuoteB ON dbo.vPCFYREGION.RegionNo = dbo.vQuoteB.RegionNo AND dbo.vPCFYREGION.PC = dbo.vQuoteB.PC LEFT OUTER JOIN " & _
                      "dbo.vBookingsB ON dbo.vPCFYREGION.PC = dbo.vBookingsB.PC AND dbo.vPCFYREGION.FY = dbo.vBookingsB.FY AND " & _
                      "dbo.vPCFYREGION.RegionNo = dbo.vBookingsB.RegionNo LEFT OUTER JOIN " & _
                      "dbo.vForecastB ON dbo.vPCFYREGION.PC = dbo.vForecastB.PC AND dbo.vPCFYREGION.FY = dbo.vForecastB.FY AND " & _
                      "dbo.vPCFYREGION.RegionNo = dbo.vForecastB.RegionNo " & _
                      "WHERE (dbo.vPCFYREGION.FY = '" & Request.QueryString("year1") & "') AND (dbo.vPCFYREGION.OsNo = '" & Request.QueryString("OS") & "')"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableDonnees(1) = New DataTable
        cmdTable.Fill(dtTableDonnees(1))

        Dim temp As String = ""
        For Each myRow In dtTableDivisions.Rows
            unTableau.Controls.Add(AfficheDonnees(myRow, dtTableCustomersRegions, dtTableDonnees, temp, TotalColonnes))
            temp = myRow(1)
        Next

        unTableau.Controls.Add(AfficheTotauxColonnes(TotalColonnes))
        TableauRapport.Controls.Add(unTableau)

        dbConnSQLServer.Close()

    End Sub

    Function AfficheNomCustomers(ByVal dtTableCustomersRegions() As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 3
        uneLigne.Cells.Add(uneCellule)
        Dim myRow As DataRow
        Dim i As Integer
        'Afficher les noms des customers

        uneLigne.CssClass = "EnteteRapport"

        For i = 0 To UBound(dtTableCustomersRegions)
            For Each myRow In dtTableCustomersRegions(i).Rows
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneCellule.ColumnSpan = 4
                uneCellule.HorizontalAlign = HorizontalAlign.Center
                'Mettre le nom du customer au lieu du numéro (bd NOMIS)
                uneCellule.Text = myRow(1).ToString()
                uneLigne.Cells.Add(uneCellule)
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 4
        uneCellule.HorizontalAlign = HorizontalAlign.Center
        'Mettre le nom du customer au lieu du numéro (bd NOMIS)
        uneCellule.Text = "Total"
        uneCellule.CssClass = "Gras"
        uneLigne.Cells.Add(uneCellule)

        AfficheNomCustomers = uneLigne
    End Function

    Function AfficheNoCustomers(ByVal dtTableCustomersRegions() As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 3
        uneLigne.Cells.Add(uneCellule)
        Dim myRow As DataRow
        Dim i As Integer
        'Afficher les noms des customers

        uneLigne.CssClass = "EnteteRapport"

        For i = 0 To UBound(dtTableCustomersRegions)
            For Each myRow In dtTableCustomersRegions(i).Rows
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneCellule.ColumnSpan = 4
                uneCellule.HorizontalAlign = HorizontalAlign.Center
                'Mettre le nom du customer au lieu du numéro (bd NOMIS)
                uneCellule.Text = myRow(0).ToString()
                uneLigne.Cells.Add(uneCellule)
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 3
        uneCellule.HorizontalAlign = HorizontalAlign.Center
        'Mettre le nom du customer au lieu du numéro (bd NOMIS)
        uneCellule.Text = "Total"
        uneCellule.CssClass = "Gras"
        uneLigne.Cells.Add(uneCellule)

        AfficheNoCustomers = uneLigne
    End Function

    Function AfficheEntete(ByVal dtTableCustomersRegions() As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        Dim myrow As DataRow
        Dim strTitres() As String = {"Forecast", "Last Year", "Actual", "Quoted"}
        Dim i As Integer
        Dim index As Integer

        uneLigne.CssClass = "SousTitresRapport"

        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = "Division"
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = "PC"
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = "Description"
        uneLigne.Cells.Add(uneCellule)

        For index = 0 To UBound(dtTableCustomersRegions)
            For Each myrow In dtTableCustomersRegions(index).Rows
                For i = 0 To UBound(strTitres)
                    uneCellule = New TableCell
                    uneCellule.BorderWidth = Unit.Pixel(1)
                    uneCellule.HorizontalAlign = HorizontalAlign.Center
                    uneCellule.Width = New Unit(80)
                    uneCellule.Text = strTitres(i)
                    uneLigne.Cells.Add(uneCellule)
                Next

                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneCellule.Width = New Unit(60)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        For i = 0 To UBound(strTitres)
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneCellule.Width = New Unit(80)
            uneCellule.Text = strTitres(i)
            uneLigne.Cells.Add(uneCellule)
        Next

        AfficheEntete = uneLigne
    End Function

    Function AfficheDonnees(ByVal myRow As DataRow, ByVal dtTableCustomersRegions() As DataTable, ByVal dtTableDonnees() As DataTable, ByVal Grouping As String, ByRef TotalColonnes() As Integer) As TableRow
        Dim dvDonnees As DataView
        Dim totals(3) As Integer

        Dim fields() As String = {"CustomerNo", "RegionNo"}
        Dim myRow1 As DataRow
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim i As Integer
        Dim indice As Integer

        'Écrit le nom de la division, du productCode et la description
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell

        uneLigne.CssClass = "TexteGRapport"

        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = IIf(Trim(myRow(1)) <> Trim(Grouping), myRow(1), "")
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = myRow(0)
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = myRow(2)
        uneLigne.Cells.Add(uneCellule)

        For i = 0 To UBound(dtTableCustomersRegions)
            dvDonnees = New DataView(dtTableDonnees(i))
            For Each myRow1 In dtTableCustomersRegions(i).Rows
                dvDonnees.RowFilter = "NO=" & myRow1(0) & " AND PC='" & myRow(0) & "'"
                Dim index As Integer
                For index = 2 To dvDonnees.Table.Columns.Count - 1
                    uneCellule = New TableCell
                    uneCellule.BorderWidth = Unit.Pixel(1)
                    If dvDonnees.Count > 0 Then
                        If Not IsDBNull(dvDonnees(0)(index)) Then
                            uneCellule.Text = dvDonnees(0)(index)
                            uneCellule.HorizontalAlign = HorizontalAlign.Center
                        End If
                    End If
                    uneLigne.Cells.Add(uneCellule)

                    If dvDonnees.Table.Columns(index).ToString = "Forecast" Then
                        totals(0) += Val(uneCellule.Text)
                    ElseIf dvDonnees.Table.Columns(index).ToString = "BOOKINGSLY" Then
                        totals(1) += Val(uneCellule.Text)
                    ElseIf dvDonnees.Table.Columns(index).ToString = "BOOKINGSAC" Then
                        totals(2) += Val(uneCellule.Text)
                    ElseIf dvDonnees.Table.Columns(index).ToString = "QuoteTotal" Then
                        totals(3) += Val(uneCellule.Text)
                    End If

                    TotalColonnes(indice) += Val(uneCellule.Text)
                    indice += 1
                Next
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        For i = 0 To UBound(totals)
            TotalColonnes(indice) += totals(i)
            indice += 1
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneCellule.Text = totals(i)
            uneLigne.Cells.Add(uneCellule)
        Next

        AfficheDonnees = uneLigne
    End Function

    Function AfficheTotauxColonnes(ByVal TotalColonnes() As Integer) As TableRow
        Dim myRow As New TableRow
        Dim uneCellule As TableCell
        Dim i As Integer

        uneCellule = New TableCell
        uneCellule.ColumnSpan = 3
        uneCellule.Text = "Total"
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.HorizontalAlign = HorizontalAlign.Left

        myRow.Cells.Add(uneCellule)

        For i = 0 To UBound(TotalColonnes)
            If i > 0 And i Mod 4 = 0 Then
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                myRow.Cells.Add(uneCellule)
            End If
            uneCellule = New TableCell
            uneCellule.Text = TotalColonnes(i)
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            myRow.Cells.Add(uneCellule)
        Next

        AfficheTotauxColonnes = myRow

    End Function

    Function WhereOs(ByVal strTable As String) As String
        Dim strWhere As String = ""
        If Request.QueryString("OS") <> Nothing Then
            If Request.QueryString("OS") <> "" And Request.QueryString("OS") <> 0 Then
                strWhere = " AND " & strTable & ".OSNo='" & IIf(Request.QueryString("OS").Length = 2, "0", "") & Request.QueryString("OS") & "'"
            End If
        End If
        WhereOs = strWhere
    End Function

    Function WhereDivision() As String
        Dim strWhere As String = "PrimaryCat <> '' OR PrimaryCat <> NULL "
        If Request.QueryString("Division") <> Nothing And Request.QueryString("Division") <> "" And UCase(Request.QueryString("Division")) <> "ALL" Then
            strWhere = " PrimaryCat='" & Request.QueryString("Division") & "'"
        End If
        WhereDivision = strWhere

    End Function


    '|--------------------------------------------Connexions--------------------------------------------|

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
        Tableau.Width = "100px"

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
