Imports System.Data
Imports System.Data.OleDb

Public Class RPT_Multi_Year_Bookings_Results
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

        Dim strReq As String = "SELECT PC, PrimaryCat, ProductDesc FROM Nomis " & WhereDivision() & " Group by PC, PrimaryCat, ProductDesc order by PrimaryCat, ProductDesc"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSQLServer)
        cmdTable.Fill(dtTableDivisions)

        strReq = "Select ForecastA.CustomerNo, CustomerName from NOMIS, ForecastA where ForecastA.CustomerNo = NOMIS.CustomerNo " & WhereOs("ForecastA") & " Group by ForecastA.CustomerNo, CustomerName order by CustomerName"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(0) = New DataTable
        dtTableCustomersRegions(0).TableName = "ForecastA"
        cmdTable.Fill(dtTableCustomersRegions(0))

        strReq = "Select ForecastB.RegionNo, Region From Regions, ForecastB where Regions.RegionNo = ForecastB.RegionNo " & WhereOs("Regions") & " Group by ForecastB.RegionNo, Region Order by Region"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(1) = New DataTable
        dtTableCustomersRegions(1).TableName = "ForecastB"
        cmdTable.Fill(dtTableCustomersRegions(1))

        'Réserve le nombre de colonnes pour les totaux en bas de page
        intNbElements = (dtTableCustomersRegions(0).Rows.Count + dtTableCustomersRegions(1).Rows.Count) * (Request.QueryString("year2") - Request.QueryString("year1") + 1) - 1 + _
                        (Request.QueryString("year2") - Request.QueryString("year1") + 1)
        ReDim TotalColonnes(intNbElements)

        'Ligne contenant les nom des customers
        unTableau.Controls.Add(AfficheNoCustomers(dtTableCustomersRegions))
        unTableau.Controls.Add(AfficheNomCustomers(dtTableCustomersRegions))

        'Affichage l'entête entre les noms des customers et les données
        unTableau.Controls.Add(AfficheEntete(dtTableCustomersRegions))

        'Remplit les données selon le customer et l'initiative
        strReq = "SELECT FY, CUSTOMERNO AS [NO], CAST(SUM(BOOKINGS) AS NUMERIC), PC " & _
                    "FROM NOMIS " & _
                    "WHERE (FY BETWEEN '" & Request.QueryString("Year1") & "' AND '" & Request.QueryString("Year2") & "') " & _
                    IIf(Request.QueryString("OS") <> Nothing And Request.QueryString("OS") <> "ALL", " and NOMIS.OsNo='" & Request.QueryString("OS") & "' ", "") & _
                    "GROUP BY FY, CUSTOMERNO, PC " & _
                    "ORDER BY CUSTOMERNO, PC"

        Trace.Warn(strReq)
        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableDonnees(0) = New DataTable
        cmdTable.Fill(dtTableDonnees(0))

        strReq = "SELECT FY, RegionNo AS [NO], CAST(SUM(BOOKINGS) AS NUMERIC) AS BOOKINGS, PC " & _
                    "FROM Customers INNER JOIN " & _
                    "NOMIS ON Customers.CustomerNo = NOMIS.CUSTOMERNO " & _
                    "WHERE (FY BETWEEN '" & Request.QueryString("Year1") & "' AND '" & Request.QueryString("Year2") & "') " & _
                    IIf(Request.QueryString("OS") <> Nothing And Request.QueryString("OS") <> "ALL", " and NOMIS.OsNo='" & Request.QueryString("OS") & "' ", "") & _
                    "GROUP BY FY, RegionNo, PC"
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
                uneCellule.ColumnSpan = Request.QueryString("Year2") - Request.QueryString("Year1") + 1
                uneCellule.HorizontalAlign = HorizontalAlign.Center

                uneCellule.Text = myRow(1).ToString()
                uneLigne.Cells.Add(uneCellule)
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = Request.QueryString("Year2") - Request.QueryString("Year1") + 1
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
                uneCellule.ColumnSpan = Request.QueryString("Year2") - Request.QueryString("Year1") + 1
                uneCellule.HorizontalAlign = HorizontalAlign.Center

                uneCellule.Text = myRow(0).ToString()
                uneLigne.Cells.Add(uneCellule)
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = Request.QueryString("Year2") - Request.QueryString("Year1") + 1
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
                For i = Request.QueryString("Year1") To Request.QueryString("Year2")
                    uneCellule = New TableCell
                    uneCellule.BorderWidth = Unit.Pixel(1)
                    uneCellule.HorizontalAlign = HorizontalAlign.Center
                    uneCellule.Text = i.ToString
                    uneLigne.Cells.Add(uneCellule)
                Next

                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneCellule.Width = New Unit(60)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        For i = Request.QueryString("Year1") To Request.QueryString("Year2")
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneCellule.Text = i.ToString
            uneLigne.Cells.Add(uneCellule)
        Next


        AfficheEntete = uneLigne
    End Function

    Function AfficheDonnees(ByVal myRow As DataRow, ByVal dtTableCustomersRegions() As DataTable, ByVal dtTableDonnees() As DataTable, ByVal Grouping As String, ByRef TotalColonnes() As Integer) As TableRow
        Dim dvDonnees As DataView

        Dim myRow1 As DataRow
        Dim strReq As String
        Dim i As Integer
        Dim index As Integer
        Dim Ligne As Integer
        Dim indice As Integer
        Dim totals(Request.QueryString("Year2") - Request.QueryString("Year1")) As Integer

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
                For index = Request.QueryString("Year1") To Request.QueryString("Year2")
                    dvDonnees.RowFilter = "NO=" & myRow1(0) & " AND FY='" & index & "'" & " AND PC='" & myRow(0) & "'"
                    uneCellule = New TableCell
                    uneCellule.BorderWidth = Unit.Pixel(1)
                    Ligne = index - Request.QueryString("Year1")
                    If dvDonnees.Count > 0 Then
                        If Not IsDBNull(dvDonnees(0)(0)) Then
                            If dvDonnees(0)(0) = index Then
                                If Not IsDBNull(dvDonnees(0)(2)) Then
                                    uneCellule.Text = dvDonnees(0)(2)
                                Else
                                    uneCellule.Text = ""
                                End If
                            End If
                            uneCellule.HorizontalAlign = HorizontalAlign.Center
                        End If
                    End If
                    totals(index - Request.QueryString("Year1")) += Val(uneCellule.Text)
                    TotalColonnes(indice) += Val(uneCellule.Text)
                    uneLigne.Cells.Add(uneCellule)
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

    Function WhereOs(ByVal strTable As String) As String
        Dim strWhere As String = ""
        If Request.QueryString("OS") <> Nothing Then
            If Request.QueryString("OS") <> "" And Request.QueryString("OS") <> 0 Then
                strWhere = " AND " & strTable & ".OSNo='" & IIf(Request.QueryString("OS").Length = 2, "0", "") & Request.QueryString("OS") & "'"
            End If
        End If
        WhereOs = strWhere
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
            If i > 0 And i Mod (Request.QueryString("year2") - Request.QueryString("year1") + 1) = 0 Then
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

    Function WhereDivision() As String
        Dim strWhere As String = ""
        If Request.QueryString("Division") <> Nothing And Request.QueryString("Division") <> "" And UCase(Request.QueryString("Division")) <> "ALL" Then
            strWhere = " where PrimaryCat='" & Request.QueryString("Division") & "'"
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
