Imports System.Data.OleDb
Imports System.Data


Public Class RPT_Product_Specialists
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Report As System.Web.UI.WebControls.PlaceHolder
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
        CreateReport()
    End Sub

    Sub CreateReport()
        Dim TableauRapport As PlaceHolder = FindControl("Report")
        Dim dtTableDivisions As New DataTable
        Dim dtTableCustomersRegions(1) As DataTable
        Dim dtTableRegions As New DataTable
        Dim dtTableDonnees As DataTable
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
        strReq = "SELECT PrimaryCat, SecondaryCat FROM PCNOMIS WHERE " & WhereDivision() & " Group by PrimaryCat, SecondaryCat Order By PrimaryCat, SecondaryCat "
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSQLServer)
        cmdTable.Fill(dtTableDivisions)

        unTableau.Controls.Add(AffichePrimaryCat(dtTableDivisions))
        unTableau.Controls.Add(AfficheSecondaryCat(dtTableDivisions))

        unTableau.Controls.Add(AfficheEntete(dtTableDivisions))

        '2
        strReq = "Select Customers.CustomerNo, CustomerName, Customers.OSNO from NOMIS, Customers where Customers.CustomerNo = NOMIS.CustomerNo AND AB='A' " & WhereOs("Customers") & " Group by Customers.CustomerNo, CustomerName, Customers.OsNO order by Customers.OsNO, Customers.CustomerNo"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(0) = New DataTable
        dtTableCustomersRegions(0).TableName = "CustomersA"
        cmdTable.Fill(dtTableCustomersRegions(0))

        strReq = "Select RegionNo, Region, OSNO From Regions where 1=1 " & WhereOs("Regions") & " Group by RegionNo, Region, OSNo Order by Region"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(1) = New DataTable
        dtTableCustomersRegions(1).TableName = "CustomersB"
        cmdTable.Fill(dtTableCustomersRegions(1))

        'Remplit les données selon des customersA
        strReq = "SELECT dbo.vPCFYCUstomers.PRIMARYCAT, dbo.vPCFYCUstomers.SECONDARYCAT, dbo.vPCFYCUstomers.CustomerNo AS [NO], " & _
                      "dbo.vPCFYCUstomers.OSNo, SUM(ROUND(dbo.vBookingsA.BOOKINGS, 0)) AS Bookings " & _
                      ", SUM(dbo.vForecastA.Forecast) AS Forecast FROM dbo.vPCFYCUstomers LEFT OUTER JOIN " & _
                      "dbo.vBookingsA ON dbo.vPCFYCUstomers.CustomerNo = dbo.vBookingsA.CUSTOMERNO AND " & _
                      "dbo.vPCFYCUstomers.OSNo = dbo.vBookingsA.OSNO AND dbo.vPCFYCUstomers.FY = dbo.vBookingsA.FY AND " & _
                      "dbo.vPCFYCUstomers.PC = dbo.vBookingsA.PC LEFT OUTER JOIN " & _
                      "dbo.vForecastA ON dbo.vPCFYCUstomers.PC = dbo.vForecastA.PC AND dbo.vPCFYCUstomers.CustomerNo = dbo.vForecastA.CustomerNo AND " & _
                      "dbo.vPCFYCUstomers.OSNo = dbo.vForecastA.OSNo And dbo.vPCFYCUstomers.FY = dbo.vForecastA.FY " & _
                      "WHERE 1=1 " & WhereOs("vPCFYCUstomers") & " AND (dbo.vPCFYCUstomers.FY = '" & Request.QueryString("year1") & "') " & _
                      "GROUP BY dbo.vPCFYCUstomers.PRIMARYCAT, dbo.vPCFYCUstomers.SECONDARYCAT, dbo.vPCFYCUstomers.CustomerNo, dbo.vPCFYCUstomers.OSNo"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableDonnees = New DataTable
        cmdTable.Fill(dtTableDonnees)

        For Each myRow In dtTableCustomersRegions(0).Rows
            unTableau.Controls.Add(AfficheDonnees(myRow, dtTableDivisions, dtTableDonnees))
        Next

        strReq = "SELECT dbo.vPCFYREGION.OsNo, dbo.vPCFYREGION.RegionNo AS [NO], dbo.vPCFYREGION.PRIMARYCAT, dbo.vPCFYREGION.SECONDARYCAT, " & _
                      "SUM(ROUND(dbo.vBookingsB.BOOKINGS, 0)) AS BOOKINGS, SUM(dbo.vForecastB.Forecast) AS Forecast " & _
                      "FROM dbo.vPCFYREGION LEFT OUTER JOIN " & _
                      "dbo.vBookingsB ON dbo.vPCFYREGION.PC = dbo.vBookingsB.PC AND dbo.vPCFYREGION.FY = dbo.vBookingsB.FY AND " & _
                      "dbo.vPCFYREGION.RegionNo = dbo.vBookingsB.RegionNo LEFT OUTER JOIN " & _
                      "dbo.vForecastB ON dbo.vPCFYREGION.PC = dbo.vForecastB.PC AND dbo.vPCFYREGION.FY = dbo.vForecastB.FY AND " & _
                      "dbo.vPCFYREGION.RegionNo = dbo.vForecastB.RegionNo " & _
                      "WHERE (dbo.vPCFYREGION.FY = '" & Request.QueryString("year1") & "') " & WhereOs("vPCFYREGION") & " " & _
                      "GROUP BY dbo.vPCFYREGION.PRIMARYCAT, dbo.vPCFYREGION.SECONDARYCAT, dbo.vPCFYREGION.RegionNo, dbo.vPCFYREGION.OsNo "

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableDonnees = New DataTable
        cmdTable.Fill(dtTableDonnees)

        For Each myRow In dtTableCustomersRegions(1).Rows
            unTableau.Controls.Add(AfficheDonnees(myRow, dtTableDivisions, dtTableDonnees))
        Next

        TableauRapport.Controls.Add(unTableau)

        dbConnSQLServer.Close()

    End Sub

    Function AffichePrimaryCat(ByVal dtTableDivisions As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell

        Dim columnSpan As Integer
        Dim division As String
        Dim valide As Boolean
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 3
        uneLigne.Cells.Add(uneCellule)
        Dim myRow As DataRow
        Dim i As Integer
        'Afficher les noms des customers

        uneLigne.CssClass = "EnteteRapport"

        While i < dtTableDivisions.Rows.Count
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.Text = dtTableDivisions.Rows(i)(0).ToString()
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneCellule.CssClass = "EspacesCTR"

            valide = False
            columnSpan = 0
            division = dtTableDivisions.Rows(i)(0)
            While i < dtTableDivisions.Rows.Count And valide = False

                If dtTableDivisions.Rows(i)(0) <> division Then valide = True
                i += 1
                columnSpan += 1
            End While

            If valide Then
                i -= 1
                columnSpan -= 1
            End If

            uneCellule.ColumnSpan = columnSpan * 2

            uneLigne.Cells.Add(uneCellule)
        End While

        AffichePrimaryCat = uneLigne
    End Function

    Function AfficheSecondaryCat(ByVal dtTableDivisions As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell

        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 3
        uneLigne.Cells.Add(uneCellule)
        Dim myRow As DataRow
        Dim i As Integer

        uneLigne.CssClass = "EnteteRapport"

        For i = 0 To dtTableDivisions.Rows.Count - 1
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneCellule.ColumnSpan = 2
            uneCellule.Text = dtTableDivisions.Rows(i)(1).ToString()
            uneLigne.Cells.Add(uneCellule)
        Next

        AfficheSecondaryCat = uneLigne
    End Function

    Function AfficheEntete(ByVal dtTableDivisions As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        Dim myrow As DataRow
        Dim strTitresGauche() As String = {"Customer#", "Customer Name", "Os"}
        Dim strTitres() As String = {"Actual", "Forecast"}
        Dim i As Integer
        Dim i2 As Integer
        Dim index As Integer

        uneLigne.CssClass = "SousTitresRapport"

        i = 0
        For i = 0 To UBound(strTitresGauche)
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.Text = strTitresGauche(i)
            uneLigne.Cells.Add(uneCellule)
        Next

        For i = 0 To dtTableDivisions.Rows.Count - 1
            For i2 = 0 To UBound(strTitres)
                uneCellule = New TableCell
                uneCellule.CssClass = "EspacesCTR"
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneCellule.Text = strTitres(i2)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        AfficheEntete = uneLigne
    End Function

    Function AfficheDonnees(ByVal myRow As DataRow, ByVal dtTableDivisions As DataTable, ByVal dtTableDonnees As DataTable) As TableRow
        Dim dvDonnees As DataView
        Dim totals(2) As Integer

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

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = myRow(0)
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = myRow(1)
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(2)
        uneCellule.Text = myRow(2)
        uneLigne.Cells.Add(uneCellule)

        For Each myRow1 In dtTableDivisions.Rows
            dvDonnees = New DataView(dtTableDonnees)
            dvDonnees.RowFilter = "PRIMARYCAT='" & Trim(myRow1(0)) & "' AND SECONDARYCAT='" & Trim(myRow1(1)) & "' AND NO=" & myRow(0) & " AND OSNO='" & myRow(2) & "'"
            Dim index As Integer
            For index = 4 To 5
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                If dvDonnees.Count > 0 Then
                    If Not IsDBNull(dvDonnees(0)(index)) Then
                        uneCellule.Text = dvDonnees(0)(index)
                        uneCellule.HorizontalAlign = HorizontalAlign.Center
                    End If
                End If
                uneLigne.Cells.Add(uneCellule)

                indice += 1
            Next
        Next

        AfficheDonnees = uneLigne
    End Function

    Function WhereOs(ByVal strTable As String) As String
        Dim strWhere As String
        If Request.QueryString("OS") <> Nothing Then
            If Request.QueryString("OS") <> "" And Request.QueryString("OS") <> 0 Then
                strWhere = " AND " & strTable & ".OSNo='" & IIf(Request.QueryString("OS").Length = 2, "0", "") & Request.QueryString("OS") & "'"
            End If
        End If
        WhereOs = strWhere
    End Function

    Function WhereDivision() As String
        Dim strWhere As String = "PrimaryCat <> '' OR PrimaryCat <> NULL"
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
