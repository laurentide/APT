Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic

Public Class RPT_Initiative_Summary
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
        Dim dtTableInitiatives As New DataTable
        Dim dtTableCustomersRegions(1) As DataTable
        Dim dtTableRegions As New DataTable
        Dim dtTableDonnees As New DataTable
        Dim dvCustomers As DataView
        Dim myRow As DataRow

        Dim unTableau As New Table
        unTableau.CssClass = "EspaceTableau"

        Dim dbConnSQLServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSQLServer.Open()

        Dim strReq As String = "SELECT InitiativeNo, Division, Initiative FROM Initiatives " & WhereDivision() & " order by Division, Initiative"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSQLServer)
        cmdTable.Fill(dtTableInitiatives)

        strReq = "Select InitiativesA.CustomerNo, CustomerName from NOMIS, InitiativesA where InitiativesA.CustomerNo = NOMIS.CustomerNo " & WhereOs("InitiativesA") & " Group by InitiativesA.CustomerNo, CustomerName order by CustomerName"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(0) = New DataTable
        dtTableCustomersRegions(0).TableName = "InitiativesA"
        cmdTable.Fill(dtTableCustomersRegions(0))

        strReq = "Select Regions.RegionNo, Region From InitiativesB, Regions where Regions.RegionNo = InitiativesB.RegionNo " & WhereOs("Regions") & " Group by Regions.RegionNo, Region Order by Region"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableCustomersRegions(1) = New DataTable
        dtTableCustomersRegions(1).TableName = "InitiativesB"
        cmdTable.Fill(dtTableCustomersRegions(1))

        'Ligne contenant les nom des customers
        unTableau.Controls.Add(AfficheNomCustomers(dtTableCustomersRegions))

        'Affichage l'entête entre les noms des customers et les données
        unTableau.Controls.Add(AfficheEntete(dtTableCustomersRegions))

        'Remplit les données selon le customer et l'initiative
        Dim temp As String = ""
        For Each myRow In dtTableInitiatives.Rows
            unTableau.Controls.Add(AfficheDonnees(myRow, dtTableCustomersRegions, dbConnSQLServer, temp))
            temp = myRow(1)
        Next

        TableauRapport.Controls.Add(unTableau)

        dbConnSQLServer.Close()

    End Sub

    Function AfficheNomCustomers(ByVal dtTableCustomersRegions() As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 2
        uneLigne.Cells.Add(uneCellule)
        Dim myRow As DataRow
        Dim i As Integer
        'Afficher les noms des customers

        uneLigne.CssClass = "EnteteRapport"

        For i = 0 To UBound(dtTableCustomersRegions)
            For Each myRow In dtTableCustomersRegions(i).Rows
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneCellule.ColumnSpan = 3
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
        uneCellule.ColumnSpan = 3
        uneCellule.HorizontalAlign = HorizontalAlign.Center
        'Mettre le nom du customer au lieu du numéro (bd NOMIS)
        uneCellule.Text = "Total"
        uneCellule.CssClass = "Gras"
        uneLigne.Cells.Add(uneCellule)

        AfficheNomCustomers = uneLigne
    End Function

    Function AfficheEntete(ByVal dtTableCustomersRegions() As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        Dim myrow As DataRow
        Dim strTitres() As String = {"Planned", "Completed", "Notes"}
        Dim i As Integer
        Dim index As Integer

        uneLigne.CssClass = "SousTitresRapport"

        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = "Division"
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = "Initiatives"
        uneLigne.Cells.Add(uneCellule)

        For index = 0 To UBound(dtTableCustomersRegions)
            For Each myrow In dtTableCustomersRegions(index).Rows
                For i = 0 To UBound(strTitres)
                    uneCellule = New TableCell
                    uneCellule.BorderWidth = Unit.Pixel(1)
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
            uneCellule.Text = strTitres(i)
            uneLigne.Cells.Add(uneCellule)
        Next

        AfficheEntete = uneLigne
    End Function

    Function AfficheDonnees(ByVal myRow As DataRow, ByVal dtTableCustomersRegions() As DataTable, ByVal dbConnSQLServer As OleDbConnection, ByVal Grouping As String) As TableRow
        Dim fields() As String = {"CustomerNo", "RegionNo"}
        Dim myRow1 As DataRow
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim i As Integer
        Dim i2 As Integer
        Dim totals(2) As Integer

        'Écrit le nom de la division et de l'initiative
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell

        uneLigne.CssClass = "TexteGRapport"

        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = IIf(Trim(myRow(1)) <> Trim(Grouping), myRow(1), "")
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = myRow(2)
        uneLigne.Cells.Add(uneCellule)

        For i = 0 To UBound(dtTableCustomersRegions)
            For Each myRow1 In dtTableCustomersRegions(i).Rows
                strReq = "Select Planned, Completed, Notes from " & dtTableCustomersRegions(i).TableName & " where FY='" & FiscalYear() & "' AND InitiativeNo=" & myRow(0) & " AND " & fields(i) & "=" & myRow1(0)
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
                Dim dtTableDonnees As New DataTable
                cmdTable.Fill(dtTableDonnees)
                Dim mycol As DataColumn
                i2 = 0
                For Each mycol In dtTableDonnees.Columns
                    uneCellule = New TableCell
                    uneCellule.BorderWidth = Unit.Pixel(1)
                    If dtTableDonnees.Rows.Count > 0 Then
                        If Not IsDBNull(dtTableDonnees.Rows(0)(mycol)) Then
                            uneCellule.Text = dtTableDonnees.Rows(0)(mycol)
                            totals(i2) += IIf(uneCellule.Text <> "", 1, 0)
                        End If
                        If mycol.ToString = "Notes" Then
                            uneCellule.HorizontalAlign = HorizontalAlign.Left
                        Else
                            uneCellule.HorizontalAlign = HorizontalAlign.Center
                        End If
                    End If
                    uneLigne.Cells.Add(uneCellule)
                    i2 += 1
                Next
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        For i = 0 To UBound(totals)
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

    Function WhereDivision() As String
        Dim strWhere As String = ""
        If Request.QueryString("Division") <> Nothing And Request.QueryString("Division") <> "" And UCase(Request.QueryString("Division")) <> "ALL" Then
            strWhere = " where Division='" & Request.QueryString("Division") & "'"
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

    Function FiscalYear() As Integer
        Dim year As Integer = Format(Now, "yyyy")

        If Now >= CDate("1-oct-" & Format(Now, "yyyy")) Then year += 1

        Return year
    End Function

End Class
