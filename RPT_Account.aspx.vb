Imports System.Data
Imports System.Data.OleDb

Public Class Account
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents title As System.Web.UI.WebControls.Label
    Protected WithEvents Report As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
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
        Dim Titre As Label = FindControl("title")
        Dim TableauRapport As PlaceHolder = FindControl("Report")
        Dim dtTableDivisions As New DataTable
        Dim dsTables As New DataSet
        Dim myRow As DataRow

        Dim TotalBookings As Integer

        Dim unTableau As New Table
        unTableau.CssClass = "EspaceTableau"

        Dim dbConnSQLServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSQLServer.Open()

        Dim strReq As String = "SELECT CustomerName FROM Nomis where CustomerNo = " & Request.QueryString("Customer")
        Dim cmdTable As New OleDbDataAdapter
        Dim dtTable As New DataTable

        strReq = "SELECT PrimaryCat FROM Nomis where FY='" & Request.QueryString("year1") & "' Group by PrimaryCat order by PrimaryCat"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        cmdTable.Fill(dtTableDivisions)

        If Request.QueryString("Type") = "A" Then
            'Title
            strReq = "SELECT CustomerName FROM Nomis where CustomerNo = " & Request.QueryString("Customer")
            cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
            cmdTable.Fill(dtTable)

            'Bookings Data
            strReq = "Select PrimaryCat, CAST(Sum(Bookings) AS NUMERIC) from NOMIS WHERE FY = '" & Request.QueryString("year1") & "' AND CustomerNo=" & Request.QueryString("Customer") & _
            IIf(Request.QueryString("OS") <> Nothing, " AND OSNo='" & Request.QueryString("OS") & "' ", "") & _
            " Group by CustomerNo, PrimaryCat"
            cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
            cmdTable.Fill(dsTables, "Bookings")

            'initiativesA Data
            strReq = "Select Division as PrimaryCat, Planned, Completed, Notes from Initiatives, InitiativesA where Initiatives.InitiativeNo = InitiativesA.InitiativeNo AND FY='" & Request.QueryString("year1") & "' and CustomerNo=" & Request.QueryString("Customer")
            cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
            cmdTable.Fill(dsTables, "Initiatives")
        ElseIf Request.QueryString("Type") = "B" Then
            'Title
            strReq = "SELECT Region FROM Regions where RegionNo = " & Request.QueryString("Customer")
            cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
            cmdTable.Fill(dtTable)

            'Bookings Data
            strReq = "Select PrimaryCat, CAST(Sum(Bookings) AS NUMERIC) from NOMIS, Customers WHERE Nomis.CustomerNo = Customers.CustomerNo AND FY = '" & Request.QueryString("year1") & "' AND RegionNo=" & Request.QueryString("Customer") & _
                       IIf(Request.QueryString("OS") <> Nothing, " AND NOMIS.OSNo='" & Request.QueryString("OS") & "' ", "") & _
                       " Group by PrimaryCat"
            Trace.Warn(strReq)
            cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
            cmdTable.Fill(dsTables, "Bookings")

            'initiativesB Data
            strReq = "Select Division as PrimaryCat, Planned, Completed, Notes from Initiatives, InitiativesB where Initiatives.InitiativeNo = InitiativesB.InitiativeNo AND FY='" & Request.QueryString("year1") & "' and RegionNo=" & Request.QueryString("Customer")
            cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
            cmdTable.Fill(dsTables, "Initiatives")


        End If

        Titre.Text = dtTable.Rows(0)(0)

        'Affichage l'entête
        unTableau.Controls.Add(AfficheEntete1())
        unTableau.Controls.Add(AfficheEntete2())

        For Each myRow In dtTableDivisions.Rows
            unTableau.Controls.Add(AfficheDonnees(myRow, dsTables, TotalBookings))
        Next

        unTableau.Controls.Add(AfficheTotauxColonnes(TotalBookings))
        TableauRapport.Controls.Add(unTableau)

        dbConnSQLServer.Close()

    End Sub

    Function AfficheEntete1() As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 2
        uneLigne.Cells.Add(uneCellule)
        Dim myRow As DataRow
        Dim i As Integer
        'Afficher les noms des customers

        uneLigne.CssClass = "EnteteRapport"

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.HorizontalAlign = HorizontalAlign.Center
        uneCellule.Text = "Bookings"
        uneLigne.Cells.Add(uneCellule)

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.ColumnSpan = 3
        uneCellule.HorizontalAlign = HorizontalAlign.Center
        uneCellule.Text = "Initiatives"
        uneLigne.Cells.Add(uneCellule)

        AfficheEntete1 = uneLigne
    End Function

    Function AfficheEntete2() As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        Dim myrow As DataRow
        Dim strTitres() As String = {"Division", "", "Actual", "Planned", "Completed", "Notes"}
        Dim i As Integer

        uneLigne.CssClass = "SousTitresRapport"

        For i = 0 To UBound(strTitres)
            uneCellule = New TableCell
            'uneCellule.Width = IIf(strTitres(i) = "", New Unit(20), New Unit(80))
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.Text = strTitres(i)
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneLigne.Cells.Add(uneCellule)
        Next

        AfficheEntete2 = uneLigne
    End Function

    Function AfficheDonnees(ByVal myRow As DataRow, ByVal dsTables As DataSet, ByRef TotalBookings As Integer) As TableRow

        'Écrit le nom de la division, du productCode et la description
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        Dim unTableau As Table
        Dim uneLigne2 As TableRow
        Dim uneCellule2 As TableCell

        uneLigne.CssClass = "TexteGRapport"

        'Affiche le nom de la division
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.VerticalAlign = VerticalAlign.Middle
        uneCellule.Text = myRow(0)
        uneLigne.Cells.Add(uneCellule)
        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneLigne.Cells.Add(uneCellule)

        Dim i As Integer
        Dim colonnes As Integer
        Dim resultats As Integer
        Dim goal As Double

        Dim dvTable As DataView

        For i = 0 To dsTables.Tables.Count - 1
            dvTable = New DataView(dsTables.Tables(i))
            dvTable.RowFilter = "PrimaryCat='" & Trim(myRow(0)) & "'"

            For colonnes = 1 To dvTable.Table.Columns.Count - 1
                uneCellule = New TableCell
                uneCellule.BorderWidth = Unit.Pixel(1)
                uneCellule.VerticalAlign = VerticalAlign.Middle
                If dvTable.Count > 0 Then
                    If i = 1 Then
                        unTableau = New Table
                        For resultats = 0 To dvTable.Count - 1
                            uneLigne2 = New TableRow
                            uneLigne2.CssClass = "TexteGRapport"
                            uneCellule2 = New TableCell
                            uneCellule2.Width = Unit.Percentage(100)
                            uneCellule.VerticalAlign = VerticalAlign.Top
                            If Not IsDBNull(dvTable(resultats)(colonnes)) Then
                                uneCellule2.Text = "- " & dvTable(resultats)(colonnes)
                            End If
                            uneLigne2.Cells.Add(uneCellule2)
                            unTableau.Controls.Add(uneLigne2)
                        Next
                        uneCellule.Controls.Add(unTableau)
                    ElseIf Not IsDBNull(dvTable(0)(colonnes)) Then
                        uneCellule.Text = dvTable(0)(colonnes)
                        TotalBookings += Val(uneCellule.Text)
                    End If
                End If
                uneCellule.HorizontalAlign = HorizontalAlign.Center
                uneLigne.Cells.Add(uneCellule)
            Next
        Next

        AfficheDonnees = uneLigne
    End Function

    Function AfficheTotauxColonnes(ByVal TotalBookings As Integer) As TableRow
        Dim myRow As New TableRow
        Dim uneCellule As TableCell
        Dim i As Integer

        uneCellule = New TableCell
        uneCellule.Text = "Total"
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.HorizontalAlign = HorizontalAlign.Left

        myRow.Cells.Add(uneCellule)

        uneCellule = New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        myRow.Cells.Add(uneCellule)

        uneCellule = New TableCell
        uneCellule.Text = TotalBookings
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.HorizontalAlign = HorizontalAlign.Center

        myRow.Cells.Add(uneCellule)



        AfficheTotauxColonnes = myRow

    End Function

    Function WhereDivision() As String
        Dim strWhere As String = ""
        If Request.QueryString("Division") <> Nothing And Request.QueryString("Division") <> "" And UCase(Request.QueryString("Division")) <> "ALL" Then
            strWhere = " where PrimaryCat='" & Request.QueryString("Division") & "'"
        End If
        WhereDivision = strWhere

    End Function

    Function FiscalYear() As Integer
        Dim year As Integer = Format(Now, "yyyy")

        If Now >= CDate("1-oct-" & Format(Now, "yyyy")) Then year += 1

        Return year
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
