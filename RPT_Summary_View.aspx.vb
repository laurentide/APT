Imports System.Data
Imports System.Data.OleDb

Public Class RPT_Summary_View
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Report As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents TabReport As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox

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
        TextBox1.Text = Now
        CreateReport()

    End Sub

    Sub CreateReport()
        Dim TableauRapport As PlaceHolder = FindControl("Report")
        Dim dtTableDivisions As New DataTable
        Dim dtTableOS As DataTable
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

        Dim strReq As String = "SELECT PrimaryCat FROM PCNOMIS " & WhereDivision() & " Group by PrimaryCat order by PrimaryCat"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSQLServer)
        cmdTable.Fill(dtTableDivisions)

        strReq = "Select Employee.OSNo,OsName from NOMIS, Employee where Employee.OSNO = NOMIS.OSNO AND (dbo.NOMIS.FY = '" & Request.QueryString("year1") & "')" & WhereOs("Employee") & " Group by Employee.OSNo, OSName order by OSName"

        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableOS = New DataTable
        cmdTable.Fill(dtTableOS)

        'Ligne contenant les nom des customers
        unTableau.Controls.Add(AfficheNomOS(dtTableOS))

        'Réserve le nombre de colonnes pour les totaux en bas de page
        intNbElements = (dtTableOS.Rows.Count * 2) - 1
        ReDim TotalColonnes(intNbElements)

        'Affichage l'entête entre les noms des customers et les données
        unTableau.Controls.Add(AfficheEntete(dtTableOS))

        strReq = "SELECT ISNULL(NOMIS.PRIMARYCAT, Goals.Division) AS PRIMARYCAT, " & _
                    "ISNULL(NOMIS.OSNO, GOALS.OSNO) AS OSNo, CAST(Goals.Goal AS NUMERIC), CAST(SUM(dbo.NOMIS.BOOKINGS) AS NUMERIC) " & _
                    "FROM NOMIS FULL OUTER JOIN Goals ON NOMIS.FY = Goals.FY AND " & _
                    "NOMIS.OSNO = Goals.OsNo AND NOMIS.PRIMARYCAT = Goals.Division " & _
                    "WHERE 1=1 " & _
                    IIf(Request.QueryString("Customer") <> Nothing And Request.QueryString("Type") = "A", " AND CustomerNo=" & Request.QueryString("Customer"), "") & _
                    IIf(Request.QueryString("Customer") <> Nothing And Request.QueryString("Type") = "B", " AND CustomerNo IN (Select CustomerNo from Customers where RegionNo=" & Request.QueryString("Customer") & ") ", " ") & _
                    "GROUP BY ISNULL(NOMIS.FY, Goals.FY), ISNULL(NOMIS.PRIMARYCAT, Goals.Division) , ISNULL(NOMIS.OSNO, GOALS.OSNO), dbo.Goals.Goal " & _
                    "HAVING (ISNULL(NOMIS.FY, Goals.FY) = '" & Request.QueryString("Year1") & "') " & _
                    IIf(Request.QueryString("OS") <> Nothing, " AND ISNULL(NOMIS.OSNO, GOALS.OSNO) = '" & Request.QueryString("OS") & "' ", "")
                    
        Trace.Warn(strReq)
        cmdTable = New OleDbDataAdapter(strReq, dbConnSQLServer)
        dtTableDonnees = New DataTable
        cmdTable.Fill(dtTableDonnees)

        For Each myRow In dtTableDivisions.Rows
            unTableau.Controls.Add(AfficheDonnees(myRow, dtTableOS, dtTableDonnees, TotalColonnes))
        Next

        unTableau.Controls.Add(AfficheTotauxColonnes(TotalColonnes))
        TableauRapport.Controls.Add(unTableau)

        dbConnSQLServer.Close()

    End Sub

    Function AfficheNomOS(ByVal dtTableOS As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneLigne.Cells.Add(uneCellule)
        Dim myRow As DataRow
        Dim i As Integer
        'Afficher les noms des customers

        uneLigne.CssClass = "EnteteRapport"

        For Each myRow In dtTableOS.Rows
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.ColumnSpan = 2
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneCellule.Text = myRow(1).ToString()
            uneLigne.Cells.Add(uneCellule)
            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneLigne.Cells.Add(uneCellule)
        Next

        AfficheNomOS = uneLigne
    End Function

    Function AfficheEntete(ByVal dtTableOS As DataTable) As TableRow
        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell
        Dim myrow As DataRow
        Dim strTitres() As String = {"YTD Goal", "YTD Actual"}
        Dim i As Integer
        Dim index As Integer

        uneLigne.CssClass = "SousTitresRapport"

        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = "Division"
        uneCellule.Width = New Unit(100)
        uneLigne.Cells.Add(uneCellule)

        For Each myrow In dtTableOS.Rows
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

        AfficheEntete = uneLigne
    End Function

    Function AfficheDonnees(ByVal myRow As DataRow, ByVal dtTableOS As DataTable, ByVal dtTableDonnees As DataTable, ByRef TotalColonnes() As Integer) As TableRow
        Dim dvDonnees As DataView

        Dim fields() As String = {"CustomerNo", "RegionNo"}
        Dim myRow1 As DataRow
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim i As Integer
        Dim indice As Integer

        Dim uneLigne As New TableRow
        Dim uneCellule As New TableCell

        Dim Goals As Integer
        Dim Actual As Integer

        'Remplit les données selon le customer et l'initiative
        Dim NbDayYear As Integer = DateDiff(DateInterval.Day, CDate("01-oct-" & FiscalYear() - 1), CDate("01-oct-" & FiscalYear()))
        Dim DayOfFY As Integer = DateDiff(DateInterval.DayOfYear, CDate("01-oct-" & FiscalYear() - 1), Now) + 1

        uneLigne.CssClass = "TexteGRapport"

        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.Text = Trim(myRow(0))
        uneLigne.Cells.Add(uneCellule)

        dvDonnees = New DataView(dtTableDonnees)
        For Each myRow1 In dtTableOS.Rows
            dvDonnees.RowFilter = "OSNO='" & myRow1(0) & "' AND PRIMARYCAT='" & myRow(0) & "'"
            Goals = 0
            Actual = 0
            For i = 0 To dvDonnees.Count - 1
                If Not IsDBNull(dvDonnees(i)(2)) Then
                    Goals += dvDonnees(i)(2)
                End If
                If Not IsDBNull(dvDonnees(i)(3)) Then
                    Actual += dvDonnees(i)(3)
                End If
            Next

            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.HorizontalAlign = HorizontalAlign.Center

            uneCellule.Text = Math.Round((Goals / NbDayYear) * DayOfFY, 2)
            uneLigne.Cells.Add(uneCellule)

            TotalColonnes(indice) += Val(uneCellule.Text)
            indice += 1

            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneCellule.HorizontalAlign = HorizontalAlign.Center
            uneCellule.Text = Actual
            uneLigne.Cells.Add(uneCellule)

            TotalColonnes(indice) += Val(uneCellule.Text)
            indice += 1

            uneCellule = New TableCell
            uneCellule.BorderWidth = Unit.Pixel(1)
            uneLigne.Cells.Add(uneCellule)
        Next

        AfficheDonnees = uneLigne
    End Function

    Function AfficheTotauxColonnes(ByVal TotalColonnes() As Integer) As TableRow
        Dim myRow As New TableRow
        Dim uneCellule As TableCell
        Dim i As Integer

        uneCellule = New TableCell
        uneCellule.Text = "Total"
        uneCellule.BorderWidth = Unit.Pixel(1)
        uneCellule.HorizontalAlign = HorizontalAlign.Left

        myRow.Cells.Add(uneCellule)

        For i = 0 To UBound(TotalColonnes)
            If i > 0 And i Mod 2 = 0 Then
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
            If Request.QueryString("OS") <> "" Then
                strWhere = " AND " & strTable & ".OSNo='" & IIf(Request.QueryString("OS").Length = 2, "0", "") & Request.QueryString("OS") & "'"
            End If
        End If
        WhereOs = strWhere
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

    '|----------------------------------------------------------------------------------------------|
    '| ExportToExcel: Export Report to Excel												        |
    '|----------------------------------------------------------------------------------------------|

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

    '|----------------------------------------------------------------------------------------------|
    '| EtablitConnexionSQLServer												                    |
    '|----------------------------------------------------------------------------------------------|
    Function EtablitConnexionSQLServer() As OleDbConnection
        Dim strConn As String = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=APT;" & _
                                "User ID=APT;Password=APT"
        EtablitConnexionSQLServer = New OleDbConnection(strConn)
    End Function


End Class
