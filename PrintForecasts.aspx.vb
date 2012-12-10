Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic

Public Class PrintForecasts
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

        Dim dbConnSqlServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSqlServer.Open()

        Dim strReq As String = ""
        Dim dtTable As New DataTable

        Dim NbDayYear As Integer = DateDiff(DateInterval.Day, CDate("01-oct-" & FiscalYear() - 1), CDate("01-oct-" & FiscalYear()))
        Dim DayOfFY As Integer = DateDiff(DateInterval.DayOfYear, CDate("01-oct-" & FiscalYear() - 1), Now) + 1
        If Now > CDate("1-oct-" & FiscalYear()) Then
            DayOfFY = NbDayYear
        End If

        If Request.QueryString("type") = "A" Then
            strReq = "Select PC, (Select ProductDesc from nomis where Nomis.PC = FA.PC group by ProductDesc) AS ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "(Select Forecast from ForecastA where PC=FA.PC AND CustomerNo=" & Request.QueryString("cusAB") & " AND FY='" & FiscalYear() & "') AS CurrentForecast, " & _
                        "(Select Forecast from ForecastA where PC=FA.PC AND CustomerNo=" & Request.QueryString("cusAB") & " AND FY='" & FiscalYear() + 1 & "') AS NextYearForecast, " & _
                        "(Select PrimaryCat from nomis where Nomis.PC = FA.PC group by PrimaryCat) AS Division " & _
                        "from ForecastA FA " & _
                        "WHERE CustomerNo=" & Request.QueryString("cusAB") & " and OsNo=" & Request.QueryString("OS") & _
                        " UNION " & _
                        "SELECT N1.PC, ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.CustomerNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.CustomerNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.CustomerNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.CustomerNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.CustomerNo = " & Request.QueryString("cusAB") & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.CustomerNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "0 AS CurrentForecast, " & _
                        "0 AS NextYearForecast, " & _
                        "PrimaryCat as Division  " & _
                        "FROM NOMIS N1 " & _
                        "WHERE NOT EXISTS (SELECT * From ForecastA where N1.PC = PC AND N1.CUSTOMERNO = CUSTOMERNO AND N1.OSNO = OSNO) " & _
                        "AND OSNO = '" & Request.QueryString("OS") & "' AND CUSTOMERNO=" & Request.QueryString("cusAB") & " AND PC is not null and PC <> '' " & _
                        "Group by N1.Pc, productDesc, PrimaryCat"
        Else
            strReq = "SELECT FB.PC, (Select ProductDesc from nomis where Nomis.PC = FB.PC group by ProductDesc) AS ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "(Select Forecast from ForecastB where PC=FB.PC AND RegionNo=" & Request.QueryString("cusAB") & " AND FY='" & FiscalYear() & "') AS CurrentForecast, " & _
                        "(Select Forecast from ForecastB where PC=FB.PC AND RegionNo=" & Request.QueryString("cusAB") & " AND FY='" & FiscalYear() + 1 & "') AS NextYearForecast, " & _
                        "(Select PrimaryCat from nomis where Nomis.PC = FB.PC group by PrimaryCat) AS Division " & _
                        "FROM ForecastB FB INNER JOIN " & _
                        "Regions ON FB.RegionNo = Regions.RegionNo INNER JOIN " & _
                        "Customers ON Regions.OsNo = Customers.OSNo " & _
                        "WHERE (dbo.Regions.RegionNo = " & Request.QueryString("cusAB") & ") " & _
                        "GROUP BY FB.PC, FB.Forecast, Regions.OsNo " & _
                        "UNION " & _
                        "SELECT N1.PC, ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Request.QueryString("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & Request.QueryString("cusAB") & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "0 AS CurrentForecast, " & _
                        "0 AS NextYearForecast, " & _
                        "PrimaryCat as Division " & _
                        "FROM dbo.Customers C1 INNER JOIN " & _
                        "dbo.NOMIS N1 ON C1.CustomerNo = N1.CUSTOMERNO AND C1.OSNo = N1.OSNO " & _
                        "WHERE (C1.RegionNo = " & Request.QueryString("cusAB") & ") AND " & _
                        "((SELECT COUNT(PC) FROM ForecastB WHERE PC = N1.PC AND RegionNo = C1.RegionNo) = 0) " & _
                        " AND N1.PC IS NOT NULL AND N1.PC <> '' " & _
                        "GROUP BY N1.PC, ProductDesc, PrimaryCat"
        End If

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        Dim dvTable As New DataView(dtTable)

        If UCase(Request.QueryString("division")) <> "ALL" Then
            dvTable.RowFilter = "Division='" & Request.QueryString("division") & "'"
        End If

        FillForecastTable(dvTable)
        dbConnSqlServer.Close()

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillForecastTable: Affiche les données dans le tableau forecasts                             |
    '|----------------------------------------------------------------------------------------------|
    Sub FillForecastTable(ByVal dvTable As DataView)
        Dim x As Integer
        Dim y As Integer
        Dim intLargeurColonnes(8) As Integer
        Dim myRow As DataRow
        Dim newRow As TableRow


        Dim placeholder As PlaceHolder = FindControl("Report")
        Dim table As New Table
        Dim newCell As TableCell
        Dim field As TextBox
        Dim strInsert As String
        Dim Totals(6) As Integer
        Dim strCssClass As String

        table.CssClass = "EspaceTableau"

        intLargeurColonnes(0) = 10
        intLargeurColonnes(1) = 150
        intLargeurColonnes(2) = 50
        intLargeurColonnes(3) = 50
        intLargeurColonnes(4) = 50
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 70
        intLargeurColonnes(8) = 70

        'Afficher l'entete
        table.Rows.Add(AfficheEntete())

        For x = 0 To dvTable.Count - 1
            newRow = New TableRow
            newRow.CssClass = "TexteGRapport"

            For y = 0 To dvTable.Table.Columns.Count - 2
                newCell = New TableCell
                strCssClass = "BordureTableau FieldForecast texte"
                newCell.Width = Unit.Pixel(intLargeurColonnes(y))

                newCell.Text = " "
                If Not IsDBNull(dvTable.Item(x)(y)) Then
                    newCell.Text = dvTable.Item(x)(y)
                End If

                If y >= 2 Then
                    newCell.HorizontalAlign = HorizontalAlign.Center
                    Totals(y - 2) += Val(newCell.Text)
                Else
                    newCell.HorizontalAlign = HorizontalAlign.Left
                End If

                newCell.CssClass = strCssClass
                newCell.BorderStyle = BorderStyle.Solid
                newRow.Cells.Add(newCell)
            Next
            table.Rows.Add(newRow)
        Next

        table.Rows.Add(FillForecastTotal(Totals))
        placeholder.Controls.Add(table)

    End Sub

    Function AfficheEntete() As TableRow
        Dim columnName As TableCell
        Dim header As New TableRow
        Dim i As Integer

        Dim strNomColonnes(8) As String

        strNomColonnes(0) = "PC"
        strNomColonnes(1) = "PC - Description"
        strNomColonnes(2) = "FY" & Mid(FiscalYear() - 2, 3, 2)
        strNomColonnes(3) = "FY" & Mid(FiscalYear() - 1, 3, 2)
        strNomColonnes(4) = "FY" & Mid(FiscalYear(), 3, 2)
        strNomColonnes(5) = "FY" & Mid(FiscalYear(), 3, 2) & " Prediction"
        strNomColonnes(6) = "FY" & Mid(FiscalYear() - 2, 3, 2) & "- FY" & Mid(FiscalYear(), 3, 2) & " Avg."
        strNomColonnes(7) = "FY" & Mid(FiscalYear(), 3, 2) & " Forecast"
        strNomColonnes(8) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Forecast"

        Dim intLargeurColonnes(8) As Integer
        intLargeurColonnes(0) = 10
        intLargeurColonnes(1) = 150
        intLargeurColonnes(2) = 50
        intLargeurColonnes(3) = 50
        intLargeurColonnes(4) = 50
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 70
        intLargeurColonnes(8) = 70

        For i = 0 To UBound(strNomColonnes)
            columnName = New TableCell
            columnName.Text = strNomColonnes(i)
            columnName.Width = New Unit(intLargeurColonnes(i))
            columnName.CssClass = "BordureTableau"
            columnName.BorderStyle = BorderStyle.Solid
            If i >= 2 Then
                columnName.HorizontalAlign = HorizontalAlign.Center
            End If
            header.Cells.Add(columnName)
        Next

        AfficheEntete = header


    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillForecastTotal: Affiche les sommes au bas du tableau Forecasts                            |
    '|----------------------------------------------------------------------------------------------|
    Function FillForecastTotal(ByVal Totals() As Integer) As TableRow
        Dim Total As New TableRow
        Dim ColumnName As TableCell

        Dim intLargeurColonnes(7) As Integer
        Dim i As Integer

        Total.CssClass = "TexteGRapport"

        intLargeurColonnes(0) = 50
        intLargeurColonnes(1) = 50
        intLargeurColonnes(2) = 50
        intLargeurColonnes(3) = 70
        intLargeurColonnes(4) = 70
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70

        ColumnName = New TableCell
        ColumnName.Text = "Total:"
        ColumnName.Width = New Unit(170)
        ColumnName.ColumnSpan = 2
        ColumnName.CssClass = "BordureTableau FieldForecast texte "
        ColumnName.BorderStyle = BorderStyle.Solid
        ColumnName.HorizontalAlign = HorizontalAlign.Left

        Total.Cells.Add(ColumnName)

        For i = 0 To UBound(Totals)
            ColumnName = New TableCell
            ColumnName.Text = Totals(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureTableau FieldForecast Centre texte "
            ColumnName.BorderStyle = BorderStyle.Solid
            ColumnName.HorizontalAlign = HorizontalAlign.Center
            Total.Cells.Add(ColumnName)
        Next

        ColumnName = New TableCell
        ColumnName.Width = New Unit(intLargeurColonnes(7))
        ColumnName.CssClass = "FieldForecast Centre texte"

        ColumnName.BorderStyle = BorderStyle.Solid

        FillForecastTotal = Total

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

    '
    '|----------------------------------------------------------------------------------------------|
    '| FiscalYear: Retourne l'année fiscale                                                         |
    '|----------------------------------------------------------------------------------------------|
    Function FiscalYear() As Integer
        Dim year As Integer = Format(Now, "yyyy")
        Return year
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

End Class
