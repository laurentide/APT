Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic

Public Class PrintInitiatives
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
        Dim strWhere As String = ""

        Dim NbDayYear As Integer = DateDiff(DateInterval.Day, CDate("01-oct-" & FiscalYear() - 1), CDate("01-oct-" & FiscalYear()))
        Dim DayOfFY As Integer = DateDiff(DateInterval.DayOfYear, CDate("01-oct-" & FiscalYear() - 1), Now) + 1

        If Request.QueryString("type") = "A" Then
            strWhere = "InitiativesA.CustomerNo = " & Request.QueryString("cusAB") & " AND OSNO='" & Request.QueryString("OS") & "' AND InitiativeNo = INI.InitiativeNo"
        Else
            strWhere = "InitiativesB.RegionNo = " & Request.QueryString("cusAB") & " AND InitiativeNo = INI.InitiativeNo"
        End If

        If Request.QueryString("type") = "A" Then
            strReq = "SELECT Division, " & _
                "Initiative, " & _
                "(Select Completed From InitiativesA WHERE " & strWhere & " AND FY='" & FiscalYear() & "') AS [Completed1], " & _
                "(Select Planned From InitiativesA WHERE " & strWhere & " AND FY='" & FiscalYear() & "') AS [Planned1], " & _
                "(Select Notes From InitiativesA WHERE " & strWhere & " AND FY='" & FiscalYear() & "') AS [Notes1], " & _
                "(Select Completed From InitiativesA WHERE " & strWhere & " AND FY='" & FiscalYear() + 1 & "') AS [Completed2], " & _
                "(Select Planned From InitiativesA WHERE " & strWhere & " AND FY='" & FiscalYear() + 1 & "')  AS [Planned2], " & _
                "(Select Notes From InitiativesA WHERE " & strWhere & " AND FY='" & FiscalYear() + 1 & "') AS [Notes2], " & _
                "INI.InitiativeNo FROM Initiatives INI"
        Else
            strReq = "SELECT Division, " & _
                "Initiative, " & _
                "(Select Completed From InitiativesB WHERE " & strWhere & " AND FY='" & FiscalYear() & "') AS [Completed1], " & _
                "(Select Planned From InitiativesB WHERE " & strWhere & " AND FY='" & FiscalYear() & "') AS [Planned1], " & _
                "(Select Notes From InitiativesB WHERE " & strWhere & " AND FY='" & FiscalYear() & "') AS [Notes1], " & _
                "(Select Completed From InitiativesB WHERE " & strWhere & " AND FY='" & FiscalYear() + 1 & "') AS [Completed2], " & _
                "(Select Planned From InitiativesB WHERE " & strWhere & " AND FY='" & FiscalYear() + 1 & "')  AS [Planned2], " & _
                "(Select Notes From InitiativesB WHERE " & strWhere & " AND FY='" & FiscalYear() + 1 & "') AS [Notes2], " & _
                "INI.InitiativeNo FROM Initiatives INI"
        End If

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        Dim dvTable As New DataView(dtTable)

        If UCase(Request.QueryString("division")) <> "ALL" Then
            dvTable.RowFilter = "Division='" & Request.QueryString("division") & "'"
        End If

        FillInitiativesTable(dvTable)

        dbConnSqlServer.Close()
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillInitiativesTable: Affiche les données dans le tableau initiatives                        |
    '|----------------------------------------------------------------------------------------------|
    Sub FillInitiativesTable(ByVal dvTable As DataView)
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
        Dim strCssClass As String

        table.CssClass = "EspaceTableau"

        intLargeurColonnes(0) = 120
        intLargeurColonnes(1) = 190
        intLargeurColonnes(2) = 70
        intLargeurColonnes(3) = 70
        intLargeurColonnes(4) = 180
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 180

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
                Else
                    newCell.HorizontalAlign = HorizontalAlign.Left
                End If

                newCell.CssClass = strCssClass
                newCell.BorderStyle = BorderStyle.Solid
                newRow.Cells.Add(newCell)
            Next
            table.Rows.Add(newRow)
        Next

        placeholder.Controls.Add(table)

    End Sub

    Function AfficheEntete() As TableRow
        Dim columnName As TableCell
        Dim header As New TableRow
        Dim i As Integer

        Dim strNomColonnes(7) As String

        strNomColonnes(0) = "Division"
        strNomColonnes(1) = "Initiative"
        strNomColonnes(2) = "FY" & Mid(FiscalYear, 3, 2) & " Completed"
        strNomColonnes(3) = "FY" & Mid(FiscalYear, 3, 2) & " Planned"
        strNomColonnes(4) = "FY" & Mid(FiscalYear, 3, 2) & " Notes"
        strNomColonnes(5) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Completed"
        strNomColonnes(6) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Planned"
        strNomColonnes(7) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Notes"

        Dim intLargeurColonnes(7) As Integer
        intLargeurColonnes(0) = 120
        intLargeurColonnes(1) = 190
        intLargeurColonnes(2) = 70
        intLargeurColonnes(3) = 70
        intLargeurColonnes(4) = 180
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 180

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
