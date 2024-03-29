Imports System
imports System.Data
Imports System.Data.OleDb
imports System.Web.UI
imports System.Web.UI.WebControls
imports Microsoft.VisualBasic

Public class APT
    inherits System.Web.UI.Page
    Protected WithEvents lblSalesman As System.Web.UI.WebControls.Label
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button
    Protected WithEvents CustomerCity As System.Web.UI.WebControls.Label
    Protected WithEvents CustomerNo As System.Web.UI.WebControls.Label
    Protected WithEvents DGForecasts As System.Web.UI.WebControls.Table
    Protected WithEvents WINNAME As System.Web.UI.WebControls.TextBox
    Protected WithEvents OSNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents RPT1 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents RPT2 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents RPT3 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents RPT4 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents RPT5 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents RPT6 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents rbCustomer As System.Web.UI.WebControls.RadioButton
    Protected WithEvents rbRegion As System.Web.UI.WebControls.RadioButton
    Protected WithEvents DGForecastsHeader As System.Web.UI.WebControls.Table
    Protected WithEvents ForecastChanged As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents dgCustomersABHeader As System.Web.UI.WebControls.Table
    Protected WithEvents dgCustomersAB As System.Web.UI.WebControls.Table
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents Description As System.Web.UI.WebControls.Label
    Protected WithEvents Forecast As System.Web.UI.WebControls.TextBox
    Protected WithEvents SalesmanPerform As System.Web.UI.WebControls.Button
    Protected WithEvents Initiatives As System.Web.UI.WebControls.Button
    Protected WithEvents DGInitiativesHeader As System.Web.UI.WebControls.Table
    Protected WithEvents Table2 As System.Web.UI.WebControls.Table
    Protected WithEvents InitiativesChanged As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Products As System.Web.UI.WebControls.Button
    Protected WithEvents DGInitiatives As System.Web.UI.WebControls.Table
    Protected WithEvents InitiativeChanged As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents DGGoalsHeader As System.Web.UI.WebControls.Table
    Protected WithEvents DGGoals As System.Web.UI.WebControls.Table
    Protected WithEvents GoalsChanged As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents DGForecastsTotal As System.Web.UI.WebControls.Table
    Protected WithEvents DGGoalsTotal As System.Web.UI.WebControls.Table
    Protected WithEvents ExtraDetails As System.Web.UI.WebControls.Button
    Protected WithEvents DGDetailsHeader As System.Web.UI.WebControls.Table
    Protected WithEvents DGDetails As System.Web.UI.WebControls.Table
    Protected WithEvents DGDetailsTotal As System.Web.UI.WebControls.Table
    Protected WithEvents RPT7 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button

    Private Sub InitializeComponent()

    End Sub
   
    '
    '|----------------------------------------------------------------------------------------------|
    '| EtablitConnexionSQLServer												                    |
    '|----------------------------------------------------------------------------------------------|
    Function EtablitConnexionSQLServer() As OleDbConnection
        Dim strConn As String = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=APT;" & _
                                "User ID=APT;Password=APT"
        EtablitConnexionSQLServer = New OleDbConnection(strConn)
    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| GetSessionOs: Retrouver le No OS si l'usager y a acc�s                                       |
    '|----------------------------------------------------------------------------------------------|
    Sub GetSessionOs(ByVal dbConnSqlServer As OleDbConnection)
        If Session("adminOS") <> Nothing Then
            OSExists(dbConnSqlServer, Session("adminOS"), )
        ElseIf Session("adminNoOS") <> Nothing Then
            OSExists(dbConnSqlServer, , Session("adminNoOS").ToString())
        Else
            Dim strTemp() As String = Split(User.Identity.Name, "\")
            OSExists(dbConnSqlServer, strTemp(1), )
        End If
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| OSExists: V�rifie l'existence de l'usager dans la base de donn�es                            |
    '|----------------------------------------------------------------------------------------------|
    Sub OSExists(ByVal dbConn As OleDbConnection, Optional ByVal strWinName As String = Nothing, Optional ByVal strOSNo As String = Nothing)

        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable

        If strWinName <> Nothing Then
            strReq = "Select Cast(OSNO as VARCHAR) from Employee where WINDOWSNAME= '" & strWinName & "'"
        ElseIf strOSNo <> Nothing Then
            strReq = "Select Cast(OSNO as VARCHAR) from Employee where OSNO= '" & strOSNo.ToString() & "'"
        End If

        cmdTable = New OleDbDataAdapter(strReq, dbConn)
        cmdTable.Fill(dtTable)

        If dtTable.Rows.Count > 0 Then
            Session("OS") = dtTable.Rows(0)(0).ToString()
        Else
            Session("OS") = Nothing
        End If

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| RedirectMenu: Section admin - Redirige vers le menu principal                                |
    '|----------------------------------------------------------------------------------------------|
    Sub RedirectMenu(ByVal sender As Object, ByVal e As EventArgs)
        If Not Request.Form("WINNAME") Is Nothing Then
            Session("adminOS") = Request.Form("WINNAME").ToString()
            Session("adminNoOS") = Request.Form("OSNo").ToString()
        End If

        Response.Redirect("Menu.aspx")
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| AfficheCustomers: Remplit les informations des customers                                     |
    '|----------------------------------------------------------------------------------------------|
    Sub AfficheCustomers(ByVal dbConnSqlServer As OleDbConnection)
        ' Create a DataSet with one table, two columns, and ten rows.

        Dim Entete() As String = {"Customer #", "Customer", "City", "A", "B", "Region of B's"}
        Dim ds As New DataSet("myDataSet")
        Dim t As New DataTable
        Dim cmdTable As OleDbDataAdapter

        ' Data for Customers
        Dim strReq As String = "Select CAST(CUSTOMERNO AS VARCHAR(10)) AS CUSTOMERNO, CustomerName, City" & _
                                 " from NOMIS" & _
                                 " Where CustomerNo IS not NULL AND FY IS NOT NULL " & _
                                 " AND OSNO='" & Session("OS").ToString() & _
                                 "' Group by CUSTOMERNO, CustomerName, City order By CustomerName"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        t = New DataTable("Customers")
        cmdTable.Fill(t)

        ' Add tables to the DataSet
        ds.Tables.Add(t)

        Dim dsTables As New DataSet
        Dim dtTable As New DataTable

        strReq = "Select CAST(CustomerNo AS VARCHAR(10)) AS CUSTOMERNO, CASE WHEN AB = 'A' THEN 1 ELSE 0 END AS A, CASE WHEN AB = 'B' THEN 1 ELSE 0 END  AS B, RegionNo from Customers where OSNo='" & Session("OS").ToString() & "'"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)
        Dim mycol As DataColumn
        Dim myCol1 As DataColumn
        Dim myRow As DataRow

        ' DataColumn array to set primary key.
        Dim keyCol(1) As DataColumn

        ' Set primary key column.
        keyCol(0) = t.Columns(0)
        t.PrimaryKey = keyCol

        ' Accept changes.
        ds.AcceptChanges()

        ' Create a second DataTable identical to the first
        ' with one extra column using the Clone method.
        Dim t2 As New DataTable
        t2 = t.Clone()

        ' Add column.
        For Each mycol In dtTable.Columns
            If ds.Tables("Customers").Columns.IndexOf(mycol.ToString()) < 0 Then
                t2.Columns.Add(mycol.ToString(), Type.GetType("System.String"))
            End If
        Next

        Dim newRow As DataRow
        mycol = New DataColumn
        For Each myRow In dtTable.Rows
            newRow = t2.NewRow()
            For Each myCol1 In dtTable.Columns
                newRow(myCol1.ToString()) = myRow(myCol1)
            Next
            For Each mycol In ds.Tables("Customers").Columns
                Dim temp As DataRow = ds.Tables("Customers").Rows.Find(myRow("CustomerNo"))
                If Not temp Is Nothing Then
                    newRow(mycol.ToString()) = temp(mycol)
                End If
            Next

            t2.Rows.Add(newRow)
        Next

        ' Merge the table into the DataSet.
        ds.Merge(t2, False, MissingSchemaAction.Add)

        strReq = "Select 'NULL' AS RegionNo, '' AS Region"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(ds, "Region")
        strReq = "Select RegionNo, Region from Regions where OSNo='" & Session("OS").ToString() & "'"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(ds, "Region")

        FillCustomersTable(ds)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillCustomersTable: Affiche le tableau des CustomersAB										|
    '|----------------------------------------------------------------------------------------------|
    Sub FillCustomersTable(ByVal dsTables As DataSet)
        Dim dgCustomersAB As Table = FindControl("dgCustomersAB")

        Dim x As Integer
        Dim y As Integer
        Dim intLargeurColonnes(6) As Integer
        Dim myRow As DataRow
        Dim newRow As TableRow
        Dim newCell As TableCell
        Dim chkfield As CheckBox
        Dim test As RadioButton
        Dim ddlfield As DropDownList
        Dim strInsert As String

        intLargeurColonnes(0) = 80
        intLargeurColonnes(1) = 250
        intLargeurColonnes(2) = 150
        intLargeurColonnes(3) = 20
        intLargeurColonnes(4) = 20
        intLargeurColonnes(5) = 180

        For x = 0 To dsTables.Tables("Customers").Rows.Count - 1
            newRow = New TableRow
            For y = 0 To dsTables.Tables("Customers").Columns.Count - 1
                newCell = New TableCell
                newCell.CssClass = "BordureTableau EspacesCustomers"
                newCell.Width = Unit.Pixel(intLargeurColonnes(y))
                If y < 3 Then ' Information
                    If Not IsDBNull(dsTables.Tables("Customers").Rows(x)(y)) Then
                        newCell.Text = dsTables.Tables("Customers").Rows(x)(y)
                    End If
                ElseIf y < 5 Then 'Checkboxes
                    chkfield = New CheckBox
                    chkfield.ID = dsTables.Tables("Customers").Columns(y).ToString & "_" & dsTables.Tables("Customers").Rows(x)(0)
                    chkfield.Attributes.Add("onclick", "javascript:CheckUncheck();")
                    If Request.Form(chkfield.ID) = Nothing Then
                        If Not IsDBNull(dsTables.Tables("Customers").Rows(x)(y)) Then
                            chkfield.Checked = dsTables.Tables("Customers").Rows(x)(y)
                        Else
                            chkfield.Checked = False
                        End If
                    End If
                    chkfield.CssClass = "TextEntry"
                    newCell.Controls.Add(chkfield)
                Else 'DropDownList
                    ddlfield = New DropDownList
                    ddlfield.DataSource = dsTables.Tables("Region")
                    ddlfield.DataTextField = "Region"
                    ddlfield.DataValueField = "RegionNo"
                    ddlfield.DataBind()
                    If Not IsDBNull(dsTables.Tables("Customers").Rows(x)(y)) Then
                        ddlfield.SelectedValue = dsTables.Tables("Customers").Rows(x)(y)
                    End If
                    ddlfield.CssClass = "TextEntry DDLRegion"
                    newCell.Controls.Add(ddlfield)
                End If
                newRow.Cells.Add(newCell)

            Next
            dgCustomersAB.Rows.Add(newRow)
        Next

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FiscalYear: Retourne l'ann�e fiscale                                                         |
    '|----------------------------------------------------------------------------------------------|
    Function FiscalYear() As Integer
        Dim year As Integer = Format(Now, "yyyy")
        Return year
    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| SaveCustomersAB: Enregistre les modifications effectu�s aux Customers                        |
    '|----------------------------------------------------------------------------------------------|
    Sub SaveCustomersAB(ByVal sender As Object, ByVal e As EventArgs)
        'Save in database
        Dim DGCustomersAB As Table = FindControl("dgCustomersAB")
        Dim CustomerRow As TableRow
        Dim chkBoxA As CheckBox
        Dim chkBoxB As CheckBox
        Dim NoCustomer As String
        Dim ddl As DropDownList
        Dim strCommand As String

        Dim dbConnSqlServer As OleDbConnection = EtablitConnexionSQLServer()
        Dim dbCommand As New OleDbCommand
        dbConnSqlServer.Open()
        dbCommand.Connection = dbConnSqlServer

        'Ligne SQL - FG
        strCommand = "Delete from Customers where OSNo='" & Session("OS") & "'"
        dbCommand.CommandText = strCommand
        dbCommand.ExecuteNonQuery()

        For Each CustomerRow In DGCustomersAB.Rows
            strCommand = "Insert into Customers Values("

            chkBoxA = CustomerRow.Cells(3).Controls(0)
            chkBoxB = CustomerRow.Cells(4).Controls(0)
            NoCustomer = CustomerRow.Cells(0).Text
            ddl = CustomerRow.Cells(5).Controls(0)

            strCommand += NoCustomer & ","
            strCommand += "'" & Session("OS") & "',"

            If chkBoxA.Checked Then
                strCommand += "'A',"
            ElseIf chkBoxB.Checked Then
                strCommand += "'B',"
            Else
                strCommand += "NULL,"
            End If

            strCommand += ddl.SelectedValue

            strCommand += ")"

            dbCommand.CommandText = strCommand
            dbCommand.ExecuteNonQuery()
        Next

        'Back to menu
        dbConnSqlServer.Close()

        Response.Redirect("Menu.aspx")
    End Sub

    Sub OnlySaveForecasts(ByVal sender As Object, ByVal e As EventArgs)

        SaveForecasts()

    End Sub

    Sub SaveAndCloseForecasts(ByVal sender As Object, ByVal e As EventArgs)
        SaveForecasts()
        Response.Redirect("Menu.aspx")
    End Sub
    '
    '|----------------------------------------------------------------------------------------------|
    '| SaveForecasts: Enregistre la section Forecast                                                |
    '|(forecastsA, forecastsB,InitiativesA, InitiativesB)                                           |
    '|----------------------------------------------------------------------------------------------|
    Sub SaveForecasts()
        'Save in database
        Dim dbConnSqlServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSqlServer.Open()

        SaveForecastEntry(dbConnSqlServer)
        SaveInitiativeEntry(dbConnSqlServer)

        dbConnSqlServer.Close()
        'Back to menu

        ' Session("CustomerSelected") = ""
        ' Session("RegionSelected") = ""
        Session("ForecastChangedA") = Nothing
        Session("ForecastChangedB") = Nothing
        Session("InitiativeChangedA") = Nothing
        Session("InitiativeChangedB") = Nothing

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| SaveForecastEntry: Enregistre les forecasts modifi�s                                         |
    '|----------------------------------------------------------------------------------------------|
    Sub SaveForecastEntry(ByVal dbConnSqlServer As OleDbConnection)
        Dim Forecast() As String
        Dim Values() As String
        Dim strReq As String
        Dim strWhere As String
        Dim strCommand As String = ""

        Dim i As Integer

        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim cmdCommand As New OleDbCommand

        cmdCommand.Connection = dbConnSqlServer

        If Session("ForecastChangedA") <> Nothing Then
            Forecast = Session("ForecastChangedA").split(";")
            For i = 0 To UBound(Forecast) - 1

                Values = Forecast(i).Split(",")
                strWhere = " WHERE CustomerNo= '" & Mid(Values(0), 2) & "'"
                strWhere += " AND OSNO=" & Values(1)
                strWhere += " AND FY=" & Values(2)
                strWhere += " AND PC=" & Values(3)

                'Voir s'il existe dans la BD
                strReq = "Select * from ForecastA" & strWhere
                cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
                dtTable = New DataTable
                cmdTable.Fill(dtTable)

                If dtTable.Rows.Count > 0 Then
                    'S'il existe, fait un update
                    strCommand = "UPDATE ForecastA Set Forecast= '" & Mid(Values(4), 2, Values(4).Length - 2) & "'" & strWhere
                Else
                    'S'il n'existe pas, l'ajouter
                    strCommand = "Insert into ForecastA Values" & Forecast(i)
                End If

                Try
                    cmdCommand.CommandText = strCommand
                    cmdCommand.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            Next
        End If

        If Session("ForecastChangedB") <> Nothing Then
            Forecast = Session("ForecastChangedB").split(";")
            For i = 0 To UBound(Forecast) - 1

                Values = Forecast(i).Split(",")
                strWhere = " WHERE RegionNo= '" & Mid(Values(0), 2) & "'"
                strWhere += " AND FY=" & Values(1)
                strWhere += " AND PC=" & Values(2)

                'Voir s'il existe dans la BD
                strReq = "Select * from ForecastB" & strWhere
                cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
                dtTable = New DataTable
                cmdTable.Fill(dtTable)

                If dtTable.Rows.Count > 0 Then
                    'S'il existe, fait un update
                    strCommand = "UPDATE ForecastB Set Forecast= '" & Mid(Values(3), 2, Values(3).Length - 2) & "'" & strWhere
                Else
                    'S'il n'existe pas, l'ajouter
                    strCommand = "Insert into ForecastB Values" & Forecast(i)
                End If

                Try
                    cmdCommand.CommandText = strCommand
                    cmdCommand.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            Next
        End If
    End Sub

    '|----------------------------------------------------------------------------------------------|
    '| SaveInitiativeEntry: Enregistre les initiatives modifi�es                                    |
    '|----------------------------------------------------------------------------------------------|
    Sub SaveInitiativeEntry(ByVal dbConnSqlServer As OleDbConnection)
        Dim Initiative() As String
        Dim Values() As String
        Dim strReq As String
        Dim strWhere As String
        Dim strCommand As String = ""

        Dim i As Integer

        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim cmdCommand As New OleDbCommand

        cmdCommand.Connection = dbConnSqlServer

        If Session("InitiativeChangedA") <> Nothing Then
            Initiative = Session("InitiativeChangedA").split(";")
            For i = 0 To UBound(Initiative) - 1

                Values = Initiative(i).Split(",")
                strWhere = " WHERE CustomerNo=" & Mid(Values(0), 2)
                strWhere += " AND OSNO=" & Values(1)
                strWhere += " AND FY=" & Values(2)
                strWhere += " AND InitiativeNo=" & Values(3)

                'Voir s'il existe dans la BD
                strReq = "Select * from InitiativesA" & strWhere
                cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
                dtTable = New DataTable
                cmdTable.Fill(dtTable)

                If dtTable.Rows.Count > 0 Then
                    'S'il existe, fait un update
                    strCommand = "UPDATE InitiativesA Set " & _
                                    "Completed = " & Mid(Values(4), 1, 300) & "," & _
                                    "Planned = " & Mid(Values(5), 1, 300) & "," & _
                                    "Notes = " & Mid(Values(6), 1, IIf(Values(6).Length <= 300, Values(6).Length - 1, 300)) & _
                                    strWhere
                Else
                    'S'il n'existe pas, l'ajouter
                    strCommand = "Insert into InitiativesA Values" & Initiative(i)
                End If

                Try
                    cmdCommand.CommandText = strCommand
                    cmdCommand.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            Next
        End If

        If Session("InitiativeChangedB") <> Nothing Then
            Initiative = Session("InitiativeChangedB").split(";")
            For i = 0 To UBound(Initiative) - 1

                Values = Initiative(i).Split(",")
                strWhere = " WHERE RegionNo=" & Mid(Values(0), 2)
                strWhere += " AND FY=" & Values(1)
                strWhere += " AND InitiativeNo=" & Values(2)

                'Voir s'il existe dans la BD
                strReq = "Select * from InitiativesB" & strWhere
                cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
                dtTable = New DataTable
                cmdTable.Fill(dtTable)

                If dtTable.Rows.Count > 0 Then
                    'S'il existe, fait un update
                    strCommand = "UPDATE InitiativesB Set " & _
                                    "Completed = " & Mid(Values(3), 1, 300) & "," & _
                                    "Planned = " & Mid(Values(4), 1, 300) & "," & _
                                    "Notes = " & Mid(Values(5), 1, IIf(Values(5).Length <= 300, Values(5).Length - 1, 300)) & _
                                    strWhere
                Else
                    'S'il n'existe pas, l'ajouter
                    strCommand = "Insert into InitiativesB Values" & Initiative(i)
                End If

                Try
                    cmdCommand.CommandText = strCommand
                    cmdCommand.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            Next
        End If
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| SaveGoals: Enregistre les Goals modifi�es                                                    |
    '|----------------------------------------------------------------------------------------------|
    Sub SaveGoals(ByVal sender As Object, ByVal e As EventArgs)
        'Save in database
        Dim dbConnSqlServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSqlServer.Open()
        SaveGoalsEntry(dbConnSqlServer)
        dbConnSqlServer.Close()

        Session("GoalsChanged") = Nothing
        Response.Redirect("Menu.aspx")
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| SaveGoalsEntry: Enregistre les Goals modifi�es                                               |
    '|----------------------------------------------------------------------------------------------|
    Sub SaveGoalsEntry(ByVal dbConnSqlServer As OleDbConnection)
        Dim Goal() As String
        Dim Values() As String
        Dim strReq As String
        Dim strWhere As String
        Dim strCommand As String = ""

        Dim i As Integer

        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim cmdCommand As New OleDbCommand

        cmdCommand.Connection = dbConnSqlServer

        If Session("GoalsChanged") <> Nothing Then
            Goal = Session("GoalsChanged").split(";")
            For i = 0 To UBound(Goal) - 1

                Values = Goal(i).Split(",")
                strWhere = " WHERE OsNo=" & Mid(Values(0), 2)
                strWhere += " AND Division=" & Values(1)
                strWhere += " AND FY=" & Values(2)

                'Voir s'il existe dans la BD
                strReq = "Select * from Goals" & strWhere
                cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
                dtTable = New DataTable
                cmdTable.Fill(dtTable)

                If dtTable.Rows.Count > 0 Then
                    'S'il existe, fait un update
                    strCommand = "UPDATE Goals Set Goal=" & Mid(Values(3), 1, Values(3).Length - 1) & strWhere
                Else
                    'S'il n'existe pas, l'ajouter
                    strCommand = "Insert into Goals Values" & Goal(i)
                End If

                Try
                    cmdCommand.CommandText = strCommand
                    cmdCommand.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            Next
        End If
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| ListCustomerType: Affiche et emplit la combobox "Customer Type" de forecasts                 |
    '|----------------------------------------------------------------------------------------------|
    Sub ListCustomerType()
        Dim CB As ComboBox = FindControl("CustomerType")
        Dim dtTable As New DataTable
        Dim newRow As DataRow
        Dim choices() As String = {"""A"" Customers", """B"" Customers"}
        Dim i As Integer

        dtTable.Columns.Add(New DataColumn("Choice"))
        dtTable.Columns.Add(New DataColumn("Value"))

        For i = 0 To UBound(choices)
            newRow = dtTable.NewRow()
            newRow("Choice") = choices(i)
            newRow("Value") = i + 1
            dtTable.Rows.Add(newRow)
        Next

        CB.AddItems(dtTable)
        CType(FindControl("Customer"), ComboBox).Enabled = False
        CType(FindControl("BRegion"), ComboBox).Enabled = False
        CType(FindControl("Division"), ComboBox).Enabled = False

        Session("CustomerSelected") = ""
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| ListCustomerType: Affiche, emplit et active les combobox de la page Forecasts                |
    '|----------------------------------------------------------------------------------------------|
    Sub AccessibleFields(ByVal dbConnSqlServer As OleDbConnection)
        Dim CBCustomerType As ComboBox = FindControl("CustomerType")
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")
        Dim CDivision As ComboBox = FindControl("Division")
        CDivision.Enabled = False
        Products.Enabled = False

        If CBCustomerType.Value = "1" Then
            CBCustomer.Enabled = True
            CBRegion.Enabled = False
        ElseIf CBCustomerType.Value = "2" Then
            CBCustomer.Enabled = False
            CBRegion.Enabled = True
        End If

        If (CBCustomer.Enabled = True And CBCustomer.Text <> "") Then
            CDivision.Enabled = True
            Products.Enabled = True
            Session("RegionSelected") = ""
            If Session("CustomerSelected") = "" Or Session("CustomerSelected") <> CBCustomer.Value Then
                ShowCustomerInfos(dbConnSqlServer)
                ListDivisions(dbConnSqlServer)
                Session("CustomerSelected") = CBCustomer.Value
            End If
        End If

        If (CBRegion.Enabled = True And CBRegion.Text <> "") Then
            CDivision.Enabled = True
            Products.Enabled = True
            Session("CustomerSelected") = ""
            If Session("RegionSelected") = "" Or Session("RegionSelected") <> CBRegion.Value Then
                ListDivisions(dbConnSqlServer)
                Session("RegionSelected") = CBRegion.Value
            End If
        End If

        If CDivision.Enabled = True And CDivision.Text <> "" Then
            If Session("ForecastEntry") = True Then
                FillForecast(dbConnSqlServer)
            ElseIf Session("InitiativesEntry") Then
                FillInitiatives(dbConnSqlServer)
            ElseIf Session("ExtraDetails") = True Then
                FillDetails(dbConnSqlServer)
            End If
        End If

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| ShowCustomerInfos: Affiche les informations du client (No et Ville) dans la page Forecasts   |
    '|----------------------------------------------------------------------------------------------|
    Sub ShowCustomerInfos(ByVal dbConnSqlServer As OleDbConnection)
        Dim CB As ComboBox = FindControl("Customer")
        Dim lblCity As Label = FindControl("CustomerCity")
        Dim lblCustomerNo As Label = FindControl("CustomerNo")

        Dim strReq = "Select City from Nomis where CustomerNo = " & CB.Value
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        Dim dtTable As New DataTable
        cmdTable.Fill(dtTable)

        If dtTable.Rows.Count > 0 Then
            lblCity.Text = dtTable.Rows(0)(0)
            lblCustomerNo.Text = CB.Value
        End If
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| ListCustomers: Remplit la liste customers de la page Forecasts                               |
    '|----------------------------------------------------------------------------------------------|
    Sub ListCustomers(ByVal dbConnSqlServer As OleDbConnection)
        Dim CB As ComboBox = FindControl("Customer")
        Dim strReq As String = "Select customerNo from Customers where AB='A' AND OSNO='" & Session("OS").ToString & "'"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        Dim dtTable As New DataTable
        cmdTable.Fill(dtTable)
        Dim myRow As DataRow

        Dim strIn As String

        For Each myRow In dtTable.Rows
            strIn += "'" & myRow(0) & "',"
        Next

        If strIn <> "" Then
            strIn = "(" & Mid(strIn, 1, strIn.Length - 1) & ")"
        Else
            strIn = "('')"
        End If

        strReq = "Select (CustomerName + ' - ' + Cast(CustomerNo as Varchar)), CustomerNo from NOMIS where CAST(CustomerNo AS  VARCHAR(15)) IN " & strIn & " group by CustomerName, CustomerNo order by CustomerName"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        dtTable = New DataTable
        cmdTable.Fill(dtTable)

        CB.AddItems(dtTable)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| ListDivisions: Remplit la liste Divisions de la page Forecasts                               |
    '|----------------------------------------------------------------------------------------------|
    Sub ListDivisions(ByVal dbConnSqlServer As OleDbConnection)
        Dim CBCustomers As ComboBox = FindControl("Customer")
        Dim CBRegions As ComboBox = FindControl("BRegion")
        Dim CBDivisions As ComboBox = FindControl("Division")
        Dim strReq As String
        Dim cmdTable As New OleDbDataAdapter
        Dim dtTable As New DataTable

        ' Data for Forecast
        strReq = "Select 'ALL' AS PrimaryCat, 'ALL' AS Value"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        If CBCustomers.Enabled = True Then
            strReq = "Select PrimaryCat, PrimaryCat AS Value from NOMIS " & _
                        " where FY IS NOT NULL AND FY >= '" & FiscalYear() - 2 & _
                        "' AND PrimaryCat <> '' group by PrimaryCat " & _
                        "order by PrimaryCat"
        ElseIf CBRegions.Enabled = True Then
            strReq = "SELECT PRIMARYCAT, NOMIS.PRIMARYCAT AS Value " & _
                        "FROM NOMIS INNER JOIN Customers ON NOMIS.CUSTOMERNO = Customers.CustomerNo INNER JOIN " & _
                        "Regions ON Customers.RegionNo = Regions.RegionNo " & _
                        "WHERE(dbo.Regions.RegionNo = " & CBRegions.Value & ") AND Regions.OsNo='" & Session("OS") & "'" & _
                        "GROUP BY dbo.NOMIS.PRIMARYCAT " & _
                            "order by PrimaryCat"
        End If

        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)
        CBDivisions.AddItems(dtTable)
        CBDivisions.SelectedIndex = 0

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| ListRegions: Remplit la liste Regions de la page Forecasts                                   |
    '|----------------------------------------------------------------------------------------------|
    Sub ListRegions(ByVal dbConnSqlServer As OleDbConnection)
        Dim CBRegions As ComboBox = FindControl("BRegion")

        Dim strReq = "Select Region, RegionNo from Regions where OSNo='" & Session("OS") & "'"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        Dim dtTable As New DataTable
        cmdTable.Fill(dtTable)

        CBRegions.AddItems(dtTable)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillForecast: Remplit les donn�es � afficher pour les forecasts                              |
    '|----------------------------------------------------------------------------------------------|
    Sub FillForecast(ByVal dbConnSqlServer As OleDbConnection)
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")
        Dim CBDivision As ComboBox = FindControl("Division")

        Dim i As Integer

        Dim dtTable As New DataTable

        Dim strReq As String

        If CBCustomer.Enabled = True Then
            strReq = ReqCustomerA()
        Else
            strReq = ReqCustomerB()
        End If

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        Dim dvTable As New DataView(dtTable)

        If UCase(CBDivision.Text) <> "ALL" Then
            dvTable.RowFilter = "Division='" & CBDivision.Text & "'"
        End If

        FillForecastTable(dvTable)
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillForecastTable: Affiche les donn�es dans le tableau forecasts                             |
    '|----------------------------------------------------------------------------------------------|
    Sub FillForecastTable(ByVal dvTable As DataView)
        Dim x As Integer
        Dim y As Integer
        Dim intLargeurColonnes(8) As Integer
        Dim myRow As DataRow
        Dim newRow As TableRow
        Dim newCell As TableCell
        Dim field As TextBox
        Dim strInsert As String
        Dim cssClass As String
        Dim Totals(6) As Integer

        intLargeurColonnes(0) = 10
        intLargeurColonnes(1) = 150
        intLargeurColonnes(2) = 50
        intLargeurColonnes(3) = 50
        intLargeurColonnes(4) = 50
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 70
        intLargeurColonnes(8) = 70

        Dim DGForecasts As Table = FindControl("DGForecasts")
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")

        For x = 0 To dvTable.Count - 1
            newRow = New TableRow
            For y = 0 To dvTable.Table.Columns.Count - 2
                cssClass = ""
                newCell = New TableCell
                newCell.Width = Unit.Pixel(intLargeurColonnes(y))
                If y < 7 Then
                    If Not IsDBNull(dvTable.Item(x)(y)) Then
                        newCell.Text = dvTable.Item(x)(y)
                    End If

                    If y >= 2 And y <= 4 Then
                        cssClass = " ColumnPaleYellow "
                    ElseIf y >= 5 And y <= 6 Then
                        cssClass = " ColumnPaleGray "
                    End If

                    If y >= 2 Then
                        cssClass += " Centre "
                        Totals(y - 2) += Val(newCell.Text)
                    End If

                Else
                    field = New TextBox

                    If CBCustomer.Enabled = True Then
                        field.ID = dvTable.Table.Columns(y).ToString & "_" & dvTable(x)(0) & "_" & CBCustomer.Value & "_A"
                    Else
                        field.ID = dvTable.Table.Columns(y).ToString & "_" & dvTable(x)(0) & "_" & CBRegion.Value & "_B"
                    End If
                    field.Attributes.Add("OnChange", "javascript:ForecastChange();")

                    If ForecastHasChanged(dvTable.Item(x), y) <> Nothing Then
                        field.Text = ForecastHasChanged(dvTable.Item(x), y)
                    Else
                        If Not IsDBNull(dvTable.Item(x)(y)) Then
                            field.Text = dvTable.Item(x)(y)
                        Else
                            field.Text = 0
                        End If
                    End If

                    If y = 7 Then
                        field.Enabled = Now >= CDate("01-jan-" & FiscalYear() - 1) And Now < CDate("31-dec-" & FiscalYear() - 1)
                    ElseIf y = 8 Then
                        field.Enabled = Now >= CDate("01-jan-" & FiscalYear()) And Now <= CDate("31-dec-" & FiscalYear())
                    End If

                    Totals(y - 2) += Val(field.Text)

                    field.CssClass = "FieldForecast Centre texte " & IIf(field.Enabled, "TextEntry", "")
                    newCell.Controls.Add(field)
                End If

                newCell.CssClass = "BordureTableau " & cssClass
                newRow.Cells.Add(newCell)
            Next
            DGForecasts.Rows.Add(newRow)
        Next

        FillForecastTotal(Totals)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillForecastTotal: Affiche les sommes au bas du tableau Forecasts                            |
    '|----------------------------------------------------------------------------------------------|
    Sub FillForecastTotal(ByVal Totals() As Integer)
        Dim DGForecastsTotal As Table = FindControl("DGForecastsTotal")
        Dim Total As New TableRow
        Dim ColumnName As TableCell

        Dim intLargeurColonnes(7) As Integer
        Dim i As Integer

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
        ColumnName.CssClass = "BordureHeader"
        Total.Cells.Add(ColumnName)

        For i = 0 To UBound(Totals)
            ColumnName = New TableCell
            ColumnName.Text = Totals(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            ColumnName.HorizontalAlign = HorizontalAlign.Center
            Total.Cells.Add(ColumnName)

            If i <= 2 Then
                ColumnName.CssClass = "BordureHeader ColumnYellow"
            ElseIf i <= 4 Then
                ColumnName.CssClass = "BordureHeader ColumnGray"
            End If
        Next

        ColumnName = New TableCell
        ColumnName.Width = New Unit(intLargeurColonnes(7))
        ColumnName.CssClass = "BordureHeader"
        Total.Cells.Add(ColumnName)

        DGForecastsTotal.Controls.Add(Total)


    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillInitiatives: Remplit les donn�es � afficher pour les initiatives                         |
    '|----------------------------------------------------------------------------------------------|
    Sub FillInitiatives(ByVal dbConnSqlServer As OleDbConnection)
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")
        Dim CBDivision As ComboBox = FindControl("Division")

        Dim i As Integer

        Dim dtTable As New DataTable

        Dim strReq As String

        Dim strWhere As String = ""

        If CBCustomer.Enabled = True Then
            strWhere = "InitiativesA.CustomerNo = " & CBCustomer.Value & " AND OSNO='" & Session("OS") & "' AND InitiativeNo = INI.InitiativeNo"
        Else
            strWhere = "InitiativesB.RegionNo = " & CBRegion.Value & " AND InitiativeNo = INI.InitiativeNo"
        End If

        If CBCustomer.Enabled = True Then
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

        If UCase(CBDivision.Text) <> "ALL" Then
            dvTable.RowFilter = "Division='" & CBDivision.Text & "'"
        End If

        FillInitiativesTable(dvTable)
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillInitiativesTable: Affiche les donn�es dans le tableau initiatives                        |
    '|----------------------------------------------------------------------------------------------|
    Sub FillInitiativesTable(ByVal dvTable As DataView)
        Dim x As Integer
        Dim y As Integer
        Dim intLargeurColonnes(8) As Integer
        Dim myRow As DataRow
        Dim newRow As TableRow
        Dim newCell As TableCell
        Dim field As TextBox
        Dim strInsert As String
        Dim cssClass As String

        intLargeurColonnes(0) = 120
        intLargeurColonnes(1) = 190
        intLargeurColonnes(2) = 70
        intLargeurColonnes(3) = 70
        intLargeurColonnes(4) = 180
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 180

        Dim DGInitiatives As Table = FindControl("DGInitiatives")
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")

        For x = 0 To dvTable.Count - 1
            newRow = New TableRow
            For y = 0 To dvTable.Table.Columns.Count - 2
                cssClass = ""
                newCell = New TableCell
                newCell.Width = Unit.Pixel(intLargeurColonnes(y))
                If y < 2 Then
                    If Not IsDBNull(dvTable.Item(x)(y)) Then
                        newCell.Text = dvTable.Item(x)(y)
                    End If
                Else
                    cssClass = " Centre "
                    field = New TextBox

                    If CBCustomer.Enabled = True Then
                        field.ID = dvTable.Table.Columns(y).ToString & "_" & dvTable(x)(8) & "_" & CBCustomer.Value & "_A"
                    Else
                        field.ID = dvTable.Table.Columns(y).ToString & "_" & dvTable(x)(8) & "_" & CBRegion.Value & "_B"
                    End If
                    field.Attributes.Add("OnChange", "javascript:InitiativesChange();")

                    If InitiativesHasChanged(dvTable.Item(x), y) <> Nothing Then
                        field.Text = InitiativesHasChanged(dvTable.Item(x), y)
                    Else
                        If Not IsDBNull(dvTable.Item(x)(y)) Then
                            field.Text = dvTable.Item(x)(y)
                        End If
                    End If

                    If y >= 2 And y <= 4 Then
                        field.Enabled = Now >= CDate("01-jan-" & FiscalYear() - 1) And Now < CDate("31-dec-" & FiscalYear() - 1)
                    ElseIf y >= 5 And y <= 7 Then
                        field.Enabled = Now >= CDate("01-jan-" & FiscalYear()) And Now <= CDate("31-dec-" & FiscalYear())
                    End If

                    If y = "4" Or y = "7" Then
                        field.CssClass = "FieldInitiativesNotes texte " & IIf(field.Enabled, "TextEntry", "")
                        cssClass += " ColumnInitiativeNotes "
                    Else
                        field.CssClass = "FieldInitiatives Centre texte " & IIf(field.Enabled, "TextEntry", "")
                    End If

                    newCell.Controls.Add(field)
                End If

                newCell.CssClass = "BordureTableau " & cssClass
                newRow.Cells.Add(newCell)
            Next
            DGInitiatives.Rows.Add(newRow)
        Next

    End Sub
    '
    '|----------------------------------------------------------------------------------------------|
    '| FillDetails: Remplit les donn�es � afficher pour les d�tails des forecasts                   |
    '|----------------------------------------------------------------------------------------------|
    Sub FillDetails(ByVal dbConnSqlServer As OleDbConnection)
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")
        Dim CBDivision As ComboBox = FindControl("Division")

        Dim i As Integer

        Dim dtTable As New DataTable

        Dim strReq As String

        If CBCustomer.Enabled = True Then
            strReq = ReqDetailsCustomerA()
        Else
            strReq = ReqDetailsCustomerB()
        End If

        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        Dim dvTable As New DataView(dtTable)

        If UCase(CBDivision.Text) <> "ALL" Then
            dvTable.RowFilter = "Division='" & CBDivision.Text & "'"
        End If

        FillDetailsTable(dvTable)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillDetailsTable: Affiche les donn�es dans le tableau details                                |
    '|----------------------------------------------------------------------------------------------|
    Sub FillDetailsTable(ByVal dvTable As DataView)
        Dim x As Integer
        Dim y As Integer
        Dim intLargeurColonnes(9) As Integer
        Dim myRow As DataRow
        Dim newRow As TableRow
        Dim newCell As TableCell
        Dim field As TextBox
        Dim strInsert As String
        Dim cssClass As String
        Dim Totals(6) As Integer

        intLargeurColonnes(0) = 10
        intLargeurColonnes(1) = 150
        intLargeurColonnes(2) = 50
        intLargeurColonnes(3) = 50
        intLargeurColonnes(4) = 50
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 70
        intLargeurColonnes(8) = 70

        Dim DGForecasts As Table = FindControl("DGDetails")
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")

        For x = 0 To dvTable.Count - 1
            newRow = New TableRow
            For y = 0 To dvTable.Table.Columns.Count - 2
                cssClass = ""
                newCell = New TableCell
                newCell.Width = Unit.Pixel(intLargeurColonnes(y))
                If Not IsDBNull(dvTable.Item(x)(y)) Then
                    newCell.Text = dvTable.Item(x)(y)
                End If

                If y >= 2 And y <= 4 Then
                    cssClass = " ColumnPaleYellow "
                ElseIf y >= 5 And y <= 6 Then
                    cssClass = " ColumnPaleGray "
                ElseIf y >= 7 And y <= 8 Then
                    cssClass = " ColumnPaleBlue "
                End If

                If y >= 2 Then
                    cssClass += " Centre "
                    Totals(y - 2) += Val(newCell.Text)
                End If

                newCell.CssClass = "BordureTableau " & cssClass
                newRow.Cells.Add(newCell)
            Next
            DGForecasts.Rows.Add(newRow)
        Next

        FillDetailsTotal(Totals)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillDetailsTotal: Affiche les sommes au bas du tableau Forecasts - D�tails                   |
    '|----------------------------------------------------------------------------------------------|
    Sub FillDetailsTotal(ByVal Totals() As Integer)
        Dim DGForecastsTotal As Table = FindControl("DGDetailsTotal")
        Dim Total As New TableRow
        Dim ColumnName As TableCell

        Dim intLargeurColonnes(7) As Integer
        Dim i As Integer

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
        ColumnName.CssClass = "BordureHeader"
        Total.Cells.Add(ColumnName)

        For i = 0 To UBound(Totals)
            ColumnName = New TableCell
            ColumnName.Text = Totals(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            ColumnName.HorizontalAlign = HorizontalAlign.Center
            Total.Cells.Add(ColumnName)

            If i <= 2 Then
                ColumnName.CssClass = "BordureHeader ColumnYellow"
            ElseIf i <= 4 Then
                ColumnName.CssClass = "BordureHeader ColumnGray"
            ElseIf i <= 6 Then
                ColumnName.CssClass = "BordureHeader ColumnBlue"
            End If
        Next

        ColumnName = New TableCell
        ColumnName.Width = New Unit(intLargeurColonnes(7))
        ColumnName.CssClass = "BordureHeader"
        Total.Cells.Add(ColumnName)

        DGForecastsTotal.Controls.Add(Total)


    End Sub


    '
    '|----------------------------------------------------------------------------------------------|
    '| FillGoals: Remplit les donn�es � afficher pour les Goals                                     |
    '|----------------------------------------------------------------------------------------------|
    Sub FillGoals(ByVal dbConnSqlServer As OleDbConnection)
        Dim cbOs As ComboBox = FindControl("CBOS")

        If cbOs.Value <> "" Then
            Dim strReq = "SELECT PRIMARYCAT, " & _
                "CAST(SUM(CASE WHEN FY = '" & FiscalYear() - 2 & "' AND OSNO='" & cbOs.Value & "' THEN BOOKINGS ELSE NULL END) AS NUMERIC) AS Year1, " & _
                "CAST(SUM(CASE WHEN FY = '" & FiscalYear() - 1 & "' AND OSNO='" & cbOs.Value & "' THEN BOOKINGS ELSE NULL END) AS NUMERIC)  AS Year2, " & _
                "CAST(SUM(CASE WHEN FY = '" & FiscalYear() & "' AND OSNO='" & cbOs.Value & "' THEN BOOKINGS ELSE NULL END) AS NUMERIC)  AS Year3, " & _
                "ISNULL((SELECT SUM(Forecast) from ForecastA FA where  " & _
                "(Select PrimaryCat from Nomis where PC = FA.PC Group by PrimaryCat) = N1.PRIMARYCAT AND FY='" & FiscalYear() & "' AND OSNO='" & cbOs.Value & "'),0) +  " & _
                "ISNULL((SELECT SUM(Forecast) from ForecastB FB, Regions where  " & _
                "FB.RegionNo = Regions.RegionNo AND (Select PrimaryCat from Nomis where PC = FB.PC Group by PrimaryCat) = N1.PRIMARYCAT AND FY='" & FiscalYear() & "' AND OSNO='" & cbOs.Value & "'),0)  AS Forecast1, " & _
                "(SELECT Goal From Goals where Division=N1.PrimaryCat AND OSNO='" & cbOs.Value & "' AND FY='" & FiscalYear() & "') AS GOALS1, " & _
                "ISNULL((SELECT SUM(Forecast) from ForecastA FA where  " & _
                "(Select PrimaryCat from Nomis where PC = FA.PC Group by PrimaryCat) = N1.PRIMARYCAT AND FY='" & FiscalYear() + 1 & "' AND OSNO='" & cbOs.Value & "'),0) +  " & _
                "ISNULL((SELECT SUM(Forecast) from ForecastB FB, Regions where  " & _
                "FB.RegionNo = Regions.RegionNo AND (Select PrimaryCat from Nomis where PC = FB.PC Group by PrimaryCat) = N1.PRIMARYCAT AND FY='" & FiscalYear() + 1 & "' AND OSNO='" & cbOs.Value & "'),0)  AS Forecast2, " & _
                "(SELECT Goal From Goals where Division=N1.PrimaryCat AND OSNO='" & cbOs.Value & "' AND FY='" & FiscalYear() + 1 & "') AS GOALS2 " & _
                "FROM dbo.NOMIS N1 " & _
                "WHERE PRIMARYCAT <> '' " & _
                "GROUP BY PRIMARYCAT " & _
                "ORDER BY PRIMARYCAT"
            Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
            Dim dtTable As New DataTable

            cmdTable.Fill(dtTable)
            FillGoalsTable(dtTable)
        End If


    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillGoalsTable: Affiche les donn�es dans le tableau Goals                                    |
    '|----------------------------------------------------------------------------------------------|
    Sub FillGoalsTable(ByVal dtTable As DataTable)
        Dim cbOs As ComboBox = FindControl("CBOS")

        Dim DGGoals As Table = FindControl("DGGoals")

        Dim intLargeurColonnes(8) As Integer
        Dim myRow As DataRow
        Dim newRow As TableRow
        Dim newCell As TableCell
        Dim field As TextBox
        Dim strInsert As String
        Dim cssClass As String

        Dim totals(6) As Integer

        intLargeurColonnes(0) = 200
        intLargeurColonnes(1) = 90
        intLargeurColonnes(2) = 90
        intLargeurColonnes(3) = 90
        intLargeurColonnes(4) = 90
        intLargeurColonnes(5) = 90
        intLargeurColonnes(6) = 90
        intLargeurColonnes(7) = 90

        Dim x As Integer
        Dim y As Integer

        For x = 0 To dtTable.Rows.Count - 1
            newRow = New TableRow
            For y = 0 To dtTable.Columns.Count - 1
                cssClass = ""
                newCell = New TableCell
                newCell.Width = Unit.Pixel(intLargeurColonnes(y))
                If y <> 5 And y <> 7 Then
                    If Not IsDBNull(dtTable.Rows(x)(y)) Then
                        newCell.Text = dtTable.Rows(x)(y)
                    End If

                    If y >= 1 And y <= 3 Then
                        cssClass += " ColumnPaleYellow "
                    ElseIf y = 4 Or y = 6 Then
                        cssClass += " ColumnPaleGray "
                    End If

                    If y > 0 Then
                        totals(y - 1) += Val(newCell.Text)
                    End If

                Else
                    field = New TextBox

                    field.ID = dtTable.Columns(y).ToString & "_" & Trim(dtTable.Rows(x)(0)) & "_" & cbOs.Value

                    field.Attributes.Add("OnChange", "javascript:GoalsChange();")

                    If GoalsHasChanged(dtTable.Rows(x), y) <> Nothing Then
                        field.Text = GoalsHasChanged(dtTable.Rows(x), y)
                    Else
                        If Not IsDBNull(dtTable.Rows(x)(y)) Then
                            field.Text = dtTable.Rows(x)(y)
                        End If
                    End If

                    field.CssClass = "FieldGoals Centre texte TextEntry"
                    newCell.Controls.Add(field)

                    totals(y - 1) += Val(field.Text)
                End If

                If y > 0 Then
                    cssClass += " Centre "
                End If
                newCell.CssClass = "BordureTableau " & cssClass
                newRow.Cells.Add(newCell)
            Next
            DGGoals.Rows.Add(newRow)
        Next

        FillGoalsTotal(totals)
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillGoalsTotal: Affiche les sommes au bas du tableau Goals                                   |
    '|----------------------------------------------------------------------------------------------|
    Sub FillGoalsTotal(ByVal Totals() As Integer)
        Dim DGGoalsTotal As Table = FindControl("DGGoalsTotal")
        Dim Total As New TableRow
        Dim ColumnName As TableCell

        Dim intLargeurColonnes(7) As Integer
        Dim i As Integer

        '200
        intLargeurColonnes(0) = 90
        intLargeurColonnes(1) = 90
        intLargeurColonnes(2) = 90
        intLargeurColonnes(3) = 90
        intLargeurColonnes(4) = 90
        intLargeurColonnes(5) = 90
        intLargeurColonnes(6) = 90

        ColumnName = New TableCell
        ColumnName.Text = "Total:"
        ColumnName.Width = New Unit(200)
        ColumnName.CssClass = "BordureHeader"
        Total.Cells.Add(ColumnName)

        For i = 0 To UBound(Totals)
            ColumnName = New TableCell
            ColumnName.Text = Totals(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            ColumnName.HorizontalAlign = HorizontalAlign.Center
            Total.Cells.Add(ColumnName)

            If i <= 2 Then
                ColumnName.CssClass = "BordureHeader ColumnYellow"
            ElseIf i = 3 Or i = 5 Then
                ColumnName.CssClass = "BordureHeader ColumnGray"
            End If
        Next

        ColumnName = New TableCell
        ColumnName.Width = New Unit(intLargeurColonnes(7))
        ColumnName.CssClass = "BordureHeader"
        Total.Cells.Add(ColumnName)

        DGGoalsTotal.Controls.Add(Total)
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| ForecastHasChanged: V�rifie si le forecast a chang�, si oui, retourner sa valeur             |
    '|----------------------------------------------------------------------------------------------|
    Function ForecastHasChanged(ByVal row As DataRowView, ByVal y As Integer) As String
        Dim strValue As String = Nothing
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")
        Dim recherche As String
        Dim elementRecherche As String
        Dim Forecasts() As String
        Dim NoCusReg As String

        Dim i As Integer

        If CBCustomer.Enabled Then
            recherche = Session("ForecastChangedA")
            NoCusReg = CBCustomer.Value
        ElseIf CBRegion.Enabled Then
            recherche = Session("ForecastChangedB")
            NoCusReg = CBRegion.Value
        End If

        elementRecherche = "("

        elementRecherche += NoCusReg & ", "

        If CBCustomer.Enabled Then
            elementRecherche += "'" & Session("OS") & "', "
        End If

        If y = 7 Then
            elementRecherche += "'" & FiscalYear() & "', "
        ElseIf y = 8 Then
            elementRecherche += "'" & FiscalYear() + 1 & "', "
        End If

        elementRecherche += "'" & row(0) & "', "

        If recherche <> "" Then
            Dim pos As Integer = recherche.LastIndexOf(elementRecherche)
            Dim whereStop As Integer

            If pos >= 0 Then
                strValue = Mid(recherche, pos + elementRecherche.Length)
                whereStop = strValue.IndexOf(")")
                strValue = Mid(strValue, 1, whereStop)
            End If
        End If

        ForecastHasChanged = strValue

    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| InitiativesHasChanged: V�rifie si l'initiatives a chang�, si oui, retourner sa valeur        |
    '|----------------------------------------------------------------------------------------------|
    Function InitiativesHasChanged(ByVal row As DataRowView, ByVal y As Integer) As String
        Dim strValue As String = Nothing
        Dim CBCustomer As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")
        Dim recherche As String
        Dim elementRecherche As String
        Dim Forecasts() As String
        Dim NoCusReg As String
        Dim NoCol As Integer
        Dim values() As String

        Dim i As Integer

        If CBCustomer.Enabled Then
            recherche = Session("InitiativeChangedA")
            NoCusReg = CBCustomer.Value
        ElseIf CBRegion.Enabled Then
            recherche = Session("InitiativeChangedB")
            NoCusReg = CBRegion.Value
        End If

        elementRecherche = "("
        elementRecherche += NoCusReg & ", "

        If CBCustomer.Enabled Then
            elementRecherche += "'" & Session("OS") & "', "
        End If

        If y >= 2 And y <= 4 Then
            elementRecherche += "'" & FiscalYear() & "', "
            NoCol = 3 - (4 - y)
        ElseIf y >= 5 And y <= 7 Then
            elementRecherche += "'" & FiscalYear() + 1 & "', "
            NoCol = 3 - (7 - y)
        End If

        elementRecherche += row(8).ToString & ", "

        If recherche <> "" Then
            Dim pos As Integer = recherche.LastIndexOf(elementRecherche)
            Dim whereStop As Integer

            If pos >= 0 Then

                Dim posDepart As Integer = pos + elementRecherche.Length
                Dim positionFin As Integer = Mid(recherche, posDepart).IndexOf(")")
                strValue = Mid(recherche, posDepart, positionFin)

                values = strValue.Split(",")
                values(NoCol - 1) = values(NoCol - 1).Replace("'", "")
                values(NoCol - 1) = values(NoCol - 1).Replace(")", "")

                strValue = values(NoCol - 1)
            End If
        End If

        InitiativesHasChanged = strValue

    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| GoalsHasChanged: V�rifie si le Goal a chang�, si oui, retourner sa valeur                    |
    '|----------------------------------------------------------------------------------------------|
    Function GoalsHasChanged(ByVal row As DataRow, ByVal y As Integer)
        Dim strValue As String = Nothing
        Dim CBOS As ComboBox = FindControl("CBOS")
        Dim recherche As String
        Dim elementRecherche As String
        Dim Forecasts() As String
        Dim NoOS As String
        Dim NoCol As Integer
        Dim values() As String

        Dim i As Integer

        recherche = Session("GoalsChanged")
        NoOS = CBOS.Value

        elementRecherche = "("
        elementRecherche += "'" & NoOS & "', "
        elementRecherche += "'" & Trim(row(0).ToString) & "', "

        If y = 5 Then
            elementRecherche += "'" & FiscalYear() & "', "
        ElseIf y = 7 Then
            elementRecherche += "'" & FiscalYear() + 1 & "', "
        End If

        If recherche <> "" Then
            Dim pos As Integer = recherche.LastIndexOf(elementRecherche)
            Dim whereStop As Integer

            If pos >= 0 Then
                strValue = Mid(recherche, pos + elementRecherche.Length)
                whereStop = strValue.IndexOf(")")
                strValue = Mid(strValue, 1, whereStop)
            End If
        End If

        GoalsHasChanged = strValue
    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| ReqCustomerA: Retourne la requ�te pour les customersA afin d'afficher le forecast            |
    '| Selects all data from forecastA and others from Nomis                                        |
    '|----------------------------------------------------------------------------------------------|
    Function ReqCustomerA() As String
        Dim CBCustomer As ComboBox = FindControl("Customer")

        Dim NbDayYear As Integer = DateDiff(DateInterval.Day, CDate("01-oct-" & FiscalYear() - 1), CDate("01-oct-" & FiscalYear()))
        Dim DayOfFY As Integer = DateDiff(DateInterval.DayOfYear, CDate("01-oct-" & FiscalYear() - 1), Now) + 1

        If Now > CDate("1-oct-" & FiscalYear()) Then
            DayOfFY = NbDayYear
        End If

        ReqCustomerA = "Select PC, (Select ProductDesc from nomis where Nomis.PC = FA.PC group by ProductDesc) AS ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "(Select Forecast from ForecastA where PC=FA.PC AND OSNO = FA.OSNO AND CustomerNo=" & CBCustomer.Value & " AND FY='" & FiscalYear() & "') AS CurrentForecast, " & _
                        "(Select Forecast from ForecastA where PC=FA.PC AND OSNO = FA.OSNO AND CustomerNo=" & CBCustomer.Value & " AND FY='" & FiscalYear() + 1 & "') AS NextYearForecast, " & _
                        "(Select PrimaryCat from nomis where Nomis.PC = FA.PC group by PrimaryCat) AS Division " & _
                        "from ForecastA FA " & _
                        "WHERE CustomerNo=" & CBCustomer.Value & " and OsNo=" & Session("OS") & _
                        " UNION " & _
                        "SELECT N1.PC, ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "0 AS CurrentForecast, " & _
                        "0 AS NextYearForecast, " & _
                        "PrimaryCat as Division  " & _
                        "FROM NOMIS N1 " & _
                        "WHERE NOT EXISTS (SELECT * From ForecastA where N1.PC = PC AND N1.CUSTOMERNO = CUSTOMERNO AND N1.OSNO = OSNO) " & _
                        "AND OSNO = '" & Session("OS") & "' AND CUSTOMERNO=" & CBCustomer.Value & " AND PC is not null and PC <> '' " & _
                        "Group by N1.Pc, productDesc, PrimaryCat"
    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| ReqCustomerB: Retourne la requ�te pour les customersB afin d'afficher le forecast            |
    '| Selects all data from forecastB and others from Nomis                                        |
    '|----------------------------------------------------------------------------------------------|
    Function ReqCustomerB() As String
        Dim CBRegion As ComboBox = FindControl("BRegion")

        Dim NbDayYear As Integer = DateDiff(DateInterval.Day, CDate("01-oct-" & FiscalYear() - 1), CDate("01-oct-" & FiscalYear()))
        Dim DayOfFY As Integer = DateDiff(DateInterval.DayOfYear, CDate("01-oct-" & FiscalYear() - 1), Now) + 1

        If Now > CDate("1-oct-" & FiscalYear()) Then
            DayOfFY = NbDayYear
        End If

        ReqCustomerB = "SELECT FB.PC, (Select ProductDesc from nomis where Nomis.PC = FB.PC group by ProductDesc) AS ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Session("OS") & "' AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Session("OS") & "' AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Session("OS") & "' AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Session("OS") & "' AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Session("OS") & "' AND Customers.RegionNo = " & CBRegion.Value & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Nomis.OSNo='" & Session("OS") & "' AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "(Select Forecast from ForecastB where PC=FB.PC AND RegionNo=" & CBRegion.Value & " AND FY='" & FiscalYear() & "') AS CurrentForecast, " & _
                        "(Select Forecast from ForecastB where PC=FB.PC AND RegionNo=" & CBRegion.Value & " AND FY='" & FiscalYear() + 1 & "') AS NextYearForecast, " & _
                        "(Select PrimaryCat from nomis where Nomis.PC = FB.PC group by PrimaryCat) AS Division " & _
                        "FROM ForecastB FB INNER JOIN " & _
                        "Regions ON FB.RegionNo = Regions.RegionNo INNER JOIN " & _
                        "Customers ON Regions.OsNo = Customers.OSNo " & _
                        "WHERE (dbo.Regions.RegionNo = " & CBRegion.Value & ") " & _
                        "GROUP BY FB.PC, FB.Forecast, Regions.OsNo " & _
                        "UNION " & _
                        "SELECT N1.PC, ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.OSNo='" & Session("OS") & "' AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "0 AS CurrentForecast, " & _
                        "0 AS NextYearForecast, " & _
                        "PrimaryCat as Division " & _
                        "FROM dbo.Customers C1 INNER JOIN " & _
                        "dbo.NOMIS N1 ON C1.CustomerNo = N1.CUSTOMERNO AND C1.OSNo = N1.OSNO " & _
                        "WHERE (C1.RegionNo = " & CBRegion.SelectedValue & ") AND " & _
                        "((SELECT COUNT(PC) FROM ForecastB WHERE PC = N1.PC AND RegionNo = C1.RegionNo) = 0) " & _
                        " AND N1.PC IS NOT NULL AND N1.PC <> '' " & _
                        "GROUP BY N1.PC, ProductDesc, PrimaryCat"
    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| ReqDetailsCustomerA: Retourne la requ�te pour les customersA afin d'afficher le forecast     |
    '| Selects all data from forecastA and others from Nomis                                        |
    '|----------------------------------------------------------------------------------------------|
    Function ReqDetailsCustomerA() As String
        Dim CBCustomer As ComboBox = FindControl("Customer")

        Dim NbDayYear As Integer = DateDiff(DateInterval.Day, CDate("01-oct-" & FiscalYear() - 1), CDate("01-oct-" & FiscalYear()))
        Dim DayOfFY As Integer = DateDiff(DateInterval.DayOfYear, CDate("01-oct-" & FiscalYear() - 1), Now) + 1

        If Now > CDate("1-oct-" & FiscalYear()) Then
            DayOfFY = NbDayYear
        End If

        ReqDetailsCustomerA = "Select PC, (Select ProductDesc from nomis where Nomis.PC = FA.PC group by ProductDesc) AS ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = FA.PC AND Nomis.CustomerNo = FA.CustomerNo AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote where PC=FA.PC AND Quote.Customer=" & CBCustomer.Value & "), 0) AS numeric), 0) AS QuoteTotal, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote where PC=FA.PC AND Status=0 AND Quote.Customer=" & CBCustomer.Value & "), 0) AS numeric), 0) AS QuoteOpen, " & _
                        "(Select PrimaryCat from nomis where Nomis.PC = FA.PC group by PrimaryCat) AS Division " & _
                        "from ForecastA FA " & _
                        "WHERE CustomerNo=" & CBCustomer.Value & _
                        " UNION " & _
                        "SELECT N1.PC, ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS WHERE Nomis.PC = N1.PC AND Nomis.CustomerNo = " & CBCustomer.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote where PC=N1.PC AND Quote.Customer=" & CBCustomer.Value & "), 2) AS numeric), 0) AS QuoteTotal, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote where PC=N1.PC AND Status=0 AND Quote.Customer=" & CBCustomer.Value & "), 2) AS numeric), 0) AS QuoteOpen, " & _
                        "PrimaryCat as Division  " & _
                        "FROM NOMIS N1 " & _
                        "WHERE NOT EXISTS (SELECT * From ForecastA where N1.PC = PC AND N1.CUSTOMERNO = CUSTOMERNO) " & _
                        "AND CUSTOMERNO=" & CBCustomer.Value & " AND PC is not null and PC <> '' " & _
                        "Group by N1.Pc, productDesc, PrimaryCat"
    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| ReqDetailsCustomerB: Retourne la requ�te pour les customersB afin d'afficher le forecast     |
    '| Selects all data from forecastB and others from Nomis                                        |
    '|----------------------------------------------------------------------------------------------|
    Function ReqDetailsCustomerB() As String
        Dim CBRegion As ComboBox = FindControl("BRegion")

        Dim NbDayYear As Integer = DateDiff(DateInterval.Day, CDate("01-oct-" & FiscalYear() - 1), CDate("01-oct-" & FiscalYear()))
        Dim DayOfFY As Integer = DateDiff(DateInterval.DayOfYear, CDate("01-oct-" & FiscalYear() - 1), Now) + 1

        If Now > CDate("1-oct-" & FiscalYear()) Then
            DayOfFY = NbDayYear
        End If

        ReqDetailsCustomerB = "SELECT FB.PC AS PrC, (Select ProductDesc from nomis where Nomis.PC = FB.PC group by ProductDesc) AS ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                        "((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = FB.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote, Customers where Quote.Customer=Customers.CustomerNo AND PC=FB.PC AND Customers.RegionNo=" & CBRegion.Value & "), 0) AS numeric), 0) AS QuoteTotal, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote, Customers  where Quote.Customer=Customers.CustomerNo  AND Status=0 AND PC=FB.PC AND Customers.RegionNo=" & CBRegion.Value & "), 0) AS numeric), 0) AS QuoteOpen, " & _
                        "(Select PrimaryCat from nomis where Nomis.PC = FB.PC group by PrimaryCat) AS Division " & _
                        "FROM ForecastB FB INNER JOIN " & _
                        "Regions ON FB.RegionNo = Regions.RegionNo " & _
                        "WHERE (dbo.Regions.RegionNo = " & CBRegion.Value & ") " & _
                        "GROUP BY FB.PC, FB.Forecast " & _
                        "UNION " & _
                        "SELECT N1.PC  AS PrC, ProductDesc, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 2 & "),0) AS Numeric), 0) AS YEAR1, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() - 1 & "),0) AS Numeric), 0) AS YEAR2, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & "),0) AS Numeric), 0) AS YEAR3, " & _
                        "ISNULL(CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & ",0) AS Numeric), 0) AS Prediction, " & _
                        "CAST(Round((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY Between '" & FiscalYear() - 2 & "' AND '" & FiscalYear() - 1 & "') + " & _
                            "((Select Sum(BOOKINGS) From NOMIS, Customers WHERE NOMIS.CustomerNo = Customers.CustomerNo AND Nomis.PC = N1.PC AND Customers.RegionNo = " & CBRegion.Value & " AND FY=" & FiscalYear() & ") * " & (NbDayYear / DayOfFY) & "), 0) / 3 AS numeric) AS AVG, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote, Customers where Quote.Customer=Customers.CustomerNo AND PC=N1.PC AND Customers.RegionNo=" & CBRegion.Value & "), 0) AS numeric), 0) AS QuoteTotal, " & _
                        "ISNULL(Cast(Round((Select Sum(Total) from Quote, Customers  where Quote.Customer=Customers.CustomerNo  AND Status=0 AND PC=N1.PC AND Customers.RegionNo=" & CBRegion.Value & "), 0) AS numeric), 0) AS QuoteOpen, " & _
                        "PrimaryCat as Division " & _
                        "FROM dbo.Customers C1 INNER JOIN " & _
                        "dbo.NOMIS N1 ON C1.CustomerNo = N1.CUSTOMERNO " & _
                        "WHERE (C1.RegionNo = " & CBRegion.SelectedValue & ") AND " & _
                        "((SELECT COUNT(PC) FROM ForecastB WHERE PC = N1.PC AND RegionNo = C1.RegionNo) = 0) " & _
                        " AND N1.PC IS NOT NULL AND N1.PC <> '' " & _
                        "GROUP BY N1.PC, ProductDesc, PrimaryCat"
    End Function

    '
    '|----------------------------------------------------------------------------------------------|
    '| CreateHeaderForecast: Cr�e l'ent�te du tableau forecasts                                     |
    '|----------------------------------------------------------------------------------------------|
    Sub CreateHeaderForecast()
        Dim DGForecastsHeader As Table = FindControl("DGForecastsHeader")
        Dim header As New TableRow
        Dim ColumnName As TableCell

        Dim strNomColonnes(9) As String
        Dim intLargeurColonnes(9) As Integer
        Dim i As Integer

        strNomColonnes(0) = "PC"
        strNomColonnes(1) = "PC - Description"
        strNomColonnes(2) = "FY" & Mid(FiscalYear() - 2, 3, 2)
        strNomColonnes(3) = "FY" & Mid(FiscalYear() - 1, 3, 2)
        strNomColonnes(4) = "FY" & Mid(FiscalYear(), 3, 2)
        strNomColonnes(5) = "FY" & Mid(FiscalYear(), 3, 2) & " Prediction"
        strNomColonnes(6) = "FY" & Mid(FiscalYear() - 2, 3, 2) & "- FY" & Mid(FiscalYear(), 3, 2) & " Avg."
        strNomColonnes(7) = "FY" & Mid(FiscalYear(), 3, 2) & " Forecast"
        strNomColonnes(8) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Forecast"

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
            ColumnName = New TableCell
            ColumnName.Text = strNomColonnes(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            If i >= 2 Then
                ColumnName.HorizontalAlign = HorizontalAlign.Center
            End If
            header.Cells.Add(ColumnName)
        Next
        DGForecastsHeader.Controls.Add(header)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| CreateHeaderInitiatives: Cr�e l'ent�te du tableau Initiatives                                |
    '|----------------------------------------------------------------------------------------------|
    Sub CreateHeaderInitiatives()
        Dim DGInitiativesHeader As Table = FindControl("DGInitiativesHeader")
        Dim header As New TableRow
        Dim ColumnName As TableCell

        Dim strNomColonnes(8) As String
        Dim intLargeurColonnes(8) As Integer
        Dim i As Integer

        strNomColonnes(0) = "Division"
        strNomColonnes(1) = "Initiative"
        strNomColonnes(2) = "FY" & Mid(FiscalYear, 3, 2) & " Completed"
        strNomColonnes(3) = "FY" & Mid(FiscalYear, 3, 2) & " Planned"
        strNomColonnes(4) = "FY" & Mid(FiscalYear, 3, 2) & " Notes"
        strNomColonnes(5) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Completed"
        strNomColonnes(6) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Planned"
        strNomColonnes(7) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Notes"

        intLargeurColonnes(0) = 120
        intLargeurColonnes(1) = 190
        intLargeurColonnes(2) = 70
        intLargeurColonnes(3) = 70
        intLargeurColonnes(4) = 180
        intLargeurColonnes(5) = 70
        intLargeurColonnes(6) = 70
        intLargeurColonnes(7) = 180

        For i = 0 To UBound(strNomColonnes)
            ColumnName = New TableCell
            ColumnName.Text = strNomColonnes(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            ColumnName.HorizontalAlign = HorizontalAlign.Center
            header.Cells.Add(ColumnName)
        Next
        DGInitiativesHeader.Controls.Add(header)

    End Sub


    '
    '|----------------------------------------------------------------------------------------------|
    '| CreateHeaderDetails: Cr�e l'ent�te du tableau Initiatives                                |
    '|----------------------------------------------------------------------------------------------|
    Sub CreateHeaderDetails()
        Dim DGInitiativesHeader As Table = FindControl("DGDetailsHeader")
        Dim header As New TableRow
        Dim ColumnName As TableCell

        Dim strNomColonnes(9) As String
        Dim intLargeurColonnes(9) As Integer
        Dim i As Integer

        strNomColonnes(0) = "PC"
        strNomColonnes(1) = "PC - Description"
        strNomColonnes(2) = "FY" & Mid(FiscalYear() - 2, 3, 2)
        strNomColonnes(3) = "FY" & Mid(FiscalYear() - 1, 3, 2)
        strNomColonnes(4) = "FY" & Mid(FiscalYear(), 3, 2)
        strNomColonnes(5) = "FY" & Mid(FiscalYear(), 3, 2) & " Prediction"
        strNomColonnes(6) = "FY" & Mid(FiscalYear() - 2, 3, 2) & "- FY" & Mid(FiscalYear(), 3, 2) & " Avg."
        strNomColonnes(7) = "FY" & Mid(FiscalYear(), 3, 2) & " Quote Total"
        strNomColonnes(8) = "FY" & Mid(FiscalYear(), 3, 2) & " Quote Open"

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
            ColumnName = New TableCell
            ColumnName.Text = strNomColonnes(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            If i >= 2 Then
                ColumnName.HorizontalAlign = HorizontalAlign.Center
            End If
            header.Cells.Add(ColumnName)
        Next
        DGDetailsHeader.Controls.Add(header)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| CreateHeaderCustomers: Cr�e l'ent�te du tableau CustomersAB                                  |
    '|----------------------------------------------------------------------------------------------|
    Sub CreateHeaderCustomers()
        Dim DGForecastsHeader As Table = FindControl("dgCustomersABHeader")
        Dim header As New TableRow
        Dim ColumnName As TableCell

        Dim strNomColonnes(6) As String
        Dim intLargeurColonnes(6) As Integer
        Dim i As Integer

        strNomColonnes(0) = "Customer #"
        strNomColonnes(1) = "Customer Name"
        strNomColonnes(2) = "City"
        strNomColonnes(3) = "A"
        strNomColonnes(4) = "B"
        strNomColonnes(5) = "Region"

        intLargeurColonnes(0) = 80
        intLargeurColonnes(1) = 250
        intLargeurColonnes(2) = 150
        intLargeurColonnes(3) = 20
        intLargeurColonnes(4) = 20
        intLargeurColonnes(5) = 180

        For i = 0 To UBound(strNomColonnes)
            ColumnName = New TableCell
            ColumnName.Text = strNomColonnes(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            If i >= 2 Then
                ColumnName.HorizontalAlign = HorizontalAlign.Center
            End If
            header.Cells.Add(ColumnName)
        Next
        DGForecastsHeader.Controls.Add(header)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| CreateHeaderGoals: Cr�e l'ent�te du tableau Goals                                            |
    '|----------------------------------------------------------------------------------------------|
    Sub CreateHeaderGoals()
        Dim DGGoalsHeader As Table = FindControl("DGGoalsHeader")
        Dim header As New TableRow
        Dim ColumnName As TableCell

        Dim strNomColonnes(8) As String
        Dim intLargeurColonnes(8) As Integer
        Dim i As Integer

        strNomColonnes(0) = "Division"
        strNomColonnes(1) = "FY" & Mid(FiscalYear() - 2, 3, 2) & " Actual"
        strNomColonnes(2) = "FY" & Mid(FiscalYear() - 1, 3, 2) & " Actual"
        strNomColonnes(3) = "FY" & Mid(FiscalYear(), 3, 2) & " Actual"
        strNomColonnes(4) = "FY" & Mid(FiscalYear(), 3, 2) & " Forecasts"
        strNomColonnes(5) = "FY" & Mid(FiscalYear(), 3, 2) & " Goals"
        strNomColonnes(6) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Forecasts"
        strNomColonnes(7) = "FY" & Mid(FiscalYear() + 1, 3, 2) & " Goals"

        intLargeurColonnes(0) = 200
        intLargeurColonnes(1) = 90
        intLargeurColonnes(2) = 90
        intLargeurColonnes(3) = 90
        intLargeurColonnes(4) = 90
        intLargeurColonnes(5) = 90
        intLargeurColonnes(6) = 90
        intLargeurColonnes(7) = 90

        For i = 0 To UBound(strNomColonnes)
            ColumnName = New TableCell
            ColumnName.Text = strNomColonnes(i)
            ColumnName.Width = New Unit(intLargeurColonnes(i))
            ColumnName.CssClass = "BordureHeader"
            If i >= 1 Then
                ColumnName.HorizontalAlign = HorizontalAlign.Center
            End If
            header.Cells.Add(ColumnName)
        Next
        DGGoalsHeader.Controls.Add(header)

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| openProducts: ouvre la page d'ajout de produits                                              |
    '|----------------------------------------------------------------------------------------------|
    Sub openProducts(ByVal sender As Object, ByVal e As EventArgs)
        Dim CBCus As ComboBox = FindControl("Customer")
        Dim CBRegion As ComboBox = FindControl("BRegion")

        Dim strComm As String = ""
        strComm += "<script language=Javascript>window.open('Product.aspx"

        If CBCus.Enabled = True Then
            strComm += "?Cus=" & CBCus.Value
            strComm += "&OSNo=" & Session("OS")
        ElseIf CBRegion.Enabled = True Then
            strComm += "?Reg=" & CBRegion.Value
            strComm += "&OSNo=" & Session("OS")
        End If

        strComm += "', 'new','menubar=no,scrollbars=no,height=450,resizable=no,width=400');</script>"
        Response.Write(strComm)
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| EnableReportControls: Activer les liste d�pendamment du raport choisi                        |
    '|----------------------------------------------------------------------------------------------|
    Sub EnableReportControls(ByVal sender As Object, ByVal e As EventArgs)
        Dim Os As ComboBox = FindControl("OS")
        Dim Division As ComboBox = FindControl("Division")
        Dim Year1 As ComboBox = FindControl("Year1")
        Dim Year2 As ComboBox = FindControl("Year2")
        Dim RbCustomer As RadioButton = FindControl("RbCustomer")
        Dim RbRegion As RadioButton = FindControl("RbRegion")
        Dim Account As ComboBox = FindControl("Account")

        Division.Enabled = True

        RbCustomer.Enabled = False
        RbRegion.Enabled = False
        Account.Enabled = False

        Select Case Request.Form("Report")
            Case "RPT2"
                Year1.Enabled = True
                Year2.Enabled = False
            Case "RPT3"
                Year1.Enabled = True
                Year2.Enabled = True
            Case "RPT4"
                Year1.Enabled = True
                Year2.Enabled = False
                RbCustomer.Enabled = True
                RbRegion.Enabled = True
                Account.Enabled = True
            Case "RPT5"
                Year1.Enabled = False
                Year2.Enabled = False
            Case "RPT6"
                Division.Enabled = False
                Year1.Enabled = True
                Year2.Enabled = False
                RbCustomer.Enabled = True
                RbRegion.Enabled = True
                Account.Enabled = True
        End Select
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| OpenReport: Ouvre le bon rapport avec les options choisies                                   |
    '|----------------------------------------------------------------------------------------------|
    Sub OpenReport(ByVal sender As Object, ByVal e As EventArgs)
        Dim Os As ComboBox = FindControl("OS")
        Dim Division As ComboBox = FindControl("Division")
        Dim Year1 As ComboBox = FindControl("Year1")
        Dim Year2 As ComboBox = FindControl("Year2")
        Dim Account As ComboBox = FindControl("Account")
        Dim strAlert As String = ""

        Dim strComm As String = "<script language=Javascript>window.open('"

        If Division.Enabled = True And Division.Text = "" Then
            strAlert = "<script language=Javascript>alert(""You must select a division."");</script>"
        End If

        Select Case Request.Form("Report")
            Case "RPT2"
                If (Os.Text <> "" And UCase(Os.Text) <> "ALL") Or (Division.Text <> "" And UCase(Division.Text) <> "ALL") Then
                    strComm += "RPT_Bookings_Results.aspx"
                    strComm += "?OS=" & IIf(Os.Text <> "ALL", Os.Value.ToString, Nothing)
                    strComm += "&Division=" & Trim(Division.Text)
                    strComm += "&Year1=" & Trim(Year1.Text)
                Else
                    strAlert = "<script language=Javascript>alert(""You must select a salesman or a division."");</script>"
                End If
            Case "RPT3"
                If (Os.Text <> "" And UCase(Os.Text) <> "ALL") Or (Division.Text <> "" And UCase(Division.Text) <> "ALL") Then
                    strComm += "RPT_Multi_Year_Bookings_Results.aspx"
                    strComm += "?OS=" & IIf(Os.Text <> "ALL", Os.Value.ToString, Nothing)
                    strComm += "&Division=" & Trim(Division.Text)
                    strComm += "&Year1=" & Trim(Year1.Text)
                    strComm += "&Year2=" & Trim(Year2.Text)
                Else
                    strAlert = "<script language=Javascript>alert(""You must select a salesman or a division."");</script>"
                End If
            Case "RPT4"
                strComm += "RPT_Summary_View.aspx"
                strComm += "?OS=" & IIf(Os.Text <> "ALL", Os.Value.ToString, Nothing)
                strComm += "&Division=" & Trim(Division.Text)
                strComm += "&Year1=" & Trim(Year1.Text)
                If CType(FindControl("rbCustomer"), RadioButton).Checked Then
                    strComm += "&Type=A"
                Else
                    strComm += "&Type=B"
                End If
                strComm += "&Customer=" & Account.Value
            Case "RPT5"
                strComm += "RPT_Initiative_Summary.aspx"
                strComm += "?OS=" & IIf(Os.Text <> "ALL", Os.Value.ToString, Nothing)
                strComm += "&Division=" & Trim(Division.Text)
            Case "RPT6"
                If Account.Text <> "" Then
                    strComm += "RPT_Account.aspx"
                    strComm += "?OS=" & IIf(Os.Text <> "ALL", Os.Value.ToString, Nothing)
                    If CType(FindControl("rbCustomer"), RadioButton).Checked Then
                        strComm += "&Type=A"
                    Else
                        strComm += "&Type=B"
                    End If
                    strComm += "&Customer=" & Account.Value
                    strComm += "&Year1=" & Trim(Year1.Text)
                Else
                    strAlert = "<script language=Javascript>alert(""You must select a customer."");</script>"
                End If
            Case "RPT7"
                strComm += "RPT_Product_Specialists.aspx"
                strComm += "?OS=" & IIf(Os.Text <> "ALL", Os.Value.ToString, Nothing)
                strComm += "&Division=" & Trim(Division.Text)
                strComm += "&Year1=" & Trim(Year1.Text)

        End Select

        strComm += "', 'New','');</script>"

        If strAlert <> "" Then
            Response.Write(strAlert)
        Else
            Response.Write(strComm)
        End If

    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillReportLists: Remplit les listes de la page reports                                       |
    '|----------------------------------------------------------------------------------------------|
    Sub FillReportLists(ByVal dbConnSqlServer As OleDbConnection)
        Dim Os As ComboBox = FindControl("OS")
        Dim Year1 As ComboBox = FindControl("Year1")
        Dim Year2 As ComboBox = FindControl("Year2")
        Dim i As Integer

        'Fill OS
        Dim strReq As String = "Select 'ALL' As OSName, '0' AS OSNo"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        Dim dtTable As New DataTable
        cmdTable.Fill(dtTable)

        If User.IsInRole("LCLMTL\LCL_APT") Then
            strReq = "Select OSName, Employee.OSNo from Employee, NOMIS where Employee.OsNo=NOMIS.OSNo Group by OsName, Employee.OSNo order by OSNAME"
        Else
            strReq = "Select OSName, Employee.OSNo from Employee, NOMIS where Employee.OsNo=NOMIS.OSNo AND Employee.OsNo='" & Session("OS") & "' Group by OsName, Employee.OSNo order by OSNAME"
        End If

        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        If Session("OS") = "243" Then
            strReq = "Select OSName, Employee.OSNo from Employee, NOMIS where Employee.OsNo=NOMIS.OSNo AND Employee.OsNo='" & "220" & "' Group by OsName, Employee.OSNo order by OSNAME"
            cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
            cmdTable.Fill(dtTable)
        End If

        Os.AddItems(dtTable)

        Os.SelectedIndex = 0

        'Fill Years
        dtTable = New DataTable
        dtTable.Columns.Add(New DataColumn("YEAR"))

        Dim newRow As DataRow
        Dim FiscalYear As Integer = Format(Now, "yyyy")
        If Now >= CDate("1-oct-" & Format(Now, "yyyy")) Then FiscalYear += 1

        For i = FiscalYear - 3 To FiscalYear
            newRow = dtTable.NewRow()
            newRow(0) = i
            dtTable.Rows.Add(newRow)
        Next

        Year1.AddItems(dtTable)
        Year1.SelectedIndex = dtTable.Rows.Count - 1
        Year2.AddItems(dtTable)
        Year2.SelectedIndex = dtTable.Rows.Count - 1

        'Fill Accounts
        FillAccountList(dbConnSqlServer)
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillDivisions: Remplit la liste des divisions de la page reports                             |
    '|----------------------------------------------------------------------------------------------|
    Sub FillDivisions(ByVal dbConnSqlServer As OleDbConnection)
        'Fill Divisions
        Dim Os As ComboBox = FindControl("OS")
        Dim Divisions As ComboBox = FindControl("Division")
        Dim strReq As String
        Dim cmdTable As New OleDbDataAdapter
        Dim dtTable As New DataTable

        If UCase(Trim(Os.Text)) = "ALL" And Not User.IsInRole("LCLMTL\LCL_APT") Then
            strReq = "Select Division from Reports " & _
                    " where OsNo = '" & Session("OS") & "' order by Division"
        Else
            strReq = "Select 'All' AS PrimaryCat"
            cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
            cmdTable.Fill(dtTable)

            strReq = "Select PrimaryCat from NOMIS " & _
                    " where PrimaryCat <> '' AND FY IS NOT NULL AND FY >= '" & FiscalYear() - 2 & _
                    "' group by PrimaryCat order by PrimaryCat"
        End If

        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)
        Divisions.AddItems(dtTable)
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillAccount: Remplit la liste pour le rapport Account                                        |
    '|----------------------------------------------------------------------------------------------|
    Sub FillAccount(ByVal sender As Object, ByVal e As EventArgs)
        Dim dbConnSqlServer As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSqlServer.Open()
        FillAccountList(dbConnSqlServer)
        dbConnSqlServer.Close()
    End Sub

    '
    '|----------------------------------------------------------------------------------------------|
    '| FillAccountList: Remplit la liste pour le rapport Account d�pendamment du choix effectu�     |
    '|----------------------------------------------------------------------------------------------|
    Sub FillAccountList(ByVal dbConnSqlServer As OleDbConnection)
        Dim Account As ComboBox = FindControl("Account")
        Dim rbCustomer As RadioButton = FindControl("rbCustomer")
        Dim rbRegion As RadioButton = FindControl("rbRegion")

        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable

        If rbCustomer.Checked Then
            strReq = "Select CustomerName + ', ' + CAST(Customers.CustomerNo AS VARCHAR) as Name, Customers.CustomerNo from Customers, NOMIS " & _
                    "where Customers.CustomerNo = NOMIS.CustomerNo " & _
                    IIf(Not User.IsInRole("LCLMTL\LCL_APT"), "and NOMIS.OsNo='" & Session("OS") & "' ", "") & _
                    "Group by CustomerName, Customers.CustomerNo  " & _
                    " order by CustomerName"
        Else
            strReq = "Select Region as Name, RegionNo as No from Regions " & _
                    IIf(Not User.IsInRole("LCLMTL\LCL_APT"), "where OsNo='" & Session("OS") & "' ", "") & _
                    " order by Region"
        End If

        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        Account.AddItems(dtTable)
    End Sub

    '
    '|--------------------------------------------------------------------------------------------------------------|
    '| AddChangedForecast                                                                                           |
    '| Ajoute temporairement les forecasts qui ont �t� modifi�s jusqu'� ce quon appuie sur Save & Close             |
    '|--------------------------------------------------------------------------------------------------------------|
    Sub AddChangedForecast()
        Dim ForecastsChangedField As HtmlInputHidden = FindControl("ForecastChanged")
        Dim ForecastsChanged() As String = Split(ForecastsChangedField.Value, ";")

        Dim NoPC As String
        Dim field As String
        Dim CusType As String
        Dim NoCusReg As String
        Dim forecast As String

        Dim Insert As String

        Dim i As Integer

        For i = 0 To UBound(ForecastsChanged)
            If ForecastsChanged(i) <> "" Then
                NoPC = ""
                field = ""
                CusType = ""
                NoCusReg = ""

                Insert = "("
                NoPC = Mid(ForecastsChanged(i), 2, ForecastsChanged(i).Length - 2)

                If Session("CustomerSelected") <> Nothing And Session("CustomerSelected") <> "" Then
                    CusType = "A"
                    NoCusReg = Session("CustomerSelected")
                ElseIf Session("RegionSelected") <> Nothing And Session("RegionSelected") <> "" Then
                    CusType = "B"
                    NoCusReg = Session("RegionSelected")
                End If

                Insert += NoCusReg & ", "

                If CusType = "A" Then
                    Insert += "'" & Session("OS") & "', "
                End If

                Dim year As Integer = Format(Now, "yyyy")
                If Now >= CDate("1-jan-" & Format(Now, "yyyy")) And Now < CDate("01-jan-" & FiscalYear() + 1) Then year += 1
                Insert += "'" & year & "', "

                If Now >= CDate("01-jan-" & FiscalYear()) And Now < CDate("01-jan-" & FiscalYear()) Then
                    field = "NextYearForecast_"
                Else
                    field = "NextYearForecast_"
                End If

                Insert += "'" & NoPC & "', "
                forecast = Request.Form(field & NoPC & "_" & NoCusReg & "_" & CusType)
                Insert += IIf(forecast <> Nothing, forecast, "0")
                Insert += ");"

                If Session("CustomerSelected") <> Nothing And Session("CustomerSelected") <> "" Then
                    Session("ForecastChangedA") += Insert
                ElseIf Session("RegionSelected") <> Nothing And Session("RegionSelected") <> "" Then
                    Session("ForecastChangedB") += Insert
                End If
            End If
        Next

        ForecastsChangedField.Value = ""
    End Sub

    '
    '|--------------------------------------------------------------------------------------------------------------|
    '| AddChangedInitiatives                                                                                        |
    '| Ajoute temporairement les initiatives qui ont �t� modifi�s jusqu'� ce quon appuie sur Save & Close           |
    '|--------------------------------------------------------------------------------------------------------------|
    Sub AddChangedInitiatives()
        Dim InitiativeChangedField As HtmlInputHidden = FindControl("InitiativeChanged")
        Dim InitiativeChanged() As String = Split(InitiativeChangedField.Value, ";")

        Dim columns1() As String = {"Completed1", "Planned1", "Notes1"}
        Dim columns2() As String = {"Completed2", "Planned2", "Notes2"}

        Dim NoInitiative As String
        Dim field As String
        Dim CusType As String
        Dim NoCusReg As String
        Dim FY As String
        Dim Fields(2) As String ' Completed, planned, notes results

        Dim Insert As String

        Dim i As Integer
        Dim i2 As Integer

        For i = 0 To UBound(InitiativeChanged)
            If InitiativeChanged(i) <> "" Then
                NoInitiative = Mid(InitiativeChanged(i), 2, InitiativeChanged(i).Length - 2)

                For i2 = 0 To UBound(Fields)
                    Fields(i2) = ""
                Next

                If Session("CustomerSelected") <> Nothing And Session("CustomerSelected") <> "" Then
                    CusType = "A"
                    NoCusReg = Session("CustomerSelected")
                ElseIf Session("RegionSelected") <> Nothing And Session("RegionSelected") <> "" Then
                    CusType = "B"
                    NoCusReg = Session("RegionSelected")
                End If

                Dim year As Integer = Format(Now, "yyyy")
                If Now >= CDate("1-oct-" & Format(Now, "yyyy")) Then year += 1
                FY = "'" & year & "'"

                For i2 = 0 To UBound(columns1)

                    If Now >= CDate("01-jan-" & FiscalYear()) And Now < CDate("01-dec-" & FiscalYear()) Then
                        Fields(i2) = Request.Form(columns2(i2) & "_" & NoInitiative & "_" & NoCusReg & "_" & CusType)
                    Else
                        Fields(i2) = Request.Form(columns1(i2) & "_" & NoInitiative & "_" & NoCusReg & "_" & CusType)
                    End If

                    If Now >= CDate("01-jan-" & FiscalYear()) And Now <= CDate("01-dec-" & FiscalYear()) Then
                        FY = "'" & FiscalYear() + 1 & "'"
                        Fields(i2) = Request.Form(columns2(i2) & "_" & NoInitiative & "_" & NoCusReg & "_" & CusType)
                    Else
                        FY = "'" & FiscalYear() & "'"
                        Fields(i2) = Request.Form(columns1(i2) & "_" & NoInitiative & "_" & NoCusReg & "_" & CusType)
                    End If

                    Fields(i2) = IIf(Fields(i2) = Nothing, "", Fields(i2))

                Next

                Insert = "("
                Insert += NoCusReg & ", "

                If CusType = "A" Then
                    Insert += "'" & Session("OS") & "', "
                End If

                Insert += FY & ", "
                Insert += NoInitiative & ", "

                For i2 = 0 To UBound(Fields)
                    Insert += "'" & Replace(Fields(i2), "'", "''") & "', "
                Next

                Insert = Mid(Insert, 1, Insert.Length - 2) & ");"

                If Session("CustomerSelected") <> Nothing And Session("CustomerSelected") <> "" Then
                    Session("InitiativeChangedA") += Insert
                ElseIf Session("RegionSelected") <> Nothing And Session("RegionSelected") <> "" Then
                    Session("InitiativeChangedB") += Insert
                End If
            End If
        Next

        InitiativeChangedField.Value = ""
    End Sub

    '
    '|--------------------------------------------------------------------------------------------------------------|
    '| AddChangedGoals                                                                                              |
    '| Ajoute temporairement les Goals qui ont �t� modifi�s jusqu'� ce quon appuie sur Save & Close                 |
    '|--------------------------------------------------------------------------------------------------------------|
    Sub AddChangedGoals()
        Dim GoalsChangedField As HtmlInputHidden = FindControl("GoalsChanged")
        Dim GoalsChanged() As String = Split(GoalsChangedField.Value, ";")

        Dim Division As String
        Dim field As String
        Dim CusType As String
        Dim NoOs As String
        Dim FY As String
        Dim Goal As String

        Dim Insert As String

        Dim i As Integer
        Dim i2 As Integer

        For i = 0 To UBound(GoalsChanged)
            If GoalsChanged(i) <> "" Then
                For i2 = 1 To 2
                    Division = Trim(Mid(GoalsChanged(i), 2, GoalsChanged(i).Length - 2))
                    NoOs = Session("OSGoals")
                    If Request.Form("GOALS" & i2 & "_" & Division & "_" & NoOs) <> Nothing Then
                        FY = FiscalYear() + i2 - 1
                        Goal = Request.Form("Goals" & i2 & "_" & Division & "_" & NoOs)

                        Insert = "("
                        Insert += "'" & NoOs & "', "
                        Insert += "'" & Division & "', "
                        Insert += "'" & FY & "', "
                        Insert += Goal
                        Insert += ");"

                        Session("GoalsChanged") += Insert
                    End If
                Next
            End If
        Next

        GoalsChangedField.Value = ""

    End Sub

    '
    '|--------------------------------------------------------------------------------------------------------------|
    '| FillOsList: Remplit la liste des os                                                                          |
    '|--------------------------------------------------------------------------------------------------------------|
    Sub FillOsList(ByVal dbConnSqlServer As OleDbConnection)
        Dim CBOS As ComboBox = FindControl("CBOS")
        Dim strReq As String = "Select OSName,Employee.OsNo from Employee, Nomis " & _
                                "where Employee.OSNo=Nomis.OsNo group by Employee.OsNo, OsName order by OsName"
        Dim cmdTable As New OleDbDataAdapter(strReq, dbConnSqlServer)
        Dim dtTable As New DataTable
        cmdTable.Fill(dtTable)
        CBOS.AddItems(dtTable)
    End Sub


    Sub printInformation(ByVal sender As Object, ByVal e As EventArgs)
        Dim Customer As ComboBox = FindControl("Customer")
        Dim BRegion As ComboBox = FindControl("BRegion")
        Dim Division As ComboBox = FindControl("division")
        Dim CustomerAB As String = ""
        Dim Type As String

        Dim strAlert As String = ""

        Dim strComm As String = "<script language=Javascript>window.open('"

        If Customer.Enabled = True And Customer.Text <> "" Then
            CustomerAB = Customer.Value
            Type = "A"
        ElseIf BRegion.Enabled = True And BRegion.Text <> "" Then
            CustomerAB = BRegion.Value
            Type = "B"
        End If

        If CustomerAB = "" Then
            strAlert = "<script language=Javascript>alert(""You must select a customer."");</script>"
        End If

        If Session("ForecastEntry") = True Then
            strComm += "PrintForecasts.aspx"
        ElseIf Session("InitiativesEntry") = True Then
            strComm += "PrintInitiatives.aspx"
        Else
            strComm += "PrintExtraDetails.aspx"
        End If

        strComm += "?OS=" & Session("OS")
        strComm += "&CusAB=" & Trim(CustomerAB)
        strComm += "&type=" & Type
        strComm += "&Division=" & Trim(Division.Text)

        strComm += "', 'New','');</script>"

        If strAlert <> "" Then
            Response.Write(strAlert)
        Else
            Response.Write(strComm)
        End If
    End Sub


End Class