Imports System.Data
Imports System.Data.OleDb

Public Class Product
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents Description As System.Web.UI.WebControls.Label
    Protected WithEvents Forecast As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents CustomerNo As System.Web.UI.WebControls.Label
    Protected WithEvents CustomerName As System.Web.UI.WebControls.Label
    Protected WithEvents City As System.Web.UI.WebControls.Label
    Protected WithEvents OsNo As System.Web.UI.WebControls.Label
    Protected WithEvents OsName As System.Web.UI.WebControls.Label
    Protected WithEvents RegionC As System.Web.UI.WebControls.Label
    Protected WithEvents RegionNo As System.Web.UI.WebControls.Label

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
        Dim CBDivision As ComboBox = FindControl("Division")

        Dim strReq As String
        Dim cmdTable As New OleDbDataAdapter
        Dim dtTable As New DataTable

        dbConnSqlServer.Open()

        If Not Page.IsPostBack Then
            Session("ProductDivision") = Nothing
            If Request.QueryString("Cus") <> Nothing Then
                Session("CusTypeProducts") = "CUS"
                ShowCustomerInfos(dbConnSqlServer)
            ElseIf Request.QueryString("Reg") <> Nothing Then
                Session("CusTypeProducts") = "REG"
                ShowRegionInfos(dbConnSqlServer)
            End If

            If Request.QueryString("OSNo") <> Nothing Then
                ShowOsInfos(dbConnSqlServer)
            End If

            ShowDivisions(dbConnSqlServer)
        Else
            If CBDivision.Text <> "" And CBDivision.Text <> Session("ProductDivision") Then
                ShowProducts(dbConnSqlServer)
                Session("ProductDivision") = CBDivision.Text
            End If
        End If

        dbConnSqlServer.Close()
    End Sub

    Sub ShowCustomerInfos(ByVal dbConnSqlServer As OleDbConnection)
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable

        strReq = "Select CustomerNo, CustomerName, city from Nomis where customerNo=" & Request.QueryString("Cus") & _
                        " Group by CustomerNo, CustomerName, city"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        CustomerNo.Text = dtTable.Rows(0)(0)
        CustomerName.Text = dtTable.Rows(0)(1)
        City.Text = dtTable.Rows(0)(2)
    End Sub

    Sub ShowRegionInfos(ByVal dbConnSqlServer As OleDbConnection)
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable

        strReq = "Select * from Regions where RegionNo=" & Request.QueryString("Reg")
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        RegionC.Text = dtTable.Rows(0)(0)
        RegionNo.Text = Request.QueryString("Reg")
    End Sub

    Sub ShowOsInfos(ByVal dbConnSqlServer As OleDbConnection)
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable

        strReq = "Select OsName from Nomis where OsNo='" & Request.QueryString("OSNo") & "'"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        OsNo.Text = Request.QueryString("OSNo")
        OsName.Text = dtTable.Rows(0)(0)
    End Sub

    Sub ShowDivisions(ByVal dbConnSqlServer As OleDbConnection)
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim CBDivision As ComboBox = FindControl("Division")

        strReq = "Select PRIMARYCAT as Division From PCNOMIS Group by PRIMARYCAT order by PRIMARYCAT"
        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        CBDivision.AddItems(dtTable)
    End Sub

    Sub ShowProducts(ByVal dbConnSqlServer As OleDbConnection)
        Dim strReq As String
        Dim cmdTable As OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim CBDivision As ComboBox = FindControl("Division")
        Dim CBPc As ComboBox = FindControl("ProductCode")

        If Session("CusTypeProducts") = "CUS" Then
            strReq = "Select (PC + ', ' + PRODUCTDESC), PC From PCNOMIS N1 WHERE PrimaryCat='" & CBDivision.Text & "' AND NOT EXISTS " & _
                        "(Select PC from ForecastA where N1.PC = PC AND OsNo= " & OsNo.Text & " AND CustomerNo=" & CustomerNo.Text & " Group by PC " & _
                        "Union Select PC From Nomis where N1.PC = PC AND OsNo= " & OsNo.Text & " AND CustomerNo=" & CustomerNo.Text & ") " & _
                        " Group by (PC + ', ' + PRODUCTDESC), PC order by PC"
        ElseIf Session("CusTypeProducts") = "REG" Then
            strReq = "Select (PC + ', ' + PRODUCTDESC), PC From PCNOMIS N1 WHERE PrimaryCat='" & CBDivision.Text & "' AND NOT EXISTS " & _
                        "(Select PC from ForecastB where N1.PC = PC AND RegionNo= " & RegionNo.Text & " Group by PC) Order by PC"
            '"Union Select PC From Nomis, Customers where N1.PC = PC AND Nomis.OSNO= " & OsNo.Text & " AND Nomis.CustomerNo = Customers.CustomerNo And RegionNo=" & RegionNo.Text & ") " & _
            '" Group by (PC + ', ' + PRODUCTDESC), PC order by PC"

        End If

        Trace.Warn(strReq)

        cmdTable = New OleDbDataAdapter(strReq, dbConnSqlServer)
        cmdTable.Fill(dtTable)

        CBPc.AddItems(dtTable)
    End Sub

    Sub AddProduct(ByVal sender As Object, ByVal e As EventArgs)
        Dim strComm As String = "<script language=Javascript>opener.document.forms[0].submit();self.close();</script>"
        Dim strReq As String = ""
        Dim dbConnSqlServer As OleDbConnection = EtablitConnexionSQLServer()
        Dim cmdCommand As New OleDbCommand
        Dim FY As String
        Dim cbPC As ComboBox = FindControl("ProductCode")

        FY = IIf(Now >= CDate("01-oct-" & FiscalYear() - 1) And Now <= CDate("31-jul-" & FiscalYear()), FiscalYear, FiscalYear() + 1)

        If ValidPage() Then
            'Enregistrer dans la bd et reloader la page
            dbConnSqlServer.Open()
            cmdCommand.Connection = dbConnSqlServer

            If Session("CusTypeProducts") = "CUS" Then
                strReq = "Insert into ForecastA Values("
                strReq += CustomerNo.Text & ", "
                strReq += "'" & OsNo.Text & "', "
                strReq += "'" & FY & "', "
                strReq += "'" & cbPC.Value & "', "
                strReq += Forecast.Text
                strReq += ")"
            ElseIf Session("CusTypeProducts") = "REG" Then
                strReq = "Insert into ForecastB Values("
                strReq += RegionNo.Text & ", "
                strReq += "'" & FY & "', "
                strReq += "'" & cbPC.Value & "', "
                strReq += Forecast.Text
                strReq += ")"
            End If

            Trace.Warn(strReq)
            cmdCommand.CommandText = strReq
            cmdCommand.ExecuteNonQuery()

            dbConnSqlServer.Close()
            Response.Write(strComm)
        End If

    End Sub
    Function ValidPage() As Boolean
        Dim cbPC As ComboBox = FindControl("ProductCode")
        Dim strError As String = ""
        Dim valid As Boolean = True

        If cbPC.Value = Nothing Or cbPC.Text = "" Then
            strError = "Please select a Product Code"
            valid = False
        ElseIf Forecast.Text = "" Then
            strError = "Please enter a valid forecast"
            valid = False
        ElseIf Not IsNumeric(Forecast.Text) Then
            strError = "Please enter a numeric forecast"
            valid = False
        End If

        If Not valid Then
            Response.Write("<script language=Javascript>alert(""" & strError & """);</script>")
        End If

        ValidPage = valid
    End Function

    '|----------------------------------------------------------------------------------------------|
    '| FiscalYear:                                                                                  |
    '|----------------------------------------------------------------------------------------------|
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

End Class
