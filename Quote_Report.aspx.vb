Imports System.Data.OleDb
Imports System.Data

Public Class Quote_Report
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BillName As System.Web.UI.WebControls.TextBox
    Protected WithEvents CustNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents PC As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents StartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents EndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents QuoteNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents FUStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents FUEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents City As System.Web.UI.WebControls.TextBox
    Protected WithEvents LastName As System.Web.UI.WebControls.TextBox
    Protected WithEvents ModelNumber As System.Web.UI.WebControls.TextBox
    Protected WithEvents NetPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents ValDate As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents Customvalidator1 As System.Web.UI.WebControls.CustomValidator

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

        Dim dbConnSQLSERVER As OleDbConnection = EtablitConnexionSQLServer()
        dbConnSQLSERVER.Open()

        If Session("OS") = Nothing Then
            GetSessionOs(dbConnSQLSERVER)
        End If

        If Session("OS") = Nothing And Not User.IsInRole("LCLMTL\LCL_APT") _
           And Not User.IsInRole("LCLMTL\LCL_AE") And Not User.IsInRole("LCLMTL\LCL_SA") And Not User.IsInRole("LCLMTL\LCL_SIE") Then
            Response.Redirect("Denied.html")
        Else
            'Remplit le combobox des os
            Dim CB As ComboBox = FindControl("OsNo")
            Dim CBQuotedBy As ComboBox = FindControl("QuotedBy")
            Dim CBWFollowUp As ComboBox = FindControl("WFollowUp")
            Dim CBStatus As ComboBox = FindControl("Status")
            Dim CBGrouping As ComboBox = FindControl("Grouping")
            Dim strReq As String
            Dim cmdTable As OleDbDataAdapter
            Dim dtTable As New DataTable
            Dim dtTable2 As New DataTable
            Dim dtTable3 As New DataTable
            Dim dtTable4 As New DataTable
            Dim dtTable5 As New DataTable

            If CB.Text = "" Then
                If User.IsInRole("LCLMTL\LCL_APT") Or User.IsInRole("LCLMTL\lcl_AE") Or User.IsInRole("LCLMTL\lcl_SA") Then
                    strReq = "Select '' As OS, 0 AS OSNo"
                    cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                    cmdTable.Fill(dtTable)
                    strReq = "Select (OsNo + ' - ' + WINDOWSNAME) As OS, OSNo from Employee "
                Else
                    strReq = "Select (OsNo + ' - ' + WINDOWSNAME) As OS, OSNo from Employee WHERE OSNo='" & Session("OS") & "'"
                End If

                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable)
                CB.AddItems(dtTable)

                If Not (User.IsInRole("LCLMTL\LCL_APT") Or User.IsInRole("LCLMTL\lcl_AE") Or User.IsInRole("LCLMTL\lcl_SA")) Then
                    CB.SelectedIndex = 0
                    CB.Enabled = False
                End If
            End If

            If CBQuotedBy.Text = "" Then
                strReq = "Select '' As QuotedBy"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable2)
                strReq = "SELECT QuotedBy FROM OSREPORT GROUP BY QuotedBy"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable2)
                CBQuotedBy.AddItems(dtTable2)
            End If

            If CBWFollowUp.Text = "" Then
                strReq = "Select '' As WhodoFollowup"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable3)
                strReq = "SELECT WhodoFollowup FROM OSREPORT GROUP BY WhodoFollowup"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable3)
                CBWFollowUp.AddItems(dtTable3)
            End If

            If CBStatus.Text = "" Then
                strReq = "Select '' As Status"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable4)
                strReq = "SELECT Status FROM vStatus"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable4)
                CBStatus.AddItems(dtTable4)
            End If

            If CBGrouping.Text = "" Then
                strReq = "Select '' As PRIMARYCAT"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable5)
                strReq = "SELECT PRIMARYCAT FROM vPrimaryCat"
                cmdTable = New OleDbDataAdapter(strReq, dbConnSQLSERVER)
                cmdTable.Fill(dtTable5)
                CBGrouping.AddItems(dtTable5)
            End If

        End If

        dbConnSQLSERVER.Close()

    End Sub

    Sub SendSearch(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strComm As String = "<script language=Javascript>"
        Dim valide As Boolean = True

        Dim OsNo As ComboBox = FindControl("OSNO")
        Dim WFollowUp As ComboBox = FindControl("WFollowUp")
        Dim QuotedBy As ComboBox = FindControl("QuotedBy")
        Dim Status As ComboBox = FindControl("Status")
        Dim Grouping As ComboBox = FindControl("Grouping")
        'Dim format As New System.Globalization.CultureInfo("en-US", True)

        If StartDate.Text <> "" And valide = True Then
            If (Not IsDate(StartDate.Text)) Then
                strComm += "alert(""Invalid Start Date Format mm/dd/yyyy"")"
                valide = False
            End If
        End If

        If EndDate.Text <> "" And valide = True Then
            If (Not IsDate(EndDate.Text)) Then
                strComm += "alert(""Invalid Start Date Format mm/dd/yyyy"")"
                valide = False
            End If
        End If

        If FUStartDate.Text <> "" And valide = True Then
            If (Not IsDate(FUStartDate.Text)) Then
                strComm += "alert(""Invalid Start Date Format mm/dd/yyyy"")"
                valide = False
            End If
        End If

        If FUEndDate.Text <> "" And valide = True Then
            If (Not IsDate(FUEndDate.Text)) Then
                strComm += "alert(""Invalid Start Date Format mm/dd/yyyy"")"
                valide = False
            End If
        End If

        If Not IsNumeric(CustNo.Text) And CustNo.Text <> "" And valide = True Then
            strComm += "alert(""Invalid Customer #"")"
            valide = False
        End If

        If Status.Text = "" And Grouping.Text = "" And ModelNumber.Text = "" And NetPrice.Text = "" And OsNo.Text = "" _
            And BillName.Text = "" And QuoteNo.Text = "" And CustNo.Text = "" And City.Text = "" _
            And LastName.Text = "" And PC.Text = "" And QuotedBy.Text = "" And WFollowUp.Text = "" And valide = True Then
            strComm += "alert(""You must enter at least one field different from dates"")"
            valide = False
        End If

        If valide = True Then
            strComm += "window.open('Quote_Report_SEARCH.aspx"
            strComm += "?QuoteNo=" & Replace(Trim(QuoteNo.Text), "'", "")
            strComm += "&StartDate=" & Replace(Trim(StartDate.Text), "'", "")
            strComm += "&EndDate=" & Replace(Trim(EndDate.Text), "'", "")
            strComm += "&FUStartDate=" & Replace(Trim(FUStartDate.Text), "'", "")
            strComm += "&FUEndDate=" & Replace(Trim(FUEndDate.Text), "'", "")
            If OsNo.SelectedValue = "" Then
                strComm += "&OsNo="
            ElseIf OsNo.SelectedValue < 100 Then 'And (Session("OS") = Nothing)
                strComm += "&OsNo=0" & OsNo.SelectedValue
            Else
                strComm += "&OsNo=" & OsNo.SelectedValue
            End If
            strComm += "&BillName=" & Replace(Trim(BillName.Text), "'", "")
            strComm += "&CustNo=" & Trim(CustNo.Text)
            strComm += "&Status=" & Replace(Trim(Status.Text), "'", "")
            strComm += "&Grouping=" & Replace(Trim(Grouping.Text), "'", "")
            strComm += "&City=" & Replace(Trim(City.Text), "'", "")
            strComm += "&PC=" & Replace(Trim(PC.Text), "'", "")
            strComm += "&ModelNumber=" & Replace(Trim(ModelNumber.Text), "'", "")
            strComm += "&NetPrice=" & Replace(Trim(NetPrice.Text), "'", "")
            strComm += "&LastName=" & Replace(Trim(LastName.Text), "'", "")
            strComm += "&WFollowUp=" & Replace(Trim(WFollowUp.Text), "'", "")
            strComm += "&QuotedBy=" & Replace(Trim(QuotedBy.Text), "'", "")

            strComm += "', 'New','');"

        End If
        strComm += "</script>"
        Response.Write(strComm)
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
    '| GetSessionOs: Retrouver le No OS si l'usager y a accès                                       |
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
    '| OSExists: Vérifie l'existence de l'usager dans la base de données                            |
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

End Class
