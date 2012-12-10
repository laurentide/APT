Imports System.Data.OleDb
Imports System.Data

Public Class CTR
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents OrderNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents Description As System.Web.UI.WebControls.TextBox
    Protected WithEvents BillName As System.Web.UI.WebControls.TextBox
    Protected WithEvents CustNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents ShipCity As System.Web.UI.WebControls.TextBox
    Protected WithEvents Orig As System.Web.UI.WebControls.TextBox
    Protected WithEvents CustPONo As System.Web.UI.WebControls.TextBox
    Protected WithEvents PC As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents StartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents EndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents List As System.Web.UI.WebControls.TextBox
    Protected WithEvents QuoteNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents CustInvItem As System.Web.UI.WebControls.TextBox
    Protected WithEvents DiscountRate As System.Web.UI.WebControls.TextBox
    Protected WithEvents SerialNumber As System.Web.UI.WebControls.TextBox
    Protected WithEvents InventoryNumber As System.Web.UI.WebControls.TextBox
    Protected WithEvents ProductCodeCategory As System.Web.UI.WebControls.TextBox

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
            GetSessionOs(dbConnSqlServer)
        End If

        If Session("OS") = Nothing And Not User.IsInRole("LCLMTL\LCL_APT") _
           And Not User.IsInRole("LCLMTL\LCL_AE") And Not User.IsInRole("LCLMTL\LCL_SA") And Not User.IsInRole("LCLMTL\LCL_SIE") Then
            Response.Redirect("Denied.html")
        Else
            'Remplit le combobox des os
            Dim CB As ComboBox = FindControl("OsNo")
            Dim strReq As String
            Dim cmdTable As OleDbDataAdapter
            Dim dtTable As New DataTable

            If CB.Text = "" Then
                If User.IsInRole("LCLMTL\LCL_APT") Or User.IsInRole("LCLMTL\lcl_AE") Or User.IsInRole("LCLMTL\lcl_SA") Or User.IsInRole("LCLMTL\LCL_SIE") Then
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

                'If Not (User.IsInRole("LCLMTL\LCL_APT") Or User.IsInRole("LCLMTL\RichardThe") Or User.IsInRole("LCLMTL\lcl_AE") Or User.IsInRole("LCLMTL\lcl_SA")) Then
                '    CB.SelectedIndex = 0
                '    CB.Enabled = False
                'End If
            End If
        End If

        dbConnSQLSERVER.Close()
    End Sub

    Sub SendSearch(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strComm As String = "<script language=Javascript>"
        Dim valide As Boolean = True

        Dim OsNo As ComboBox = FindControl("OSNO")

        If Not IsNumeric(CustNo.Text) And CustNo.Text <> "" Then
            strComm += "alert(""Invalid Customer #"")"
            valide = False
        End If

        If (Not IsNumeric(StartDate.Text) Or StartDate.Text.Length <> 7) And StartDate.Text <> "" Then
            strComm += "alert(""Invalid date"")"
            valide = False
        End If

        If (Not IsNumeric(EndDate.Text) Or EndDate.Text.Length <> 7) And EndDate.Text <> "" Then
            strComm += "alert(""Invalid date"")"
            valide = False
        End If

        'If OrderNo.Text = "" And Description.Text = "" _
        '    And OsNo.Text = "" And BillName.Text = "" And QuoteNo.Text = "" And CustInvItem.Text = "" And CustNo.Text = "" And List.Text = "" _
        '    And ShipCity.Text = "" And Orig.Text = "" And CustPONo.Text = "" And PC.Text = "" And SerialNumber.Text = "" And DiscountRate.Text = "" Then
        '    strComm += "alert(""You must enter at least one field different from dates"")"
        '    valide = False
        'End If

        If valide = True Then
            strComm += "window.open('CTRSEARCH.aspx"
            strComm += "?OrderNo=" & Replace(Trim(OrderNo.Text), "'", "")
            strComm += "&OrderDateStart=" & Replace(Trim(StartDate.Text), "'", "")
            strComm += "&OrderDateEnd=" & Replace(Trim(EndDate.Text), "'", "")
            strComm += "&Desc=" & Replace(Trim(Description.Text), "'", "")
            If OsNo.SelectedValue = "" Then
                strComm += "&OsNo="
            ElseIf OsNo.SelectedValue < 100 Then 'And (Session("OS") = Nothing)
                strComm += "&OsNo=0" & OsNo.SelectedValue
            Else
                strComm += "&OsNo=" & OsNo.SelectedValue
            End If
            strComm += "&BillName=" & Replace(Trim(BillName.Text), "'", "")
            strComm += "&CustNo=" & Trim(CustNo.Text)
            strComm += "&List=" & Replace(Trim(List.Text), "'", "")
            strComm += "&ShipCity=" & Replace(Trim(ShipCity.Text), "'", "")
            strComm += "&Orig=" & Replace(Trim(Orig.Text), "'", "")
            strComm += "&CustPoNo=" & Replace(Trim(CustPONo.Text), "'", "")
            strComm += "&PC=" & Replace(Trim(PC.Text), "'", "")
            strComm += "&QuoteNo=" & Replace(Trim(QuoteNo.Text), "'", "")
            strComm += "&CustInvItem=" & Replace(Trim(CustInvItem.Text), "'", "")
            strComm += "&DiscountRate=" & Replace(Trim(DiscountRate.Text), "'", "")
            strComm += "&SerialNumber=" & Replace(Trim(SerialNumber.Text), "'", "")
            strComm += "&InventoryNumber=" & Replace(Trim(InventoryNumber.Text), "'", "")
            strComm += "&ProductCodeCategory=" & Replace(Trim(ProductCodeCategory.text), "'", "")
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
