Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
Imports System.Web.UI.WebControls

Public Class ComboBox
    Inherits System.Web.UI.UserControl

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents listBox As System.Web.UI.WebControls.ListBox
    Protected WithEvents Texte As System.Web.UI.WebControls.TextBox
    Protected WithEvents Arrow As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents DivList As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents TextisChanged As System.Web.UI.HtmlControls.HtmlInputHidden

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
        If Not listBox.SelectedItem Is Nothing Then
            Texte.Text = listBox.SelectedItem.Text
        End If
        If Width = "" Then
            Width = "50"
        End If
    End Sub

    Public Property CssClass() As String
        Get
            Return listBox.CssClass
        End Get
        Set(ByVal Value As String)
            Texte.CssClass = Server.HtmlEncode(Value)
            listBox.CssClass = Server.HtmlEncode(Value)
        End Set
    End Property

    Public WriteOnly Property AutoPostBack() As Boolean
        Set(ByVal Value As Boolean)
            If Value Then
                listBox.AutoPostBack = True
                Texte.Attributes.Add("onBlur", "javascript:PostBack();")
            Else
                Texte.Attributes.Add("onBlur", "")
                listBox.AutoPostBack = False
            End If
        End Set
    End Property

    Public ReadOnly Property Text() As String
        Get
            If Not listBox.SelectedItem Is Nothing Then
                Return listBox.SelectedItem.Text
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property Value() As String
        Get
            Return listBox.SelectedValue
        End Get
    End Property

    Public WriteOnly Property SelectedIndex() As Integer
        Set(ByVal Value As Integer)
            If Value < listBox.Items.Count Then
                listBox.SelectedIndex = Value
            End If
        End Set
    End Property

    Public Property SelectedValue() As String
        Get
            Return listBox.SelectedValue
        End Get
        Set(ByVal Value As String)
            listBox.SelectedValue = Value
        End Set
    End Property

    Public Property Width() As String
        Get
            Return listBox.Width.Value
        End Get
        Set(ByVal Value As String)
            Texte.Style.Add("width", (Server.HtmlEncode(Value) - 15).ToString)
            listBox.Style.Add("width", Server.HtmlEncode(Value).ToString)
        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Return Texte.Enabled
        End Get
        Set(ByVal Value As Boolean)
            Texte.Enabled = Value
            Texte.CssClass = IIf(Value, CssClass, "Enabled")
            Arrow.Src = "Images/Arrow" & IIf(Value, "", "E") & ".jpg"
            Arrow.Disabled = Not Value
        End Set
    End Property

    Public Sub AddItems(ByVal dtTable As DataTable)
        listBox.Items.Clear()
        Texte.Text = ""
        Dim myRow As DataRow

        For Each myRow In dtTable.Rows
            If dtTable.Columns.Count = 1 Then
                If Not IsDBNull(myRow(0)) Then
                    listBox.Items.Add(myRow(0))
                End If
            Else
                If Not IsDBNull(myRow(0)) And Not IsDBNull(myRow(1)) Then
                    listBox.Items.Add(New ListItem(myRow(0).ToString(), myRow(1).ToString()))
                End If
            End If
        Next

        If listBox.Items.Count <= 1 Then
            listBox.Rows = 2
        ElseIf listBox.Items.Count < 10 Then
            listBox.Rows = listBox.Items.Count
        Else
            listBox.Rows = 10
        End If
    End Sub

    Public Sub AddItems(ByVal dtView As DataView, ByVal TextField As String, ByVal TextValue As String)
        listBox.Items.Clear()
        Texte.Text = ""
        Dim myRow As DataRow

        listBox.DataSource = dtView
        listBox.DataTextField = TextField
        listBox.DataValueField = TextValue
        listBox.DataBind()

        If listBox.Items.Count <= 1 Then
            listBox.Rows = 2
        ElseIf listBox.Items.Count < 10 Then
            listBox.Rows = listBox.Items.Count
        Else
            listBox.Rows = 10
        End If
    End Sub

End Class
