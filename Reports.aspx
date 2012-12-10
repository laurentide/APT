<%@ Register TagPrefix="CC" TagName="Combobox" Src="Combobox.ascx" %>
<%@ Page Language="vb" Codebehind="APT.aspx.vb" Inherits="APT.APT" clientTarget="downlevel" %>
<%@ import Namespace="System.Data.OleDb" %>
<%@ import Namespace="System.Data" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
		<title>Reports</title>
<LINK href="styles.css" type=text/css rel=stylesheet >
<script language=vb runat="server">
			sub Page_Load()
				if Session("OS") <> Nothing or User.IsInRole("LCLMTL\LCL_APT") then
					Dim dbConnSqlServer as oleDbConnection = EtablitConnexionSQLServer()
					dbConnSqlServer.open
					
					if not page.ispostback then
						Session("OSReport") = ""
						FillReportLists(dbConnSqlServer)
						FillDivisions(dbConnSqlServer)
						
						Year2.Enabled = False
						Account.Enabled = False
						
						If Not User.IsInRole("LCLMTL\LCL_APT") Then
							os.SelectedIndex = 1
						End If
					end if
					
					if Session("OSReport") <> Os.text then
						Session("OSReport") = Os.text
						FillDivisions(dbConnSqlServer)
					end if
					
					
					dbConnSqlServer.close
				else
					response.redirect("Denied.html")
				end if
			end sub
			</script>
</HEAD>
<body bottomMargin=0 bgColor=whitesmoke leftMargin=0 topMargin=0 rightMargin=0 marginwidth="0" marginheight="0">
<form id=Form1 runat="server">
<table height="100%" cellSpacing=0 width="100%" align=center>
  <tr>
    <td>
      <table cellSpacing=0 width="100%">
        <tr>
          <td background=images/TitleFill.gif></td>
          <td width=680 background=images/title.gif height=95 
           border="0"></td>
          <td background=images/TitleFill.gif 
        ></td></tr></table></td></tr>
  <tr>
    <td background=images/body.gif height="100%" 
border="0">
						<!-- Contenu -->
      <table style="LEFT: 40px; POSITION: relative" cellSpacing=20 
      >
        <tr vAlign=top>
          <td class=TexteRapport>Bookings</td>
          <td><asp:radiobutton class=LiensRapports id=RPT2 runat="server" Checked="True" Text="Forecasts-Bookings-Quotes" AutoPostBack="True" OnCheckedChanged="EnableReportControls" GroupName="Report"></asp:radiobutton><br 
            ><asp:radiobutton class=LiensRapports id=RPT3 runat="server" Text="Multi-Year Bookings Results by Salesperson / Customer / Product" AutoPostBack="True" OnCheckedChanged="EnableReportControls" GroupName="Report"></asp:radiobutton><br 
            ><asp:radiobutton class=LiensRapports id=RPT4 runat="server" Text="Summary View by Division" AutoPostBack="True" OnCheckedChanged="EnableReportControls" GroupName="Report"></asp:radiobutton><br 
            ></td></tr>
        <tr vAlign=top>
          <td class=TexteRapport>Initiatives</td>
          <td><asp:radiobutton class=LiensRapports id=RPT5 runat="server" Text="Initiative Summary" AutoPostBack="True" OnCheckedChanged="EnableReportControls" GroupName="Report"></asp:radiobutton><br 
            ></td></tr>
        <tr vAlign=top>
          <td class=TexteRapport>Accounts</td>
          <td><asp:radiobutton class=LiensRapports id=RPT6 runat="server" Text="Account" AutoPostBack="True" OnCheckedChanged="EnableReportControls" GroupName="Report"></asp:radiobutton><br 
            ></td></tr>
        <tr vAlign=top>
          <td class=TexteRapport>Products</td>
          <td><asp:radiobutton class=LiensRapports id=RPT7 runat="server" Text="Product specialists" AutoPostBack="True" OnCheckedChanged="EnableReportControls" GroupName="Report"></asp:radiobutton><br 
            ></td></tr>
        <tr>
          <td></td>
          <td>
            <table class="TableauRapports espaceCellules" 
              >
              <tr>
                <td class="Texte Gras Espaces">OS:</td>
                <td><cc:combobox id=OS runat="server" width="150" cssClass="Texte Combobox" AutoPostback="true"></cc:combobox></td>
                <td class="Texte Gras Espaces">Year:</td>
                <td><cc:combobox id=Year1 runat="server" width="75" cssClass="Texte Combobox"></cc:combobox></td></tr>
              <tr>
                <td class="Texte Gras Espaces" 
                >Division:</td>
                <td><cc:combobox id=Division runat="server" width="150" cssClass="Texte Combobox"></cc:combobox></td>
                <td class="Texte Gras Espaces">End 
                Year:</td>
                <td><cc:combobox id=Year2 runat="server" width="75" cssClass="Texte Combobox"></cc:combobox></td></tr>
              <tr>
                <td align=left colSpan=4>
                  <table>
                    <tr>
                      <td class="Texte Gras" 
                        >&nbsp;&nbsp;&nbsp;Customer:</td>
                      <td><asp:radiobutton id=rbCustomer Checked="True" AutoPostBack="True" OnCheckedChanged="FillAccount" GroupName="CustomerOrRegion" Runat="server" Enabled="False"></asp:radiobutton></td>
                      <td class="Texte Gras ">Region:</td>
                      <td><asp:radiobutton id=rbRegion AutoPostBack="True" OnCheckedChanged="FillAccount" GroupName="CustomerOrRegion" Runat="server" Enabled="False"></asp:radiobutton>&nbsp;&nbsp;&nbsp;&nbsp; 
                      </td>
                      <td><cc:combobox id=Account runat="server" width="270" cssClass="Texte Combobox"></cc:combobox></td></tr></table></td></tr></table></td></tr>
        <tr vAlign=bottom>
          <td align=right width=* colSpan=2><asp:button class=Bouton id=Button1 onclick=OpenReport Text="OK" Runat="server"></asp:button><asp:button class=BoutonGrand id=Button2 onclick=RedirectMenu Text="Back to Main Menu" Runat="server"></asp:button></td></tr></table></td></tr></table></form>
	</body>
</HTML>
