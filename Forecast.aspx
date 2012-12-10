<%@ Register TagPrefix="CC" TagName="Combobox" Src="Combobox.ascx" %>
<%@ Page Language="vb" Codebehind="APT.aspx.vb" Inherits="APT.APT" clientTarget="downlevel"  %>
<%@ import Namespace="System.Data.OleDb" %>
<%@ import Namespace="System.Data" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Next Year Forecast</title>
		<script language="javascript" src="APT.js"></script>
		<script language="VB" runat="server">

			Sub Page_Load()
			
				Dim dbConnSqlServer as oleDbConnection = EtablitConnexionSQLServer()
				dbConnSqlServer.open
				
				if session("OS") = Nothing then
					response.redirect("Denied.html")
				End if
				
				'Déterminer quelle section afficher
				if request.form("SalesmanPerform") <> Nothing then
					Session("ForecastEntry") = true
					SalesmanPerform.cssclass = "CurrentDiv"
					
					Session("InitiativesEntry") = false
					Initiatives.cssclass = "OtherDiv"
					
					Session("ExtraDetails") = false
					ExtraDetails.cssclass = "OtherDiv"
					
				Else if request.form("Initiatives") <> Nothing then
					Session("ForecastEntry") = False
					SalesmanPerform.cssclass = "OtherDiv"
					
					Session("InitiativesEntry") = true
					Initiatives.cssclass = "CurrentDiv"
					
					Session("ExtraDetails") = false
					ExtraDetails.cssclass = "OtherDiv"
					
				Else if request.form("ExtraDetails") <> Nothing then
					Session("ForecastEntry") = False
					SalesmanPerform.cssclass = "OtherDiv"
					
					Session("InitiativesEntry") = false
					Initiatives.cssclass = "OtherDiv"
					
					Session("ExtraDetails") = True
					ExtraDetails.cssclass = "CurrentDiv"
				End if
				
				if not page.ispostback then			
					Session("ForecastChangedA") = Nothing
					Session("ForecastChangedB") = Nothing
					
					Session("ForecastEntry") = true
					SalesmanPerform.cssclass = "CurrentDiv"
					
					Session("InitiativesEntry") = false
					Initiatives.cssclass = "OtherDiv"
					
					Session("ExtraDetails") = false
					ExtraDetails.cssclass = "OtherDiv"
					
					If session("OS") = Nothing then
						GetSessionOs(dbConnSqlServer)
					end if
					
					ListCustomerType()
					ListCustomers(dbConnSqlServer)
					ListRegions(dbConnSqlServer)
				Else
					CustomerCity.text = ""
					CustomerNo.text = ""
					AddChangedForecast()
					AddChangedInitiatives()
					AccessibleFields(dbConnSqlServer)
				End If
				
				if Session("ForecastEntry") = true then
					CreateHeaderForecast()
				Elseif Session("InitiativesEntry") = true then
					CreateHeaderInitiatives()
				Elseif Session("ExtraDetails") = true then
					CreateHeaderDetails()
				end if
				
				dbConnSqlServer.close
				
			End Sub

		</script>
		<LINK href="styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body bottomMargin="0" bgColor="whitesmoke" leftMargin="0" topMargin="0" rightMargin="0"
		marginwidth="0" marginheight="0">
		<form id="Form1" runat="server">
			<table height="100%" cellSpacing="0" width="100%" align="center">
				<tr>
					<td>
						<table cellSpacing="0" width="100%">
							<tr>
								<td background="images/TitleFill.gif"></td>
								<td width="680" background="images/title.gif" height="95" border="0"></td>
								<td background="images/TitleFill.gif"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td background="images/body.gif" height="100%" border="0">
						<!-- Contenu -->
						<table height="60%" width="100%" align="center">
							<tr>
								<td class="Gras">
									<table width="100%">
										<tr vAlign="top">
											<td>
												<table class="Texte">
													<tr>
														<td class="Gras" align="right">Forecast for:</td>
														<td><cc:combobox id="CustomerType" runat="server" cssClass="Texte Combobox" width="150" AutoPostBack="true"></cc:combobox></td>
														<td class="Gras" align="right">&nbsp;&nbsp;&nbsp;Customer:&nbsp;&nbsp;&nbsp;</td>
														<td><cc:combobox id="Customer" runat="server" cssClass="Texte Combobox" width="250" AutoPostBack="true"></cc:combobox></td>
													</tr>
													<tr>
														<td class="Gras" align="right">Region:</td>
														<td><cc:combobox id="BRegion" runat="server" cssClass="Texte Combobox" width="150" AutoPostBack="true"></cc:combobox></td>
														<td class="Gras" align="right">&nbsp;&nbsp;&nbsp;Division:&nbsp;&nbsp;&nbsp;</td>
														<td><cc:combobox id="Division" runat="server" cssClass="Texte Combobox" width="250" AutoPostBack="true"></cc:combobox></td>
													</tr>
													<tr>
														<td align="right">City:</td>
														<td><asp:label class="Texte" id="CustomerCity" Runat="server"></asp:label></td>
														<td align="right">&nbsp;&nbsp;&nbsp;Customer #:&nbsp;&nbsp;&nbsp;</td>
														<td><asp:label class="Texte" id="CustomerNo" Runat="server" text=""></asp:label></td>
													</tr>
												</table>
											</td>
											<td>&nbsp;&nbsp;&nbsp;<asp:button class="BoutonGrand" id="Products" onclick="openProducts" Runat="server" Text="Add New Product Code"
													Enabled="False"></asp:button>
												<br>
												&nbsp;&nbsp;&nbsp;<asp:button class="BoutonGrand" id="Button2" onclick="SaveAndCloseForecasts" Runat="server"
													Text="Save And Close"></asp:button>
												<br>
												&nbsp;&nbsp;&nbsp;<asp:button class="BoutonGrand" id="Button1" onclick="OnlySaveForecasts" Runat="server" Text="Save"></asp:button>
												<br>
												&nbsp;&nbsp;&nbsp;<asp:button class="BoutonGrand" id="Button3" onclick="printInformation" Runat="server" Text="Print"></asp:button><br>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr vAlign="bottom">
								<td align="left">
									<asp:button id="SalesmanPerform" Runat="server" Text="Salesman Performance"></asp:button>
									<asp:button id="ExtraDetails" Runat="server" Text="Extra Details"></asp:button>
									<asp:button id="Initiatives" Runat="server" Text="Initiatives"></asp:button>
									<% If Session("ForecastEntry") = true then %>
									<asp:table id="DGForecastsHeader" Runat="server" CssClass="Gras Header BordureHeader" Width="100%"></asp:table>
									<div style="OVERFLOW: auto; HEIGHT: 330px"><asp:table class="texte BGDataGrid EspaceTableau" id="DGForecasts" Runat="server"></asp:table></div>
									<asp:Table ID="DGForecastsTotal" Runat="server" CssClass="Gras Header BordureHeader" Width="100%"></asp:Table>
									<% ElseIf Session("InitiativesEntry") = true then %>
									<asp:table id="DGInitiativesHeader" Runat="server" CssClass="Gras Header BordureHeader" Width="100%"></asp:table>
									<div style="OVERFLOW: auto; HEIGHT: 350px"><asp:table class="texte BGDataGrid EspaceTableau" id="DGInitiatives" Runat="server"></asp:table></div>
									<% ElseIf Session("ExtraDetails") = true then %>
									<asp:table id="DGDetailsHeader" Runat="server" CssClass="Gras Header BordureHeader" Width="100%"></asp:table>
									<div style="OVERFLOW: auto; HEIGHT: 320px"><asp:table class="texte BGDataGrid EspaceTableau" id="DGDetails" Runat="server"></asp:table></div>
									<asp:Table ID="DGDetailsTotal" Runat="server" CssClass="Gras Header BordureHeader" Width="100%"></asp:Table>
									<% End If %>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<!--  Contient la liste des lignes qui ont été modofiées --><input id="ForecastChanged" type="hidden" name="ForecastChanged" runat="server">
			<input id="InitiativeChanged" type="hidden" name="ForecastChanged" runat="server">
		</form>
	</body>
</HTML>
