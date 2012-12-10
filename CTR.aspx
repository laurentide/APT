<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CTR.aspx.vb" Inherits="APT.CTR"%>
<%@ Register TagPrefix="CC" TagName="Combobox" Src="Combobox.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CTR</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body bgColor="whitesmoke" leftMargin="0" bottommargin="0" topmargin="0" RightMargin="0"
		marginheight="0" marginwidth="0">
		<form runat="server" ID="CTR">
			<table cellspacing="0" align="center" width="100%" height="100%">
				<tr>
					<td>
						<table width="100%" cellspacing="0">
							<tr>
								<td background="images/TitleFill.gif"></td>
								<td background="images/title.gif" width="680" height="95" border="0"></td>
								<td background="images/TitleFill.gif"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr valign="top">
					<td background="images/body.gif" border="0" height="100%">
						<!-- Contenu -->
						<table align="center">
							<TBODY>
								<tr height="70" valign="middle">
									<td colspan="4" align="center" class="liens" style="HEIGHT: 70px">CTR Search</td>
								</tr>
								<tr>
									<td width="1" style="WIDTH: 1px">Order #</td>
									<td width="130">Order Date (Start) &nbsp;&nbsp;&nbsp; <span class="smalltext">(Nomis: 
											1AAMMJJ)</span>
									</td>
									<td width="130">Order Date (End) &nbsp;&nbsp;&nbsp; <span class="smalltext">(Nomis: 
											1AAMMJJ)</span>
									</td>
									<td width="50">Model Number</td>
								</tr>
								<tr height="30">
								</tr>
								<tr>
									<td style="WIDTH: 1px"><asp:TextBox Runat="server" ID="OrderNo" CssClass="Combobox"></asp:TextBox></td>
									<td><asp:TextBox Runat="server" ID="StartDate" CssClass="Combobox"></asp:TextBox></td>
									<td><asp:TextBox Runat="server" ID="EndDate" CssClass="Combobox"></asp:TextBox></td>
									<td><asp:TextBox Runat="server" ID="Description" CssClass="Combobox"></asp:TextBox></td>
								</tr>
								<tr height="30">
								</tr>
								<tr>
									<td style="WIDTH: 1px">Os #</td>
									<td>Customer&nbsp;Name</td>
									<td>Unit Price</td>
									<td>Customer #</td>
								</tr>
								<tr>
									<td style="WIDTH: 1px"><cc:combobox id="OSNO" runat="server" cssClass="Texte Combobox" width="150"></cc:combobox></td>
									<td><asp:TextBox Runat="server" ID="BillName" CssClass="Combobox"></asp:TextBox></td>
									<td><asp:TextBox Runat="server" ID="List" CssClass="Combobox"></asp:TextBox></td>
									<td><asp:TextBox Runat="server" ID="CustNo" CssClass="Combobox"></asp:TextBox></td>
								</tr>
				</tr>
				<tr height="30">
				</tr>
				<tr>
					<td style="WIDTH: 1px">Shipping City</td>
					<td>Originator</td>
					<td>Customer PO #</td>
					<td>Product Code</td>
				</tr>
				<tr>
					<td style="WIDTH: 1px"><asp:TextBox Runat="server" ID="ShipCity" CssClass="Combobox"></asp:TextBox></td>
					<td><asp:TextBox Runat="server" ID="Orig" CssClass="Combobox"></asp:TextBox></td>
					<td><asp:TextBox Runat="server" ID="CustPONo" CssClass="Combobox"></asp:TextBox></td>
					<td><asp:TextBox Runat="server" ID="PC" CssClass="Combobox"></asp:TextBox></td>
				</tr>
				<tr>
					<td style="WIDTH: 1px">Quote #</td>
					<td>Customer Inventory Item</td>
					<td>Serial Number</td>
					<td>Discount Rate
					</td>
				</tr>
				<tr>
					<td style="WIDTH: 1px"><asp:TextBox Runat="server" ID="QuoteNo" CssClass="Combobox"></asp:TextBox></td>
					<td><asp:TextBox Runat="server" ID="CustInvItem" CssClass="Combobox"></asp:TextBox></td>
					<td><asp:TextBox Runat="server" ID="SerialNumber" CssClass="Combobox"></asp:TextBox></td>
					<td><asp:TextBox Runat="server" ID="DiscountRate" CssClass="Combobox"></asp:TextBox></td>
				</tr>
				<tr>
					<td style="WIDTH: 1px">Inventory Number</td>
					<td>Product Code Category</td>
					<td></td>
					<td></td>
				</tr>
				<tr>
					<td style="WIDTH: 1px"><asp:TextBox Runat="server" ID="InventoryNumber" CssClass="Combobox"></asp:TextBox></td>
					<td>
						<asp:TextBox id="ProductCodeCategory" CssClass="Combobox" Runat="server"></asp:TextBox></td>
					<td></td>
					<td></td>
				</tr>
				<tr height="70" valign="middle">
					<td colspan="4" align="center" class="liens">
						<asp:Button Runat="server" Text="Search" OnClick="SendSearch" id="Button1"></asp:Button>
					</td>
				</tr>
			</table>
			<!--Fin Contenu--> </TD></TR></TBODY></TABLE>
		</form>
	</body>
</HTML>
