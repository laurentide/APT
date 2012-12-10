<%@ Register TagPrefix="CC" TagName="Combobox" Src="Combobox.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Product.aspx.vb" Inherits="APT.Product" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>New Product Code</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="styles.css" type="text/css" rel="stylesheet">
		<script>
			//places current window in the center
			var xpos = (screen.width - 400) / 2
			var ypos = (screen.height - 450) / 2
			moveTo(xpos, ypos);
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout" background="images/body.gif" leftMargin="0" bottommargin="0"
		topmargin="0" RightMargin="0" marginheight="0" marginwidth="0">
		<form id="Form1" method="post" runat="server" style="LEFT: 25px; POSITION: relative" class="Texte">
			<%If request.QueryString("Cus") <> Nothing then %>
			<br>
			<span class="Gras">Customer Information:</span>
			<table class="Bordure Texte">
				<tr>
					<td align="right" width="100">Customer #:</td>
					<td>
						<asp:Label Runat="server" id="CustomerNo" /><br>
					</td>
				</tr>
				<tr>
					<td align="right" width="100">Customer Name:</td>
					<td>
						<asp:Label Runat="server" id="CustomerName" /><br>
					</td>
				</tr>
				<tr>
					<td align="right" width="100">City:</td>
					<td>
						<asp:Label Runat="server" id="City" /><br>
					</td>
				</tr>
			</table>
			<%Else if request.QueryString("Reg") <> Nothing then %>
			<br>
			<span class="Gras">Region Information:</span>
			<table class="Bordure Texte">
				<tr>
					<td align="right" width="100">Region:</td>
					<td>
						<asp:Label Runat="server" id="RegionC" /><br>
						<asp:Label Runat="server" ID="RegionNo" Visible="False" />
					</td>
				</tr>
			</table>
			<% End If %>
			<br>
			<span class="Gras">Salesman Information:</span>
			<table class="Bordure Texte">
				<tr>
					<td align="right" width="100">Salesman #:</td>
					<td>
						<asp:Label Runat="server" id="OsNo" />
					</td>
				</tr>
				<tr>
					<td align="right" width="100">Salesman:</td>
					<td>
						<asp:Label Runat="server" id="OsName" /><br>
					</td>
				</tr>
			</table>
			<br>
			<span class="Gras">New Product Code:</span>
			<table class="Bordure Texte">
				<tr>
					<td>
						Division:
					</td>
					<td>
						<cc:combobox id="Division" runat="server" cssClass="Texte Combobox" width="250" AutoPostBack="true" />
					</td>
				</tr>
				<tr>
					<td>
						Product Code:
					</td>
					<td>
						<cc:combobox id="ProductCode" runat="server" cssClass="Texte Combobox" width="250" AutoPostBack="true" />
					</td>
				</tr>
				<tr>
					<td>
						Description:
					</td>
					<td>
						<asp:Label ID="Description" class="Combobox" Runat="server" />
					</td>
				</tr>
				<tr>
					<td>
						Forecast:
					</td>
					<td>
						<asp:TextBox class="Combobox" Text="0" size="10" Runat="server" ID="Forecast" />
					</td>
				</tr>
			</table>
			<br>
			<table width="95%">
				<tr>
					<td align="center">
						<asp:Button Text="Finish" Runat="server" OnClick="AddProduct" id="Button1" />
						<input type="button" value="Cancel" onclick="javascript:self.close();">
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
