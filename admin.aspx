<%@ Page Language="vb" Codebehind="APT.aspx.vb" Inherits="APT.APT" clientTarget="downlevel" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Administrator Page</title>
		<link href="styles.css" type="text/css" rel="stylesheet">
			<script language="VB" runat="server">

			Sub Page_Load()
				session("OS") = Nothing 
				
				If Not User.IsInRole("LCLMTL\LCL_APT") Then
					response.redirect("Denied.html")
				End if	
			End Sub

			</script>
	</HEAD>
	<body leftMargin="0" bottommargin="0" topmargin="0" RightMargin="0" marginheight="0" marginwidth="0">
		<form runat="server">
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
				<tr>
					<td background="images/body.gif" border="0" height="100%">
						<!-- Contenu -->
						<table align="center">
							<tr height="50">
								<td>
								<td></td>
							<tr> 
								<td width="30"></td>
								<td class="Gras">
									Salesman's Windows Name:
								</td>
								<td>
									<asp:TextBox id="WINNAME" runat="server" />
								</td>
							<tr height="50">
								<td></td>
								<td colspan="2" align="center" class="Gras">Or</td>
							<tr>
							<tr>
								<td></td>
								<td class="Gras" align="right">
									Salesman's No:
								</td>
								<td>
									<asp:TextBox id="OSNo" runat="server" />
								</td>
							</tr>
							<tr height="90">
								<td></td>
								<td colspan="2" align="center">
									<asp:Button Class="Bouton" text="Ok" onClick="RedirectMenu" runat="server" id="Button1" Width="60" />
									<asp:Button class="BoutonGrand" Runat="server" Text="Back to Main Menu" OnClick="RedirectMenu" ID="Button2"/>
								</td>
							<tr>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
