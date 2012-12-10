<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.OleDb" %>
<%@ Page Language="vb" Codebehind="APT.aspx.vb" Inherits="APT.APT" clientTarget="downlevel" %>
<%@ Register TagPrefix="CC" TagName="Combobox" Src="Combobox.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Next Year Goals</title>
		<script language="javascript" src="APT.js"></script>
		<script language="VB" runat="server">

			Sub Page_Load()
				Dim dbConnSqlServer as oleDbConnection = EtablitConnexionSQLServer()
				dbConnSqlServer.open
				
				If User.IsInRole("LCLMTL\LCL_APT") = true Then
					CreateHeaderGoals()
					if not page.isPostBack then
						FillOsList(dbConnSqlServer)
						Session("OSGoals") = ""
					Else
						AddChangedGoals()
						if Session("OSGoals") <> CBOS.value then
							FillGoals(dbConnSqlServer)
							Session("OSGoals") = CBOS.value
						end if
					end if
					
					dbConnSqlServer.close
				Else
					response.redirect("Denied.html")
				End If
				
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
						<table height="60%" width="100%">
							<tr>
								<td>
									<table>
										<tr>
											<td align="left" class="Gras">Salesman:</td>
											<td><cc:combobox id="CBOS" runat="server" cssClass="Texte Combobox" width="150" AutoPostBack="true"></cc:combobox></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr vAlign="bottom">
								<td align="left">
									<asp:table id="DGGoalsHeader" Runat="server" CssClass="Gras Header BordureHeader" Width="100%"></asp:table>
									<div style="OVERFLOW: auto; HEIGHT: 400px">
										<asp:table class="texte BGDataGrid EspaceTableau" id="DGGoals" Runat="server"></asp:table>
									</div>
									<asp:Table ID="DGGoalsTotal" Runat="server" CssClass="Gras Header BordureHeader" Width="100%"></asp:Table>
								</td>
							</tr>
							<tr>
								<td align="right">
									<asp:Button Text="Save and Close" class="BoutonGrand" OnClick="SaveGoals" Runat="server" id="Button1" />
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<!--  Contient la liste des lignes qui ont été modofiées -->
			<input id="GoalsChanged" type="hidden" name="GoalsChanged" runat="server">
		</form>
	</body>
</HTML>
