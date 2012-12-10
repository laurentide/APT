<%@ Page Language="vb" Codebehind="APT.aspx.vb" Inherits="APT.APT" clientTarget="downlevel" %>
<%@ import Namespace="System.Data.OleDb" %>
<%@ import Namespace="System.Data" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Choose Customers</title>
		<script language="javascript" src="APT.js"></script>
		<script language="VB" runat="server">
			Sub Page_Load()
				Dim dbConnSqlServer as oleDbConnection = EtablitConnexionSQLServer()
				dbConnSqlServer.open
					
				if not page.ispostback then
					If session("OS") <> Nothing then
						Dim strReq as string = "Select OsName from NOMIS where OSNo='" & _
												Session("OS") & "'"
						Dim cmdTable as New OleDbDataAdapter(strReq, dbConnSqlServer)
						Dim dtTable as New DataTable
						cmdTable.fill(dtTable)
						if dtTable.rows.Count > 0 then
							lblSalesman.Text = dtTable.rows(0)(0)
						Else
							lblSalesman.Text = Session("OS")
						end if
					Else
						response.redirect("Denied.html")
					End If
				end if
				
				AfficheCustomers(dbConnSqlServer)
				CreateHeaderCustomers()
				dbConnSqlServer.close
			End Sub

		</script>
		<LINK href="styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body bgColor="whitesmoke" leftMargin="0" bottommargin="0" topmargin="0" RightMargin="0"
		marginheight="0" marginwidth="0">
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
						<table height="60%" width="99%" align="left">
							<tr>
								<td width="45"></td>
								<td class="Gras"><br>
									Salesman:
									<asp:label id="lblSalesman" runat="server"></asp:label></td>
							</tr>
							<tr>
								<td></td>
								<td><asp:table id="dgCustomersABHeader" Width="95%" Runat="server" CssClass="Gras Header BordureHeader"></asp:table>
									<div style="OVERFLOW: auto; WIDTH: 95%; HEIGHT: 390px">
										<asp:table class="texte BGDataGrid EspaceTableau" id="dgCustomersAB" Runat="server" EnableViewState="False" />
									</div>
								</td>
							</tr>
							<tr height="40">
								<td></td>
								<td align="right"><asp:button class="BoutonGrand" id="Button1" onclick="SaveCustomersAB" runat="server" text="Save And Close"></asp:button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
