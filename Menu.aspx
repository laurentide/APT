<%@ Page Language="vb" Codebehind="APT.aspx.vb" Inherits="APT.APT" clientTarget="downlevel" %>
<%@ import Namespace="System.Data.OleDb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title></title>
		<script language="VB" runat="server">

	Sub Page_Load()
		
		If session("OS") = Nothing then
			Dim dbConnSQLSERVER as oleDbConnection = EtablitConnexionSQLServer()
			dbConnSQLSERVER.open
			GetSessionOs(dbConnSqlServer)
			dbConnSQLSERVER.close
		end if
		
		If session("OS") = Nothing and not User.IsInRole("LCLMTL\LCL_APT") then		
			response.redirect("Denied.html")
		end if
		
	End Sub

		</script>
		<link href="styles.css" type="text/css" rel="stylesheet">
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
				<tr valign="top">
					<td background="images/body.gif" border="0" height="100%">
						<!-- Contenu -->
						<table align="left">
							<tr height="40">
								<td></td>
							</tr>
							<tr>
								<td width="100"></td>
								<% if Session("OS") <> Nothing then %>
								<td>
									<a href="Customers.aspx"><img src="images/Board.gif" border="0" width="85" height="85"></a>
								</td>
								<td width="250"><a href="Customers.aspx" class="liens">Select A And B Customers</a></td>
								<% Else %>
								<td></td>
								<td width="250"></td>
								<% End If %>
								<td width="100"></td>
								<td align="right">
									<a href="Reports.aspx"><img src="images/Reports.gif" border="0" width="85" height="85"></a>
								</td>
								<td><a href="Reports.aspx" class="liens">Reports</a></td>
							</tr>
							<tr height="20">
								<td></td>
							</tr>
							<tr>
								<td width="100"></td>
								<% if Session("OS") <> Nothing then %>
								<td>
									<a href="Forecast.aspx"><img src="images/Forecasts.gif" width="80" border="0"></a>
								</td>
								<td><a href="Forecast.aspx" class="liens">Next Year Forecast</a></td>
								<% ELSE %>
								<td></td>
								<td></td>
								<% End If %>
								<td></td>
								<% If User.IsInRole("LCLMTL\LCL_APT") = true Then %>
								<td align="right">
									<a href="Admin.aspx"><img src="images/Admin.gif" border="0" width="85" height="85"></a>
								</td>
								<td><a href="Admin.aspx" class="liens">Admin</a></td>
								<% ELSE %>
								<td></td>
								<td></td>
								<% End If %>
							</tr>
							<tr height="20">
								<td></td>
							</tr>
							<tr>
								<td width="100"></td>
								<% If User.IsInRole("LCLMTL\LCL_APT") = true Then %>
								<td>
									<a href="Goals.aspx"><img src="images/Goals.gif" border="0" width="75" height="75"></a>
								</td>
								<td><a href="Goals.aspx" class="liens">Next Year Goals</a></td>
								<% ELSE %>
								<td></td>
								<td></td>
								<% End If %>
								<td></td>
								<td width="100" align="right">
									<a href="javascript:self.close();"><img src="images/Exit.gif" border="0" width="80" height="80"></a>
								</td>
								<td><a href="javascript:self.close();" class="liens">Exit</a></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
