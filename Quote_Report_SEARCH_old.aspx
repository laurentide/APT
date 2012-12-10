<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Quote_Report_SEARCH.aspx.vb" Inherits="APT.Quote_Report_SEARCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Quote Report Results</title>
		<link href="styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<TABLE height="3523" cellSpacing="0" cellPadding="0" width="132" border="0" ms_2d_layout="TRUE">
			<TR vAlign="top">
				<TD width="132" height="3523">
					<form id="Form2" method="post" runat="server">
						<TABLE height="115" cellSpacing="0" cellPadding="0" width="3511" border="0" ms_2d_layout="TRUE">
							<TR vAlign="top">
								<TD width="10" height="15"></TD>
								<TD width="3501"></TD>
							</TR>
							<TR vAlign="top">
								<TD height="24"></TD>
								<TD>
									<asp:Button Runat="server" Text="Export to Excel" OnClick="ExportToExcel" ID="Button2" /></TD>
							</TR>
							<TR vAlign="top">
								<TD height="76"></TD>
								<TD>
									<table width="3500" id="TabReport" runat="server" height="75">
										<tr>
											<td>
												<span class="TitreRapport">Quote Report - Results</span>
												<br>
												<br>
												<asp:Table ID="Report" Runat="server" />
											</td>
										</tr>
									</table>
								</TD>
							</TR>
						</TABLE>
					</form>
				</TD>
			</TR>
		</TABLE>
	</body>
</HTML>
