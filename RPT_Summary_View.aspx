<%@ Page Language="vb" AutoEventWireup="false" Codebehind="RPT_Summary_View.aspx.vb" Inherits="APT.RPT_Summary_View"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Summary View</title>
		<link href="styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<TABLE height="100021" cellSpacing="0" cellPadding="0" width="181" border="0" ms_2d_layout="TRUE">
			<TR vAlign="top">
				<TD width="181" height="100021">
					<form id="Form1" method="post" runat="server">
						<TABLE height="164" cellSpacing="0" cellPadding="0" width="100009" border="0" ms_2d_layout="TRUE">
							<TR vAlign="top">
								<TD width="8" height="8"></TD>
								<TD width="100001"></TD>
							</TR>
							<TR vAlign="top">
								<TD height="40"></TD>
								<TD>
									<asp:TextBox id="TextBox1" runat="server" BorderStyle="None" BorderColor="Transparent" ForeColor="#0000C0"></asp:TextBox></TD>
							</TR>
							<TR vAlign="top">
								<TD height="40"></TD>
								<TD>
									<asp:Button Runat="server" Text="Export to Excel" OnClick="ExportToExcel" ID="Button1" NAME="Button1" /></TD>
							</TR>
							<TR vAlign="top">
								<TD height="76"></TD>
								<TD>
									<table width="100000" id="TabReport" runat="server" height="75">
										<tr>
											<td>
												<span class="TitreRapport">Summary View</span>
												<br>
												<br>
												<asp:PlaceHolder id="Report" Runat="server"></asp:PlaceHolder>
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
