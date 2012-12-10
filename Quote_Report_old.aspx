<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Quote_Report.aspx.vb" Inherits="APT.Quote_Report" %>
<%@ Register TagPrefix="CC" TagName="Combobox" Src="Combobox.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
		<title>Quote Report</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="styles.css" type="text/css" rel="stylesheet">
		<link rel="stylesheet" type="text/css" href="Calendar.css">
		<script language="JavaScript" src="CalendarPopup.js"></script>
  </head>
	<body bottomMargin="0" bgColor="whitesmoke" leftMargin="0" topMargin="0" rightMargin="0"
		marginwidth="0" marginheight="0">
		<div id="Calendrier" style="VISIBILITY:hidden;POSITION:absolute;"></div>
		<form id="QUOTE_REPORT" runat="server">
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
				<tr vAlign="top">
					<td background="images/body.gif" height="100%" border="0">
						<!-- Contenu -->
						<table align="center">
								<tr vAlign="middle" height="70">
									<td class="liens" align="center" colSpan="4"><U><FONT size="4">Quote Report Search</FONT></U></td>
								</tr>
								<TR>
									<TD style="HEIGHT: 19px" width="150" colSpan="4"><U><STRONG><FONT size="4">Quote / LCL Info</FONT></STRONG>
										</U>
									</TD>
								</TR>
								<tr>
									<td style="WIDTH: 159px; HEIGHT: 13px" width="159">Quote #</td>
					</td>
					<td style="HEIGHT: 13px" width="130">Start Date <FONT size="2">(mm/dd/yy)</FONT>
					</td>
					<td style="HEIGHT: 13px" width="130">End Date <FONT size="2">(mm/dd/yy)</FONT>
					</td>
					<td style="HEIGHT: 13px" width="50"></td>
				</tr>
				<tr height="30">
				</tr>
				<tr>
					<td style="WIDTH: 159px"><asp:textbox id="QuoteNo" CssClass="Combobox" Runat="server"></asp:textbox></td>
					<td>
						<asp:textbox id="StartDate" CssClass="Combobox" size="15" Runat="server"></asp:textbox>
						<a href="#" id="anchor1" onclick="cal1xx.select(document.forms[0].StartDate,'anchor1','MM/dd/yyyy'); return false;">
							<img src="images/calendar.jpg" border="0" width="21" height="17">
						</a>
					</td>
					<td>
						<asp:textbox id="EndDate" CssClass="Combobox" size="15" Runat="server"></asp:textbox>
						<a href="#" id="anchor2" onclick="cal1xx.select(document.forms[0].EndDate,'anchor2','MM/dd/yyyy'); return false;">
							<img src="images/calendar.jpg" border="0" width="21" height="17">
						</a>
					</td>
					<td></td>
				</tr>
				<tr height="30"></tr>
				<TR>
					<TD style="WIDTH: 159px">Quoted By</TD>
					<TD>Os #</TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 159px"><cc:combobox id="QuotedBy" runat="server" width="150" cssClass="Texte Combobox"></cc:combobox></TD>
					<TD><cc:combobox id="OSNO" runat="server" width="150" cssClass="Texte Combobox"></cc:combobox></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 159px; HEIGHT: 23px">&nbsp;</TD>
					<TD style="HEIGHT: 23px"></TD>
					<TD style="HEIGHT: 23px"></TD>
					<TD style="HEIGHT: 23px"></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 159px; HEIGHT: 23px" colSpan="4"><U><STRONG><FONT size="4">Follow-Up</FONT></STRONG></U></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 159px; HEIGHT: 23px">Who to Follow-Up</TD>
					<TD style="HEIGHT: 23px">F-U Start Date <FONT size="2">(mm/dd/yy)</FONT></TD>
					<TD style="HEIGHT: 23px">F-U End&nbsp;Date <FONT size="2">(mm/dd/yy)</FONT></TD>
					<TD style="HEIGHT: 23px"></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 159px"><cc:combobox id="WFollowUp" runat="server" width="150" cssClass="Texte Combobox"></cc:combobox></TD>
					<TD><asp:textbox id="FUStartDate" CssClass="Combobox" Runat="server"></asp:textbox></TD>
					<TD><asp:textbox id="FUEndDate" CssClass="Combobox" Runat="server"></asp:textbox></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD colSpan="4">&nbsp;
					</TD>
				</TR>
				<TR>
					<TD colSpan="4"><U><STRONG><FONT size="4">Customer / Contact</FONT></STRONG></U></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 159px">Customer #</TD>
					<TD>Customer&nbsp;Name</TD>
					<TD>City</TD>
					<TD>Contact's Last Name</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 159px"><asp:textbox id="CustNo" CssClass="Combobox" Runat="server"></asp:textbox></TD>
					<TD><asp:textbox id="BillName" CssClass="Combobox" Runat="server"></asp:textbox></TD>
					<TD><asp:textbox id="City" CssClass="Combobox" Runat="server"></asp:textbox></TD>
					<TD><asp:textbox id="LastName" CssClass="Combobox" Runat="server"></asp:textbox></TD>
				</TR>
				<tr height="30">
				</tr>
				<TR>
					<TD style="WIDTH: 150px; HEIGHT: 19px" colSpan="4">&nbsp;</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 150px; HEIGHT: 19px" colSpan="4"><U><STRONG><FONT size="4">Quoted Products</FONT></STRONG></U></TD>
				</TR>
				<tr>
					<td style="WIDTH: 159px; HEIGHT: 19px">Status</td>
					<td style="HEIGHT: 19px">Model Number</td>
					<td style="HEIGHT: 19px">Product Code</td>
					<td style="HEIGHT: 19px">Net Price</td>
				</tr>
				<tr>
					<td style="WIDTH: 159px; HEIGHT: 25px"><asp:textbox id="Status" CssClass="Combobox" Runat="server"></asp:textbox></td>
					<td style="HEIGHT: 25px"><asp:textbox id="ModelNumber" CssClass="Combobox" Runat="server"></asp:textbox></td>
					<td style="HEIGHT: 25px"><asp:textbox id="PC" CssClass="Combobox" Runat="server"></asp:textbox></td>
					<td style="HEIGHT: 25px"><asp:textbox id="NetPrice" CssClass="Combobox" Runat="server"></asp:textbox></td>
				</tr>
				<tr vAlign="middle" height="70">
					<td class="liens" align="center" colSpan="4"><asp:button id="Button1" onclick="SendSearch" Runat="server" Text="Search"></asp:button></td>
				</tr>
				
			
			</table>
		</form>
	</body>
</html>
