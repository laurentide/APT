<%@ Page Language="vb" AutoEventWireup="false" Codebehind="RPT_Multi_Year_Bookings_Results.aspx.vb" Inherits="APT.RPT_Multi_Year_Bookings_Results" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
		<title>Multi-Year Bookings Results</title>
		<link href="styles.css" type="text/css" rel="stylesheet">
  </HEAD>
	<body MS_POSITIONING="GridLayout">
<TABLE height=100023 cellSpacing=0 cellPadding=0 width=132 border=0 
ms_2d_layout="TRUE">
  <TR vAlign=top>
    <TD width=132 height=100023>
		<form id="Form1" method="post" runat="server">
      <TABLE height=115 cellSpacing=0 cellPadding=0 width=100011 border=0 
      ms_2d_layout="TRUE">
        <TR vAlign=top>
          <TD width=10 height=15></TD>
          <TD width=100001></TD></TR>
        <TR vAlign=top>
          <TD height=24></TD>
          <TD>
			<asp:Button Runat="server" Text="Export to Excel" OnClick="ExportToExcel" ID="Button1"/></TD></TR>
        <TR vAlign=top>
          <TD height=76></TD>
          <TD>
			<table width="100000" id="TabReport" runat="server" height=75>
				<tr>
					<td>
						<span class="TitreRapport">Multi-Year Bookings Results</span>
						<br ><br >
						<asp:PlaceHolder ID="Report" Runat="server" />		
					</td>
				</tr>
			</table></TD></TR></TABLE>	
		</form></TD></TR></TABLE>
	</body>
</HTML>
