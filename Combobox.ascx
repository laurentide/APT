<%@ Control Language="vb" AutoEventWireup="false" Codebehind="ComboBox.ascx.vb" Inherits="APT.ComboBox" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<script language="javascript" src="ComboBox.js"></script>
<table cellpadding="0" cellspacing="0">
	<tr valign="top">
		<td>
			<input type="hidden" id="TextisChanged" runat="server">
			<asp:textbox id="Texte" Runat="server" onKeyUp="javascript:TextChange();" />
		</td>
		<td>
			<img src="images\arrow.jpg" align="top" height="22" id="Arrow" onclick="ShowList();"
				runat="server">
		</td>
	</tr>
</table>
<DIV ID="DivList" runat="server" STYLE="VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white">
	<asp:ListBox Runat="server" id="listBox" onClick="javascript:FillText();" />
</DIV>
<script language="javascript">
	document.body.attachEvent("onclick", HideList);
</script>
