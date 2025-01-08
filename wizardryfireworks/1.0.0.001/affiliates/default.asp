<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Wizardry Fireworks Partners and Affilitates</div>
	</td>
</tr>
<tr>
	<td height="485" valign="top" align="left" style="padding: 10px; border: 1px ridge #FFFF66;">
		<div style="font-size:12px; color:white; text-align:center;">

<% 
Dim strSQL
strSQL = "SELECT * FROM affiliates"
%>
<!-- #include virtual="/DBconnect.wiz"//-->
<%

Do While Not (dbRst.EOF)
	Response.Write "<a href='"& dbRst("affiliatelink") &"' target='_blank'><img border='0' src='"& dbRst("affiliateimage") &"' alt='Link to "& dbRst("affiliatename") &" '></a><br />"
	Response.Write dbRst("affiliatename")
	Response.Write " <hr /> "
	dbRst.MoveNext
loop
%>
		</div>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->
