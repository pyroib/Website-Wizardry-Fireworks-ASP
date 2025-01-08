<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Employment</div>
	</td>
</tr>
<tr>
	<td height="485" valign="top" align="center" style="padding: 10px; border: 1px ridge #FFFF66;">
		<table border="0" width="504" cellpadding="0" cellspacing="0">
					<%
						Dim strSQL
						strSQL = "SELECT * FROM employment"
					%>
						<!-- #include virtual="/DBconnect.wiz"//-->
					<%
						Do While Not (dbRst.EOF)
						IF NOT (dbRst.EOF) THEN
					%>
			<tr>
				<td colspan="9" height="20"></td>
			</tr>
			<tr>
				<td></td>
				<td class="employment" style="text-align:left;" height="16">&nbsp;&nbsp;&nbsp;&nbsp;Position Vacant</td>
				<td></td>
				<td class="employment" style="text-align:left;" height="16">&nbsp;&nbsp;&nbsp;&nbsp;Position Vacant</td>
				<td></td>
			</tr>
			<tr>
				<td width="20"></td>
				<td width="186" class="whitetext" style="text-align:left; padding: 10px; border: 1px ridge #FFFF66;">
					<% 
							Response.Write "<div class='text'>Job Description :</div>"
							Response.Write dbRst("JobDescription")
							Response.Write "<br /><br /><div class='text'>Skills and Requirements :</div>"
							Response.Write dbRst("Requirements")
							Response.Write "<br /><br /><div class='text'>Sallery :</div>"
							Response.Write dbRst("Wage")
							Response.Write "<br /><br /><div class='text'>Date needed to start work :</div>"
							Response.Write dbRst("Needed")
							Response.Write "<br />"

							dbRst.MoveNext
						END IF
					%>
				</td>
				<td width="50"></td>
					<%
						IF NOT (dbRst.EOF) THEN
					%>
				<td width="186" class="whitetext" style="text-align:left; padding: 10px; border: 1px ridge #FFFF66;">
					<%
							Response.Write "<div class='text'>Job Description :</div>"
							Response.Write dbRst("JobDescription")
							Response.Write "<br /><br /><div class='text'>Skills and Requirements :</div>"
							Response.Write dbRst("Requirements")
							Response.Write "<br /><br /><div class='text'>Sallery :</div>"
							Response.Write dbRst("Wage")
							Response.Write "<br /><br /><div class='text'>Date needed to start work :</div>"
							Response.Write dbRst("Needed")
							Response.Write "<br />"

							dbRst.MoveNext
					%>

				</td>
				<td width="20"></td>
			</tr>
					<% ELSE %>
				<td width="207" style="padding: 0px; border: 1px ridge #FFFF66;"> &nbsp;</td>
				<td width="20"></td>
			</tr>
					<%
						END IF
						loop
					%>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->
