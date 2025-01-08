<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if
%>
<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Add a position vacant information</div>
	</td>
</tr>
<tr>
	<td valign="top" height="485" align="center" style="padding: 10px; border: 1px ridge #FFFF66;">
		<table width="550" border="0" cellpadding="0" cellspacing="0">

<%

	If Request.Form("Submit") = "submit" Then
		Dim strSQL, staff_id, EventDate, affiliatename, affiliatelink, affiliateimage

		staff_id = Session("staff_id")

		affiliatename = Request.form("name")
		affiliatename = Replace(affiliatename,"'","")
		affiliatelink = Request.form("link")
		affiliatelink = Replace(affiliatelink,"'","")
		affiliateimage = Request.form("image")
		affiliateimage = Replace(affiliateimage,"'","")

		strSQL = "INSERT INTO affiliates (affiliatename, affiliatelink, affiliateimage) VALUES ('" & affiliatename & "','" & affiliatelink & "','" & affiliateimage & "')"
		%><!-- #include virtual="/DBconnect.wiz"//--><%
		Response.redirect ("/staff_only/login.asp")

	else
%>

			<form method="post" action="/staff_only/addaffiliate.asp">
				<tr> 
					<td colspan="2" valign="top" height="50">&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2" valign="top" class="maintble"><img src="/images/logosmall.gif" alt="wizardry Logo" /><br />:: Add Position Vacant ::<br /><br /></td>
				</tr>
					<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Name :</td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="name" size="30" maxlength="100"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Link : </td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="link" size="30" maxlength="255"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Image URL : </td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="image" size="30" maxlength="255"></td>
				</tr>
				<tr> 
					<td colspan="2" valign="middle" align="center" height="50"><input type="submit" name="Submit" value="submit"></td>
				</tr>
				<tr>
					<td colspan="2"><br /><br /></td>
				</tr>
			</form>
			<tr>
				<td colspan="4"><a href="/staff_only"> Back to Staff Page </a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->








<%
	end if
%>