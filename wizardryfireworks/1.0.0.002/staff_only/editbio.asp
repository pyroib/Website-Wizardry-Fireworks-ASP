<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") < 0 then
		Response.redirect ("login.asp")
	end if
	End if
%>
<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Edit your Biography</div>
	</td>
</tr>
<tr>
	<td valign="top" height="485" align="center" style="{padding: 10px; border: 1px ridge #FFFF66;}">
		<table width="550" border="0" cellpadding="0" cellspacing="0">

<%
	If Request.Form("Submit") = "Change Biography" Then
		Dim strSQL, staff_id, EventDate, new_name, new_age, new_bio

		If Session("staff_id") = "" Then
			Response.redirect ("/staff_only/login.asp")
		Else
			staff_id = Session("staff_id")

			new_name = Request.form("new_name")
			new_name = Replace(new_name,"'","")
			new_age = Request.form("new_age")
			new_age = Replace(new_age,"'","")
			new_bio = Request.form("new_bio")
			new_bio = Replace(new_bio,"'","")

			strSQL = "UPDATE staffProfiles SET Name = '" & new_name & "', age = '" & new_age & "', experience = '" & new_bio & "' WHERE staff_id = '" & staff_id & "'"
%>
			<!-- #include virtual="/DBconnect.wiz"//-->
<%
			Response.redirect "/profiles"
		End If

	else
%>
			<form method="post" action="/staff_only/editbio.asp">
				<tr> 
					<td valign="top" colspan="2" height="50">&nbsp;</td>
				</tr>
				<tr> 
					<td valign="top" colspan="2" class="maintble"><img src="/images/logosmall.gif" alt="wizardry Logo" /></td>
				</tr>
				<tr>
					<td colspan="2">:: Change Staff Biography ::</td>
				</tr>
				<tr>
					<td width="275" class="text" style="{text-align:right;}">Full Name : </td>
					<td style="{text-align:left;}"><input type="text" name="new_name" size="30" maxlength="20"></td>
				</tr>
				<tr>
					<td width="275" class="text" style="{text-align:right;}">Age : </td>
					<td style="{text-align:left;}"><input type="text" name="new_age" size="30" maxlength="2"></td>
				</tr>
				<tr>
					<td colspan="2" height="30"></td>
				</tr>
				<tr>
					<td colspan="2" class="text">:: Edit This to Update your Biography ::</td>
				</tr>
				<tr>
					<td colspan="2"><input type="text" name="new_bio" maxlength="450" size="85"></td>
				</tr>
				<tr> 
					<td colspan="2" valign="middle" height="50"><input type="submit" name="Submit" value="Change Biography"></td>
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
		<div style="{font-size:10px; color:white;}">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->








<%
	end if
%>