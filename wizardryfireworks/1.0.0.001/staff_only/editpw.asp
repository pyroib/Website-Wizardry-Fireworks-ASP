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
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Edit your Password</div>
	</td>
</tr>
<tr>
	<td valign="top" height="485" align="center" style="padding: 10px; border: 1px ridge #FFFF66;">
		<table width="550" border="0" cellpadding="0" cellspacing="0">

<%
	If Request.Form("Submit") = "Change Password" Then
		Dim strSQL, staff_id, EventDate

		If Session("staff_id") = "" Then
			Response.redirect ("/staff_only/login.asp")
		Else
			staff_id = Session("staff_id")
			strSQL = "UPDATE logins SET staff_pw = '" & Request.Form("confirm_pw") & "' WHERE staff_id = '" & Session("staff_id") & "'"
%>
			<!-- #include virtual="/DBconnect.wiz"//-->
<%
			Response.redirect ("/staff_only")
		End If

	else
%>


			<form method="post" action="/staff_only/editpw.asp">
				<tr> 
					<td colspan="2" valign="top" height="50">&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2" valign="top" class="maintble"><img src="/images/logosmall.gif" alt="wizardry Logo" /><br />:: Change Staff Member Password ::<br /><br /></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">New Password : </td>
					<td width="250" valign="middle" height="50" align="left"><input type="password" name="confirm_pw" size="30" maxlength="10"></td>
				</tr>
				<tr> 
					<td colspan="2" valign="middle" height="50"><input type="submit" name="Submit" value="Change Password"></td>
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