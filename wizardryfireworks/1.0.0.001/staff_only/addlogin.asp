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
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Add a New Staff Member</div>
	</td>
</tr>
<tr>
	<td valign="top" height="485" align="center" style="padding: 10px; border: 1px ridge #FFFF66;">
		<table width="550" border="0" cellpadding="0" cellspacing="0">

<%

	If Request.Form("Submit") = "submit" Then
		Dim strSQL, staff_id, EventDate, staff_pw, permit


		EventDate=formatdatetime(date,vbshortdate)


		If Session("staff_id") = "" Then
			Response.redirect ("login.asp")
		Else
			staff_id = Session("staff_id")

			staff_id = Request.form("staff_id")
			staff_id = Replace(staff_id,"'","")
			staff_pw = Request.form("staff_pw")
			staff_pw = Replace(staff_pw,"'","")
			permit = Request.form("permit")
			permit = Replace(permit,"'","")

			strSQL = "INSERT INTO logins (staff_id, staff_pw, permit) VALUES ('" & staff_id & "','" & staff_pw & "','" & permit & "')"
			%><!-- #include virtual="/DBconnect.wiz"//--><%
			Response.redirect ("/staff_only/login.asp")
		End If

	else
%>


			<form method="post" action="/staff_only/addlogin.asp">
				<tr> 
					<td colspan="2" valign="top" height="50">&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2" valign="top" class="maintble"><img src="/images/logosmall.gif" alt="wizardry Logo" /><br />:: Add Public News ::<br /><br /></td>
				</tr>
					<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Staff Username :</td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="staff_id" size="30" maxlength="25"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Password : </td>
					<td width="250" valign="middle" height="50" align="left"><input type="password" name="staff_pw" size="30" maxlength="25"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Access Level : </td>
					<td width="250" valign="middle" height="50" align="left">
						<select name="permit">
							<option value="1" >1 (Admin, All access)</option>
							<option value="2" selected >2 (Restricted)</option>
						</select>
					</td>
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