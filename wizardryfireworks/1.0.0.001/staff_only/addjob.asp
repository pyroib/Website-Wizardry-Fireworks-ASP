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
		Dim strSQL, staff_id, EventDate, JobDescription, Requirements, Wage, Needed

		JobDescription = Request.form("JobDescription")
		JobDescription = Replace(JobDescription,"'","")
		Requirements = Request.form("Requirements")
		Requirements = Replace(Requirements,"'","")
		Wage = Request.form("Wage")
		Wage = Replace(Wage,"'","")
		Needed = Request.form("Needed")
		Needed = Replace(Needed,"'","")

		staff_id = Session("staff_id")
		strSQL = "INSERT INTO employment (JobDescription, Requirements, Wage, Needed) VALUES ('"& JobDescription & "','"& Requirements & "','" & Wage &"','" & Needed & "')"
		%><!-- #include virtual="/DBconnect.wiz"//--><%
		Response.redirect ("/staff_only/login.asp")

	else
%>

			<form method="post" action="/staff_only/addjob.asp">
				<tr> 
					<td colspan="2" valign="top" height="50">&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2" valign="top" class="maintble"><img src="/images/logosmall.gif" alt="wizardry Logo" /><br />:: Add Position Vacant ::<br /><br /></td>
				</tr>
					<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Job Description :</td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="JobDescription" size="30" maxlength="50"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Requirements : </td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="Requirements" size="30" maxlength="255"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Wage : </td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="Wage" size="30" maxlength="50"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="text" style="text-align:right;">Needed : </td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="Needed" size="30" maxlength="50"></td>
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