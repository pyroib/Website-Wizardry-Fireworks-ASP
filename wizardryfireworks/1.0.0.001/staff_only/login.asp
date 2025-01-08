<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Staff Only Section</div>
	</td>
</tr>
<tr>
	<td height="485" valign="top" align="center" class="TextNoBG" style="padding: 10px; border: 1px ridge #FFFF66;">
<%
if Session("staff_id") = "" Then

	If Request.Form("Submit") = "Login" Then
		Dim staff_id, staff_pw, dbConn, strSQL, dbRst

		staff_id = Request.Form("txtStaff_id")
		staff_pw = Request.Form("txtStaff_pw")

		staff_id = Replace(staff_id,"'","")
		staff_pw = Replace(staff_pw,"'","")


		Set dbConn = Server.CreateObject("ADODB.Connection")

		dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\Dbase\wizDB.mdb") & ";"
		dbConn.Open
		'check user name and password
		strSQL = "SELECT * FROM logins WHERE staff_id ='" & staff_id & "' AND staff_pw='" & staff_pw & "'"
		
		Set dbRst = dbConn.Execute(strSQL)
		
		If dbRst.EOF And dbRst.BOF Then
			'details incorrect
			dbConn.Close
			Set dbConn = Nothing
			Response.redirect ("login.asp")
		Else

			Session("staff_id") = staff_id
			Session("permit") = dbRst("permit")
			dbConn.Close
			Set dbConn = Nothing
			Response.redirect ("login.asp")
		End if
	else
%>
		<table width="550" border="0" cellpadding="0" cellspacing="0">
			<form method="post" action="login.asp">
				<tr> 
					<td colspan="2" valign="top" height="50">&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2" valign="top" class="maintble"><img src="/images/logosmall.gif" alt="wizardry Logo" /><br />Login here: <br /><br /></td>
				</tr>
					<tr> 
					<td width="250" valign="middle" height="50" align="right" class="login">Staff ID:</td>
					<td width="250" valign="middle" height="50" align="left"><input type="text" name="txtstaff_Id" size="30" maxlength="25"></td>
				</tr>
				<tr> 
					<td width="250" valign="middle" height="50" align="right" class="login"><b>Password:</b></td>
					<td width="250" valign="middle" height="50" align="left"><input type="password" name="txtstaff_pw" size="30" maxlength="25"></td>
				</tr>
				<tr> 
					<td colspan="2" valign="middle" height="50"><input type="submit" name="Submit" value="Login"></td>
				</tr>
				<tr>
					<td colspan="2"><br /><br /></td>
				</tr>
			</form>
		</table>
<%
	end if
else
%>
		<table width="550" border="0" cellpadding="0" cellspacing="0">
			<tr><td valign="top" height="20"></td></tr>
			<tr> 
				<td valign="top" width="150" style="padding: 10px; border: 1px ridge #FFFF66;">
					<div class="text">Edit Your staff Details<br /><br /><br /></div>
					Information Availiable<br /> to all staff<br />
					<a href="shows.asp">View all Shows</a><br />
					<a href="news.asp">View STAFF ONLY News</a><br /><br />
					Change your details<br />
					<a href="editpw.asp">Change my Password</a><br />
					<a href="editbio.asp">Change my Biography</a><br /><br />
				</td>
				<td valign="top" width="250" style="padding: 10px; border: 1px ridge #FFFF66;">
					<div class="text">Edit Staff Only sections<br /><br /><br /></div>
		<% if Session("permit") < 2 then %>
					Edit PUBLIC news pages<br />
					<a href="addnews.asp">Add News</a><br />
					<a href="editnews.asp">Edit News</a><br />
					<a href="editnews.asp">Delete News</a><br /><br />
					Information on future<br /> and past Pyrotehcnical Displays<br />
					<a href="addshow.asp">Add show Information</a><br />
					<a href="editshow.asp">Edit show Information</a><br />
					<a href="editshow.asp">Delete show Information</a><br /><br />
					Add, Edit and Delete staff<br /> login Priveledges<br />
					<a href="addlogin.asp">Add Login User</a><br />
					<a href="editlogin.asp">Edit Login User</a><br />
					<a href="editlogin.asp">Delete Login User</a><br /><br />
		<% END IF %>
				</td>
				<td valign="top" width="150" style="padding: 10px; border: 1px ridge #FFFF66;">
					<div class="text">Edit Public sections<br /><br /><br /></div>
		<% if Session("permit") < 2 then %>
					Add, Edit and Delete <br /> Employment Details<br />
					<a href="addjob.asp">Add Employment</a><br />
					<a href="editjob.asp">Edit Employment</a><br />
					<a href="editjob.asp">Delete Employment</a><br /><br />
					Add, Edit and Delete<br /> Affiliate details<br />
					<a href="addaffiliate.asp">Add Affiliate</a><br />
					<a href="editaffiliate.asp">Edit Affiliate</a><br />
					<a href="editaffiliate.asp">Delete Affiliate</a><br /><br />
		<% END IF %>
				</td>
			</tr>
		</table>
		<% END IF %>
	</td>
<!-- #include virtual="/footer.wiz"//-->

