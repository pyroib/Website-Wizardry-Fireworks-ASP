<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include virtual="/header.wiz"//-->
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 2 then
		Response.redirect ("login.asp")
	end if
	End if
%>
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Future and Past Shows</div>
	</td>
</tr>
<tr>
	<td height="485" valign="top" align="center" style="{padding: 10px; border: 1px ridge #FFFF66;}">
		<table border ="1" width="550" cellpadding="0" cellspacing="0">
			<tr>
				<td colspan="5">
					<div style="{font-size:22px; color:white;}">
<%
	Dim Showit, bolFound, strSQL, a, counting, id, accesslevel

	id = request.querystring("id")

	IF id = "" THEN
		Showit = "all"	
	End if

	bolFound = false
	Response.Write("Welcome " & Session("staff_id") & "</div></td></tr>")

	strSQL = "SELECT * FROM logins"
%>
	<!-- #include virtual="/DBconnect.wiz"//-->
<%
	If showit = "all" Then
%>
			<tr>
				<td class="whitetext">Staff ID</td> 
				<td class="whitetext">Access Level</td> 
				<td class="whitetext">Edit</td> 
				<td class="whitetext">Delete</td> 
			</tr>
<%
		a = 0
		Do While Not (dbRst.EOF OR a = 15)
			bolFound = False

			if dbRst("permit") = 1 then
				accesslevel = "Admin / all access"
			else
				accesslevel = "Restricted"
			end if				
%>
			<tr>
				<td class="whitetext"><%Response.Write dbRst("staff_id") %></td>
				<td class="whitetext"><% response.write accesslevel %></td>
				<td class="whitetext"><a href="/staff_only/editlogin.asp?id=<% response.write dbRst("staff_id") %>">Edit</a></td> 
				<td class="whitetext"><a href="/staff_only/deletelogin.asp?id=<% response.write dbRst("staff_id") %>">Delete</a></td> 
			</tr>
<%
			a= a + 1
		dbRst.MoveNext
		if (dbRst.EOF) then
			bolfound = true
		end if
		loop


	ELSE
%>

		<form method="post" action="/staff_only/editinglogin.asp?ID=<% response.write id %>">
<%
		Do until (dbRst.EOF or bolfound = "true")
			IF (strComp(dbRst("staff_ID"), id, vbTextCompare) = 0) THEN
			BolFound = true
%>
			<tr>
				<td style="{font-size:16px; color:white;}">Staff Login ID</td> 
				<td width="125" style="{font-size:16px; color:white;}">Access level</td> 
				<td width="125" style="{font-size:16px; color:white;}">Password</td> 
			</tr>
			<tr>
				<td class="whitetext"><input type="text" name="staffid" size="20" maxlength="50" value="<% Response.Write dbRst("staff_ID") %>"></td>
				<td class="whitetext">
<% if dbRst("permit") = 1 then %>
						<select name="permit">
							<option value="1" selected >1 (all access)</option>
							<option value="2" >2 (restricted)</option>
						</select>
<% ELSE %>
						<select name="permit">
							<option value="1" >1 (all access)</option>
							<option value="2" selected >2 (restricted)</option>
						</select>
<% END IF %>
				</td>
				<td class="whitetext"><input type="password" name="staffpw" size="20" maxlength="50" value="<% Response.Write dbRst("staff_pw") %>"></td>
			</tr>
			<tr>
				<td colspan="3" class="whitetext"><input type="submit" name="Submit" value="submit"></td>
			</tr>
			<tr>
				<td colspan="3" class="whitetext"><a href="/staff_only/editlogin.asp">Back to Login List</a></td>
			</tr>
<%
		End if
		dbRst.MoveNext
		loop
%>
		</form>
<%
	end if
	if bolFound = false then
response.write show_id
response.write id
%>
		<tr><td class="whitetext" colspan="4">Sorry, There is no show for this Client</td></tr>
<%
	end if
	dbRst.Close
	Set dbRst = Nothing
%>
			<tr>
				<td colspan="4" class="whitetext"><a href="/staff_only/addlogin.asp"> Add a new staff member </a></td>
			</tr>
			<tr>
				<td colspan="4" class="whitetext"><a href="/staff_only"> Back to Staff Page </a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="{font-size:10px; color:white;}">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->