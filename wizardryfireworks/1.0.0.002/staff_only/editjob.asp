<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include virtual="/header.wiz"//-->
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
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
	Dim Show_id, Showit, bolFound, strSQL, a, counting, id

	Show_id = request.querystring("id")

	IF (Show_id) = "" THEN
		Showit = "all"	
	End if

	bolFound = false
	Response.Write("Welcome " & Session("staff_id") & "</div></td></tr>")

	strSQL = "SELECT * FROM employment"
%>
	<!-- #include virtual="/DBconnect.wiz"//-->
<%
	If showit = "all" Then
%>
			<tr>
				<td class="whitetext">Job Description</td> 
				<td class="whitetext">Requirements</td> 
				<td width="100" class="whitetext">Edit Job Details</td> 
				<td width="100" class="whitetext">Delete position vacant</td> 
			</tr>
<%
		a = 0
		Do While Not (dbRst.EOF OR a = 15)
			bolFound = False
%>
			<tr>
				<td class="whitetext"><a href="/employment/default.asp" target="_blank"><%Response.Write dbRst("jobDescription")%></a></td>
				<td class="whitetext"><% Response.Write dbRst("Requirements") %></td>
				<td class="whitetext"><a href="/staff_only/editjob.asp?id=<% response.write dbRst("id")%>">Edit</a></td> 
				<td class="whitetext"><a href="/staff_only/deletejob.asp?id=<% response.write dbRst("id")%>">Delete</a></td> 
			</tr>
		</td>
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
		<form method="post" action="/staff_only/editingjob.asp?ID=<% response.write show_id %>">
<%
		Do until (dbRst.EOF or bolfound = "true")
			IF (strComp(dbRst("id"), Show_id, vbTextCompare) = 0) THEN
			BolFound = true
%>
			<tr>
				<td width="200" style="{font-size:16px; color:white; text-align:right;}">Job Description : </td> 
				<td class="whitetext" style="{text-align:left;}"><input type="text" name="jobdescription" size="50" maxlength="50" value="<% Response.Write dbRst("JobDescription") %>"></td>

			</tr>
			<tr>
				<td style="{font-size:16px; color:white; text-align:right;}">Requirements : </td> 
				<td class="whitetext" style="{text-align:left;}"><input type="text" name="requirements" size="50" maxlength="50" value="<% Response.Write dbRst("requirements") %>"></td>
			</tr>
			<tr>
				<td style="{font-size:16px; color:white; text-align:right;}">Wage : </td> 
				<td class="whitetext" style="{text-align:left;}"><input type="text" name="wage" size="50" maxlength="50" value="<% Response.Write dbRst("Wage") %>"></td>
			</tr>
			<tr>
				<td style="{font-size:16px; color:white; text-align:right;}">Start Date : </td> 
				<td class="whitetext" style="{text-align:left;}"><input type="text" name="needed" size="50" maxlength="50" value="<% Response.Write dbRst("needed") %>"></td>
			</tr>
			<tr>
				<td colspan="2" class="whitetext"><input type="submit" name="Submit" value="submit"></td>
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
		<tr><td class="whitetext" colspan="4">Sorry, There is no Job vacancies at this time</td></tr>
<%
	end if
	dbRst.Close
	Set dbRst = Nothing
%>
			<tr>
				<td colspan="4" calss="whitetext"><a href="/staff_only"> Back to Staff Page </a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="{font-size:10px; color:white;}">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->