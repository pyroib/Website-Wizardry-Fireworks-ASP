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
	<td height="485" valign="top" align="center" style="padding: 10px; border: 1px ridge #FFFF66;">
		<table border ="1" width="550" cellpadding="0" cellspacing="0">
			<tr>
				<td colspan="5">
					<div style="font-size:22px; color:white;">
<%
	Dim Show_id, Showit, bolFound, strSQL, a, counting, id

	Show_id = request.querystring("id")

	IF (Show_id) = "" THEN
		Showit = "all"	
	End if

	bolFound = false
	Response.Write("Welcome " & Session("staff_id") & "</div></td></tr>")

	strSQL = "SELECT * FROM publicnews ORDER BY date DESC"
%>
	<!-- #include virtual="/DBconnect.wiz"//-->
<%
	If showit = "all" Then
%>
			<tr>
				<td class="whitetext">News Topic</td> 
				<td class="whitetext">Posted by</td> 
				<td width="100" class="whitetext">Edit News information</td> 
				<td width="100" class="whitetext">Delete News information</td> 
			</tr>
<%
		a = 0
		Do While Not (dbRst.EOF OR a = 15)
			bolFound = False
%>
			<tr>
				<td class="whitetext"><a href="/news/view.asp?id=<%Response.Write dbRst("id")%>" target="_blank"><%Response.Write dbRst("topic")%></a></td>
				<td class="whitetext"><% Response.Write dbRst("postedby") %></td>
				<td class="whitetext"><a href="/staff_only/editnews.asp?id=<% response.write dbRst("id")%>">Edit</a></td> 
				<td class="whitetext"><a href="/staff_only/deletenews.asp?id=<% response.write dbRst("id")%>">Delete</a></td> 
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
		<form method="post" action="/staff_only/editingnews.asp?ID=<% response.write show_id %>">
<%
		Do until (dbRst.EOF or bolfound = "true")
			IF (strComp(dbRst("id"), Show_id, vbTextCompare) = 0) THEN
			BolFound = true
%>
			<tr>
				<td style="font-size:16px; color:white;">Topic</td> 
				<td width="125" style="font-size:16px; color:white;">Posted By</td> 
				<td width="125" style="font-size:16px; color:white;">Posted On</td> 
			</tr>
			<tr>
				<td class="whitetext"><input type="text" name="topic" size="40" maxlength="50" value="<% Response.Write dbRst("topic") %>"></td>
				<td class="whitetext"><% Response.Write dbRst("postedby") %></td>
				<td class="whitetext"><% Response.Write dbRst("date") %></td>
			</tr>
			<tr>
				<td colspan="3" style="font-size:16px; color:white; text-align:left;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;News</td>
			</tr>
			<tr>
				<td colspan="3" class="whitetext"><input type="text" name="news" size="81" maxlength="100" value="<% Response.Write dbRst("news") %>"></td>
			</tr>
			<tr>
				<td colspan="3" class="whitetext"><input type="submit" name="Submit" value="submit"></td>
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
		<tr><td class="whitetext" colspan="4">Sorry, There is no news at this time</td></tr>
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
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->