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

	strSQL = "SELECT * FROM affiliates"
%>
	<!-- #include virtual="/DBconnect.wiz"//-->
<%
	If showit = "all" Then
%>
			<tr>
				<td class="whitetext">Name</td> 
				<td class="whitetext">Link</td> 
				<td class="whitetext">Current Image</td> 
				<td width="100" class="whitetext">Edit </td> 
				<td width="100" class="whitetext">Delete</td> 
			</tr>
<%
		a = 0
		Do While Not (dbRst.EOF OR a = 15)
			bolFound = False
%>
			<tr>
				<td class="whitetext"><% Response.Write dbRst("affiliatename") %></td>
				<td class="whitetext"><a href="<% Response.Write dbRst("affiliatelink") %>"><% Response.Write dbRst("affiliatelink") %></a></td>
				<td class="whitetext"><img src="<% Response.Write dbRst("affiliateimage") %>" width="150" alt="<% Response.Write dbRst("affiliatename") %>" /></td>
				<td class="whitetext"><a href="/staff_only/editaffiliate.asp?id=<% response.write dbRst("id")%>">Edit</a></td> 
				<td class="whitetext"><a href="/staff_only/deleteaffiliate.asp?id=<% response.write dbRst("id")%>">Delete</a></td> 
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
		<form method="post" action="/staff_only/editingaffiliate.asp?ID=<% response.write show_id %>">
<%
		Do until (dbRst.EOF or bolfound = "true")
			IF (strComp(dbRst("id"), Show_id, vbTextCompare) = 0) THEN
			BolFound = true
%>
			<tr>
				<td width="200" style="font-size:16px; color:white; text-align:right;">Name : </td> 
				<td class="whitetext" style="text-align:left;"><input type="text" name="affiliate_name" size="50" maxlength="50" value="<% Response.Write dbRst("affiliatename") %>"></td>

			</tr>
			<tr>
				<td style="font-size:16px; color:white; text-align:right;">Link : </td> 
				<td class="whitetext" style="text-align:left;"><input type="text" name="affiliate_link" size="50" maxlength="50" value="<% Response.Write dbRst("affiliatelink") %>"></td>
			</tr>
			<tr>
				<td style="font-size:16px; color:white; text-align:right;">Image URL : </td> 
				<td class="whitetext" style="text-align:left;"><input type="text" name="affiliate_image" size="50" maxlength="50" value="<% Response.Write dbRst("affiliateimage") %>"></td>
			</tr>
			<tr>
				<td style="font-size:16px; color:white; text-align:right;">Current Image : </td> 
				<td class="whitetext" style="text-align:left;"><img src="<% Response.Write dbRst("affiliateimage") %>" width="150" alt="<% Response.Write dbRst("affiliateimage") %>" /></td>
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
		<tr><td class="whitetext" colspan="5">Sorry, There is no Affiliate links at this time</td></tr>
<%
	end if
	dbRst.Close
	Set dbRst = Nothing
%>
			<tr>
				<td colspan="5" calss="whitetext"><a href="/staff_only"> Back to Staff Page </a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->