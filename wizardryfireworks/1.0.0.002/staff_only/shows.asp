<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include virtual="/header.wiz"//-->
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") < 0 then
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
				<td colspan="3">
					<div style="{font-size:22px; color:white;}">
<%
	Dim Show_id, Showit, bolFound, strSQL

	Show_id = request.querystring("id")

	IF (Show_id) = "" THEN
		Showit = "all"	
	End if

	bolFound = false
	Response.Write("Welcome " & Session("staff_id") & "</div></td></tr>")

	strSQL = "SELECT * FROM Showdetails ORDER BY show_date DESC"
%>
	<!-- #include virtual="/DBconnect.wiz"//-->
<%

	If showit = "all" Then
%>
			<tr>
				<td style="{font-size:16px; color:white;}">Show date</td> 
				<td style="{font-size:16px; color:white;}">Client</td> 
				<td style="{font-size:16px; color:white;}">Click for More Details</td> 
			</tr>
<%	
		Do While Not (dbRst.EOF) 
%>
			<tr><td class="whitetext">
<%
			bolFound = False
			Response.Write dbRst("Show_date")
%>
				</td><td class="whitetext">
<%
			Response.Write dbRst("Show_client")
%>
				</td><td class="whitetext">
<%
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;Info : "
			Response.Write "<a href='/staff_only/shows.asp?id="& dbRst("id") &"'>Show details </a><br />"
%>
				</td></tr>
<%
			bolFound = True
		dbRst.MoveNext
		loop

	ELSE 
		Do until (dbRst.EOF or bolfound = "true")
			IF (strComp(dbRst("id"), Show_id, vbTextCompare) = 0) THEN

			BolFound = true
%>
			<tr>
				<td style="{font-size:16px; color:white;}">Show date</td> 
				<td style="{font-size:16px; color:white;}">Location</td> 
				<td style="{font-size:16px; color:white;}">Client</td> 
			</tr>
			<tr>
				<td class="whitetext">
<%
			Response.Write dbRst("show_date")
%>
				</td>
				<td class="whitetext">
<%
			Response.Write dbRst("show_location")
%>
				</td>
				<td class="whitetext">
<%
			Response.Write dbRst("show_client")
%>
				</td>
			</tr>
			<tr><td colspan="3">&nbsp;</td></tr>
			<tr>
				<td style="{font-size:16px; color:white;}" colspan="3">Show Type</td> 
			</tr>
			<tr>
				<td class="whitetext" colspan="3">
<%
			Response.Write dbRst("show_type")
%>
				</td>

			<tr><td colspan="3">&nbsp;</td></tr>
			<tr>
				<td style="{font-size:16px; color:white;}">Fire Time</td> 
				<td style="{font-size:16px; color:white;}">Work Start</td> 
				<td style="{font-size:16px; color:white;}">Work Finish</td> 
			</tr>
			<tr>
				<td class="whitetext">
<%
			Response.Write dbRst("fire_hour") & dbRst("fire_min") & dbRst("fire_am_pm")
%>
				</td>
				<td class="whitetext">
<%
			Response.Write dbRst("start_hour") & dbRst("start_min") & dbRst("start_am_pm")
%>
				</td>
				<td class="whitetext">
<%
			Response.Write dbRst("finish_hour") & dbRst("finish_min") & dbRst("finish_am_pm")
%>
				</td>

			<tr><td colspan="3">&nbsp;</td></tr>
			<tr>
				<td style="{font-size:16px; color:white;}" colspan="2">Show Price</td> 
				<td style="{font-size:16px; color:white;}">Staff needed</td> 
			</tr>
			<tr>
				<td class="whitetext" colspan="2">
<%
			Response.Write dbRst("show_price")
%>
				</td>
				<td class="whitetext">
<%
			Response.Write dbRst("staff_needed")
%>
				</td>

			<tr><td colspan="3">&nbsp;</td></tr>
			<tr>
				<td style="{font-size:16px; color:white;}" colspan="3">Description</td> 
			</tr>
			<tr>
				<td class="whitetext" colspan="3">
<%
			Response.Write dbRst("show_description")
%>
			<tr>
				<td class="whitetext" colspan="3"><a href="/staff_only/shows.asp">Show list</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="/staff_only/editshow.asp">Edit Show Information</a></td>
<%
			End if
		dbRst.MoveNext
		loop
	end if
	if bolFound = false then
%>
		<tr><td class="whitetext">Sorry, There is no show for this Client</td></tr>
<%
	end if
	dbRst.Close
	Set dbRst = Nothing
%>
				</td>
			</tr>
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