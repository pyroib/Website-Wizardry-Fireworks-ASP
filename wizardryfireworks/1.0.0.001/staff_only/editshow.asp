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
	Dim Show_id, Showit, bolFound, strSQL, a, counting

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
				<td style="font-size:16px; color:white;">Show date</td> 
				<td width="175" style="font-size:16px; color:white;">Client</td> 
				<td style="font-size:16px; color:white;">Click for More Details</td>
				<td style="font-size:16px; color:white;">Edit Show information</td> 
				<td style="font-size:16px; color:white;">Delete Show information</td> 
			</tr>
<%
		a = 0
		Do While Not (dbRst.EOF OR a = 15)
			bolFound = False
%>
			<tr>
				<td class="whitetext"><%Response.Write dbRst("Show_date")%></td>
				<td class="whitetext"><% Response.Write dbRst("Show_client") %></td>
				<td class="whitetext"><a href="/staff_only/shows.asp?id=<% response.write dbRst("id") %>">Show details</a></td>
				<td class="whitetext"><a href="/staff_only/editshow.asp?id=<% response.write dbRst("id")%>">Edit</a></td> 
				<td class="whitetext"><a href="/staff_only/deleteshow.asp?id=<% response.write dbRst("id")%>">Delete</a></td> 
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
		<form method="post" action="/staff_only/editingshow.asp?ID=<% response.write show_id %>">
<%
		Do until (dbRst.EOF or bolfound = "true")
			IF (strComp(dbRst("id"), Show_id, vbTextCompare) = 0) THEN

			BolFound = true
%>
			<tr>
				<td style="font-size:16px; color:white;">Show date</td> 
				<td style="font-size:16px; color:white;">Location</td> 
				<td style="font-size:16px; color:white;">Client</td> 
			</tr>
			<tr>
				<td class="whitetext"><input type="text" name="show_date" size="20" maxlength="10" value="<% Response.Write dbRst("show_date") %>"></td>
				<td class="whitetext"><input type="text" name="show_location" size="20" maxlength="50" value="<% Response.Write dbRst("show_location") %>"></td>
				<td class="whitetext"><input type="text" name="show_client" size="20" maxlength="50" value="<% Response.Write dbRst("show_client") %>"></td>
			</tr>
			<tr>
				<td colspan="3">&nbsp;</td>
			</tr>
			<tr>
				<td style="font-size:16px; color:white; text-align:left;" colspan="3">Show Type</td> 
			</tr>
			<tr>
				<td class="whitetext" style="font-size:16px; color:white; text-align:left;" colspan="3"><input type="text" name="show_type" size="85" maxlength="10" value="<% Response.Write dbRst("show_type") %>"></td>
			<tr>
				<td colspan="3">&nbsp;</td></tr>
			<tr>
				<td style="font-size:16px; color:white;">Fire Time</td> 
				<td style="font-size:16px; color:white;">Work Start</td> 
				<td style="font-size:16px; color:white;">Work Finish</td> 
			</tr>
			<tr>
				<td class="whitetext">
					<select name="fire_hour">
<%
	counting=1

	do until counting = 13

 		response.write "<option value='"& counting &"'"

		if counting = cint(dbRst("fire_hour")) then
 			response.write (" selected ")
		end if

 		response.write ">"& counting &"</option>"

		counting= counting + 1
	loop
%>
					</select>
						<select name="fire_min">
<%
	counting=00

	do until counting = 60

 		response.write "<option value='"& counting &"'"

		if counting = dbRst("fire_min") then
 			response.write (" selected ")
		end if

 		response.write ">"& counting &"</option>"

		counting= counting + 5
	loop
%>
						</select>
						<select name="fire_am_pm">
<% if dbRst("fire_am_pm") = "am" then %>
						<option value="am" selected >AM</option>
						<option value="pm">PM</option>
<% else %>
						<option value="am">AM</option>
						<option value="pm" selected >PM</option>
<% end if %>
						</select>
				</td>
				<td class="whitetext">
					<select name="start_hour">
<%
	counting=1

	do until counting = 13

 		response.write "<option value='"& counting &"'"

		if counting = cint(dbRst("start_hour")) then
 			response.write (" selected ")
		end if

 		response.write ">"& counting &"</option>"

		counting= counting + 1
	loop
%>
					</select>
						<select name="start_min">
<%
	counting=0

	do until counting = 60

 		response.write "<option value='"& counting &"'"

		if counting = cint(dbRst("start_min")) then
 			response.write (" selected ")
		end if

 		response.write ">"& counting &"</option>"

		counting= counting + 5
	loop
%>
						</select>
						<select name="start_am_pm">
<% if dbRst("start_am_pm") = "am" then %>
						<option value="am" selected >AM</option>
						<option value="pm">PM</option>
<% else %>
						<option value="am">AM</option>
						<option value="pm" selected >PM</option>
<% end if %>
						</select>
				</td>
				<td class="whitetext">
					<select name="finish_hour">
<%
	counting=1

	do until counting = 13

 		response.write "<option value='"& counting &"'"

		if counting = cint(dbRst("finish_hour")) then
 			response.write (" selected ")
		end if

 		response.write ">"& counting &"</option>"

		counting= counting + 1
	loop
%>
					</select>
						<select name="finish_min">
<%
	counting=0

	do until counting = 60

 		response.write "<option value='"& counting &"'"

		if counting = cint(dbRst("finish_min")) then
 			response.write (" selected ")
		end if

 		response.write ">"& counting &"</option>"

		counting= counting + 5
	loop
%>
						</select>
						<select name="finish_am_pm">
<% if dbRst("finish_am_pm") = "am" then %>
						<option value="am" selected >AM</option>
						<option value="pm">PM</option>
<% else %>
						<option value="am">AM</option>
						<option value="pm" selected >PM</option>
<% end if %>
						</select>
				</td>

			<tr>
				<td colspan="3">&nbsp;</td></tr>
			<tr>
				<td style="font-size:16px; color:white;">Show Price</td>
				<td style="font-size:16px; color:white;">Money Collection</td>
				<td style="font-size:16px; color:white;">Staff needed</td> 
			</tr>
			<tr>
				<td class="whitetext"><input type="text" name="show_price" size="20" maxlength="50" value="<% Response.Write dbRst("show_price") %>"></td>
				<td class="whitetext">
						<select name="money_collection">
<% if dbRst("finish_am_pm") = "am" then %>
						<option value="Yes" selected >Yes</option>
						<option value="No">No</option>
<% else %>
						<option value="Yes">Yes</option>
						<option value="No" selected >No</option>
<% end if %>
						</select>
				</td>
				<td class="whitetext">
					<select name="staff_needed">
<%
	counting=1

	do until counting = 6

 		response.write "<option value='"& counting &"'"

		if counting = cint(dbRst("staff_needed")) then
 			response.write (" selected ")
		end if

 		response.write ">"& counting &"</option>"

		counting= counting + 1
	loop
%>
					</select>
				</td>
			<tr>
				<td colspan="3">&nbsp;</td></tr>
			<tr>
				<td style="font-size:16px; color:white; text-align:left;" colspan="3">Description</td> 
			</tr>
			<tr>
				<td class="whitetext" colspan="3" style="font-size:16px; color:white; text-align:left;" ><input type="text" name="show_description" size="85" maxlength="100" value="<% Response.Write dbRst("show_description") %>">
			</tr>
			<tr>
				<td colspan="3"><input type="submit" name="Submit" value="submit"></td>
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
%>
		<tr><td class="whitetext">Sorry, There is no show for this Client</td></tr>
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