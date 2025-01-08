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
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;News</div>
	</td>
</tr>
<tr>
	<td height="485" valign="top" align="center" style="{padding: 10px; border: 1px ridge #FFFF66;}">
		<div style="{font-size:22px; color:white;}">
<% 
	Dim newsid, a, b, display, strSQL

	If Request.QueryString("id") > 0 Then
		newsid = Request.QueryString("id")
	Else
		newsid = "0"
	End If

	strSQL = "SELECT * FROM news ORDER BY date DESC"

	if newsid > 0 then
		strSQL = "SELECT * FROM news WHERE id = "& Request.QueryString("ID") & ""
	end if

%>
	<!-- #include virtual="/DBconnect.wiz"//-->
<%
	if newsid = 0 then
%>
			<table border="1" width="500">
				<tr>
					<td class="newstitle"> Date : </td>
					<td class="newstitle"> Topic : </td>
					<td class="newstitle"> Posted By : </td>
				</tr>
<%
		Do While Not (dbRst.EOF)
%>
				<tr>
 					<td class="whitetext"><% Response.Write dbRst("Date") %></td>
					<td class="whitetext"><a href="/staff_only/news.asp?id=<% response.write dbRst("id") %>"><% Response.Write dbRst("Topic") %></a></td>
					<td class="whitetext"><% Response.Write dbRst("PostedBy") %></td>
				</tr>

<%
		dbRst.MoveNext
		loop
	else
%>
			<table border="1" width="500">
			<tr>
				<td colspan="3">
					<SCRIPT LANGUAGE="Javascript">
					<!--  
						function image() { };  
							image = new image(); 
							number = 0;  

							image[number++] = "<img src='/Images/outdoor/outdoor001.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor002.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor004.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor005.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor006.jpg' border='0'>"  
							image[number++] = "<img src='/Images/outdoor/outdoor007.jpg' border='0'>"  
							image[number++] = "<img src='/Images/outdoor/outdoor008.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor009.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor010.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor011.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor012.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor013.jpg' border='0'>" 
							image[number++] = "<img src='/Images/outdoor/outdoor014.jpg' border='0'>"  

							increment = Math.floor(Math.random() * number);  document.write(image[increment]);  
						//-->
						</SCRIPT>
					</td>
				</tr>
				<tr>
					<td class="newstitle"> Date : </td>
					<td class="newstitle"> Topic : </td>
					<td class="newstitle"> Posted By : </td>
				</tr>
<%
		Do While Not (dbRst.EOF)
%>
				<tr>
 					<td class="whitetext"><% Response.Write dbRst("Date") %></td>
					<td class="whitetext"><% Response.Write dbRst("Topic") %></td>
					<td class="whitetext"><% Response.Write dbRst("PostedBy") %></td>
				</tr>
				<tr>
 					<td class="whitetext" colspan ="3"><% Response.Write dbRst("news") %></td>
				</tr>
<% 
		dbRst.MoveNext
		loop
	end if 
	dbrst.Close
%>
			<tr>
				<td colspan="4" class="whitetext"><a href="/staff_only"> Back to Staff Page </a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="/staff_only/news.asp"> Back to Staff News Page </a></td>
			</tr>
			</table>
			<br />
		</div>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="{font-size:10px; color:white;}">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->

%>