<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;News</div>
	</td>
</tr>
<tr>
	<td height="485" valign="top" align="center" style="padding: 10px; border: 1px ridge #FFFF66;">
		<div style="font-size:12px; color:white;">
			<img src="/images/logosmall.gif" alt="Wizardry Logo" /><br /><br />
				This page Is sorted by Time of posting<br /><br />
				The newest Posts are at the top, Feel free to browse the Archives at your own will<br /><br />
<% 
Dim strSQL, start, a, b, display

	If Request.QueryString("startat") = "" Then
		start = -1
	Else
		If Request.QueryString("startat") < 0 Then
			start = -1
		else
			start = Request.QueryString("startat")
		End if
	End If
strSQL = "SELECT * FROM publicNews ORDER BY date DESC"
%>
<!-- #include virtual="/DBconnect.wiz"//-->
		<table border="1" width="500">
			<tr>
				<td class="newstitle"> Topic : </td>
				<td class="newstitle"> Posted by : </td>
				<td class="newstitle"> On : </td>
			</tr>
<%
if (start = "all") then
	Do While Not dbRst.EOF
%>
			<tr>
				<td  height="22"><a href="/staff_only/staff_news.asp?id=<% Response.Write dbRst("id")%>"><% Response.Write dbRst("Topic")%></a> </td>
				<td><% Response.Write dbRst("PostedBy")%> </td>
				<td><% Response.Write dbRst("Date")%> </td>
			</tr>
<%
	dbRst.MoveNext
	loop
%>
<%
else
	a = 0
	b = 0
	display = 10
	Do While Not dbRst.EOF AND (b < display)
 		if (a < start + 1) Then
			a = a + 1
			dbRSt.MoveNext
			else
%>
			<tr>
				<td height="22"><a href="/news/view.asp?id=<% Response.Write dbRst("id")%>"><% Response.Write dbRst("Topic")%></a></td>
				<td><% Response.Write dbRst("PostedBy")%> </td>
				<td><% Response.Write dbRst("Date")%> </td>
			</tr>
<%
				a = a + 1
				b = b + 1
				dbRst.MoveNext
		end if
	loop
%>
			<tr>
				<td colspan="3"><p style="text-align:center}">
<%
	if start > 0 then
%>
					<a href="/news/default.asp?startat=<%response.write start - 10 %>">View Previous page</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
	end if
	if not dbRst.EOF then
%>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="/news/default.asp?startat=<%response.write start + 10 %>">View Next page</a></p>
<%
	end if
%>
				</td>
			</tr>
<%
	dbrst.Close
	Set dbrst = Nothing
end if
%>
		</table>
		</div>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->
