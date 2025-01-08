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

<% 
Dim strSQL, postid

	If Request.QueryString("ID") = "" Then
		response.write "Sorry, No News Selected"
	Else
		postid = Request.QueryString("ID")
	End If

strSQL = "SELECT * FROM PublicNews WHERE id = "& Request.QueryString("ID") & ""

%>
			<!-- #include virtual="/DBconnect.wiz"//-->
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
			<table border="1" width="550">
<%
	Do While Not dbRst.EOF
%>



				<tr> 
					<td colspan="2" align="left" valign="top">Topic : <br /><% Response.Write dbRst("Topic")%></td>
					<td width="350" rowspan="2" valign="top" align="left"><% Response.Write dbRst("News")%></td>
				</tr>
				<tr> 
					<td width="100" valign="top">Posted By : <br /><% Response.Write dbRst("PostedBy")%></td>
					<td valign="top">Posted On : <br /><% Response.Write dbRst("Date")%></td>
				</tr>

<%
	dbRst.MoveNext
	loop

%>
			</table>
			<table>
				<tr>
					<td valign="bottom"><a href="/news">Back to News page</a></td>
				</tr>
			</table>
		</div>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->
