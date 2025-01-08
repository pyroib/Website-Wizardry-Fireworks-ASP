<%@ Language=VBScript %>
<% Option Explicit %>

<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Wizardry Magic</div>
	</td>
</tr>
<tr>
	<td height="505" valign="top" align="left" style="padding: 0px; border: 1px ridge #FFFF66; text-align:center;">
		<script language="JavaScript" type="">
		<!--
			function moveover(name,suffix,type){
				eval("document."+name+".src='/images/indoor/"+suffix + type+ "'")
			}

			Image1= new Image()
			Image1.src = "/images/indoor001.jpg"
			Image2= new Image()
			Image2.src = "/images/indoor002.jpg"
			Image3= new Image()
			Image3.src = "/images/indoor003.jpg"
			Image4= new Image()
			Image4.src = "/images/indoor004.jpg"
			Image5= new Image()
			Image5.src = "/images/indoor005.jpg"
			Image6= new Image()
			Image6.src = "/images/indoor006.jpg"
			Image7= new Image()
			Image7.src = "/images/indoor007.jpg"
			Image8= new Image()
			Image8.src = "/images/indoor008.jpg"
		// --> 
		</script>
		<table width="590" border="0" cellpadding="0" cellspacing="0">
			<tr>

<%
	dim start, prefix, finish, count, fs

	If Request.QueryString("startat") = "" OR Request.QueryString("startat") < 1 Then
		start = 1
	Else
		start = Request.QueryString("startat")
	End If


	finish = start + 8
	count = 1


	IF start < 10 Then
		prefix = "00"
		else if start > 9 Then
			prefix = "0" 
		End if
	End if
%>
				<td colspan="3"><img src="/images/indoor/indoor<% response.write prefix & start %>.jpg" name="big_one" alt="Large Image" /><br /></td>
			</tr>
			<tr>
				<td width="45" align="center">
<%if start > 2 then%>
				<a href="indoor.asp?startat=<% response.write start - 8%>"><img src="/images/last.gif" alt="Next Page" border="0" /></a>
<%end if%>
				</td>
				<td width="500">
<%

	do until start = finish

	IF start < 10 Then
		prefix = "00"
		else if start > 9 Then
			prefix = "0" 
		End if
	End if

if start < 8 then
%>
					<img src='/images/tn/indoor<% response.write prefix & start %>.jpg' alt="indoor Image<% response.write prefix & start %>" onclick="moveover('big_one','indoor<% response.write prefix & start %>','.jpg')" />
<%
end if 
	if count = "4" then
		response.write ("<br />")
	end if

	start = start + 1
	count = count + 1
loop
%>
				</td>
				<td width="45" align="center">
<%if start < 8 then%>
					<a href="indoor.asp?startat=<% response.write start%>"><img src="/images/next.gif" alt="Next Page" border="0" /></a>
<%end if%>
				</td>
			<tr>
				<td colspan="3">
					<a href="outdoor.asp">Don't forget to take a look at our outdoor display images!</a>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->