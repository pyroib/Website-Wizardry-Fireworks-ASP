<%@ Language=VBScript %>
<% Option Explicit %>
<div id="Text" class="text" style="text-align:left; position:absolute; left:300px; top:450px; width:530px; height:200px; z-index:1; border: 1px none #000000; visibility: visable;"></div>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_setTextOfLayer(objName,x,newText) { //v4.01
  if ((obj=MM_findObj(objName))!=null) with (obj)
    if (document.layers) {document.write(unescape(newText)); document.close();}
    else innerHTML = unescape(newText);
}
//-->
</script>

<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Services</div>
	</td>
</tr>
<tr>
	<td height="485" valign="top" align="center" style="padding: 10px; border: 1px ridge #FFFF66;">
		<div name="text"></div>
		<table border="0" width="560" height="485" cellpadding="0" cellspacing="0">
			<tr>
				<td colspan="5" height="20"></td>
			</tr>
			<tr>
				<td colspan="5" style="text-align:left;" class="text" height="20">
				COMPANY PROFILE<br />
				<br />
				Wizardry Fireworks Pty Ltd is a team of trained licensed pyrotechnicians performing fireworks effects that are truly different.<br />
				<br />
				A team that bring in a multitude of skills in many industries and many years of experience.<br />
				With this skill set working together all clients can expect the safest displays and the most unique.<br />
				<br />
				We our proud of what we have achieved in developing a specialised team and a loyal clientele to date and have the biggest pyrotechnic accolades in our sights.<br />
				<br />
				We look forward in making your next event even more special and challenge ourselves to bring out our best.<br />
				<br />
				<br />
				</td>
			</tr>

		<% 
			Dim strSQL
			strSQL = "SELECT * FROM StaffProfiles order BY staff_id ASC"
		%>
			<!-- #include virtual="/DBconnect.wiz"//-->
		<%
			Do While Not (dbRst.EOF)
				If (dbRst.EOF) Then
		%>
			<tr><td></td></tr>
		<%
				ELSE
		%>
			<tr>
				<td width="20"></td>
<td width="186" class="whitetext" style="text-align:center; padding: 5px; border: 1px ridge #FFFF66;" onmouseover="MM_setTextOfLayer('text','',' <% Response.Write (" Name : "& dbRst("Name") &"<br /> Age : "& dbRst("age") &"<br /> Experience : "& dbRst("experience") &" ")%> ')">
		<%
					Response.Write dbRst("Name")
					dbRst.MoveNext
		%>
				</td>
		<%
				END iF
				If (dbRst.EOF) Then
		%>
			<td></td>
		<%
				ELSE
		%>
				<td width="50"></td>
<td width="186" class="whitetext" style="text-align:center; padding: 5px; border: 1px ridge #FFFF66;" onmouseover="MM_setTextOfLayer('text','',' <% Response.Write (" Name : "& dbRst("Name") &"<br /> Age : "& dbRst("age") &"<br /> Experience : "& dbRst("experience") &" ")%> ')">
		<%
					Response.Write dbRst("Name")
		%>
				</td>
				<td width="20"></td>
			</tr>
			<tr>
				<td colspan="5" height="10"></td>
			</tr>
<%
					dbRst.MoveNext
				END IF
			loop
			dbRst.Close
			Set dbRst = Nothing
%>
		<tr><td colspan="5"><br /><br /><br /><br /><br /><br /></td></tr>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="font-size:10px; color:white;">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->

