<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
  <head>
  <!--please respect all copyright laws. all content on this site belongs to its rightful owner. //-->
  <!--All scripts, pages and images written and created by Ian Blott. Unless specified//-->
  <!--website created and Maintained by Ian Blott. email me at |ian(a)iblott.com|//-->
  <meta name="author" content="Ian Blott" />
  <meta name="generator" content="100% notepad" />
  <meta name="copyright" content="iblott.com 2007" />
  <meta name="publisher" content="Ian Blott" />
  <meta name="description" content="Portfolio of all things created by Ian Blott" />
  <meta name="keywords" content="Ian, Blott, Portfolio" />
  <meta name="robots" content="Follow,Index" />
  <meta http-equiv="content-language" content="en" />
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <title>Ian Blott's Online Portfolio</title>
  <link rel="stylesheet" type="text/css" href="css/mainpage.css" title="TOCStyle" />
  <script type="text/JavaScript">
  <!-- 
    function NewWindow(URL) {
      day = new Date();
      id = day.getTime();
      eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=400,height=420,left = 100,top = 150');");
    }
  //-->
  </script>
  </head>
<%
' Define all your variables
dim dbConn, dbRst, icount

'I use a function to connect to my database, It makes for simpler coding during the page
FUNCTION dbconnect()
	Set dbConn = Server.CreateObject("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\39587ggjddnee\45554dfdvccxx.mdb") & ";"
	dbConn.Open
	Set dbRst = dbConn.Execute(strSQL)
END FUNCTION


page_id = Request.Querystring("id")

%>
<body>
	<div class="center">
		<div class="top_banner"></div>
		<div class="iblott_logo">
			<img src="images/iblott_logo.jpg" usemap="#Map" border="0" />
			<map name="Map" id="Map">
				<area shape="rect" coords="333,2,383,15" href="portfolio.asp?id=web" />
				<area shape="rect" coords="396,3,465,15" href="portfolio.asp?id=photography" />
				<area shape="rect" coords="473,3,615,16" href="portfolio.asp?id=3d" />
				<area shape="rect" coords="626,2,768,16" href="portfolio.asp?id=flash" />
			</map>

		</div>

		<div class="menu_control">
			<img src="images/ian.jpg" width="240" height="240" alt="Ian Blott pic" />
		</div>

		<div class="showcase_control">
<%
			'Selete only the info from the table portfolio that is required for this section
			strSQL = "SELECT * FROM showcase WHERE code_title = '"& page_id & "'"

			'Run the database connection function
			dbconnect()
			DO WHILE NOT dbRst.EOF 
%>

          	<div class="showcase_details">
<a href="breakout.asp?id=<% Response.Write dbRst("code_title")%>" onclick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onfocus="this.blur()">
	<img src="images/<% Response.Write dbRst("code_title")%>_sc_tn.jpg" alt="<% Response.Write dbRst("title") %>" class="sc_image" />
</a>
<span class="sc_title">
	<% Response.Write dbRst("title") %>
</span>
<br />
<br />
<br />
<span class="sc_header">
Purpose
</span>
<br />
<% Response.Write dbRst("purpose")%><br /><br />
<span class="sc_header">
Programs used
</span>
<br />
<% Response.Write dbRst("program")%><br /><br />
<span class="sc_header">
Estimate time spent
</span>
<br />
<% Response.Write dbRst("est_time")%><br /><br />
<span class="sc_header">
Date
</span>
<br />
<% Response.Write dbRst("complete")%><br /><br />
<span class="sc_header">
Other Information
</span>
<br />
&nbsp;&nbsp;&nbsp;<% Response.Write dbRst("other") %>
</div>
<%
			'tell the database control that i am finished with the currently selected data
			dbRst.MoveNext

			'restart the loop to regenerate the next part of the code
			loop

			'close the database connection
			dbrst.Close

			'reseting variables is allways good practice if your going to reuse them later on in the page.
			SET dbrst = NOTHING

%>
		</div>
		<div class="ian_bio">
			<!--#include file="menu.asp"-->
		</div>
	</div>
</body>
</html>