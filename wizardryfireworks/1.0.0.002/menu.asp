	<a href="default.asp"> HOME</a><br /><br />
<%
'******************************************************************************************************************************
'****************************************------------------WEB DESIGN------------------****************************************
'******************************************************************************************************************************
%>
	<strong>Website Design</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='web' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
         - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************




'******************************************************************************************************************************
'****************************************-----------------FLASH DESIGN-----------------****************************************
'******************************************************************************************************************************
%>
	<strong>Flash Animation / Games</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='flash' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
          - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************




'******************************************************************************************************************************
'****************************************-------------------3D DESIGN------------------****************************************
'******************************************************************************************************************************
%>
	<strong>3D Animation / Scenes</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='3d' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
         - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************




'******************************************************************************************************************************
'****************************************---------------PHOTOSHOP DESIGN---------------****************************************
'******************************************************************************************************************************
%>
	<strong>Photoshop Design / Manipulation</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='photoshop' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
         - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************




'******************************************************************************************************************************
'****************************************-------------------ASP DESIGN-----------------****************************************
'******************************************************************************************************************************
%>
	<strong>ASP Scripts</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='asp' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
         - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************




'******************************************************************************************************************************
'****************************************------------------PHP DESIGN-----------------****************************************
'******************************************************************************************************************************
%>
	<strong>PHP Scripts</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='php' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
         - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************




'******************************************************************************************************************************
'****************************************-----------------QBASIC DESIGN----------------****************************************
'******************************************************************************************************************************
%>
	<strong>QBasic programs</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='qbasic' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
         - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************




'******************************************************************************************************************************
'****************************************--------------PHOTOGRAPHY DESIGN--------------****************************************
'******************************************************************************************************************************
%>
	<strong>Photography Galleries</strong><br />
<%
'Selete only the info from the table portfolio that is required for this section
strSQL = "SELECT * FROM portfolio WHERE class='photography' ORDER by title"

'Run the database connection function
dbconnect()
DO WHILE NOT dbRst.EOF
%>
         - <a href="showcase.asp?id=<% Response.Write dbRst("code_title") %>"><% Response.Write dbRst("title") %></a><br />
<%

dbRst.MoveNext

loop

'close the database connection
dbrst.Close

'reset variables is allways good practice if your going to reuse them later on in the page.
SET dbrst = NOTHING
'******************************************************************************************************************************

%>




