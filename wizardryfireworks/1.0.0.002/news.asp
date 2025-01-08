<%@ Language=VBScript %>
<% Option Explicit %>
<%

DIM strSQL, dbConn, dbRst, string, icount, random_number, id

FUNCTION dbconnect()
	Set dbConn = Server.CreateObject("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\014753\wizDB.mdb") & ";"
	dbConn.Open
	Set dbRst = dbConn.Execute(strSQL)
END FUNCTION

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"> 
<html>
	<head>
		<!--please respect all copyright laws. all content on this site belongs to Wizardry Fireworks PTY LTD. //-->
		<!--All scripts, pages and images written and created by Ian Blott. Unless specified//-->
		<!--website created and Maintained by Ian Blott. email me at |ian(a)wizardryfireworks.com|//-->
		<meta name="author" content="Ian Blott" />
		<meta name="generator" content="100% notepad" />
		<meta name="copyright" content="copyright 2006 Wizardry Fireworks PTY LTD" />
		<meta name="publisher" content="Ian Blott" />
		<meta name="description" content="One of Australia's most creative fireworks companies. With impecable safety and huge attention to detail we are rated amoungst the top in the country.">
		<meta name="keywords" content="Fireworks,Sydney Fireworks,Fireworks Sydney, Sydney Pyrotechnics,Fireworks galleries,Australia,Fireworks Australia,Australian Fireworks,family fireworks,consumer fireworks,Fireworks,Fireworks Distributor,Fireworks Wholesale,Fireworks Supplier,Wholesale Fireworks in australia,wholesale fireworks,fireworks distributors,fireworks displays,wedding fireworks,firework packs,corporate firework displays,brilliant fireworks,fireworks for parties,bonfire night,mail order fireworks,low noise firework displays,finale fireworks,professional fireworks,fireworks for sale,firework retailer,discounted fireworks,quality fireworks,pyrotechnics,chinese new year fireworks,halloween fireworks,diwali fireworks,Olympic Fireworks, Olympic Flame, fireworks, pyrotechnics, pyro, New Years Eve, Indoor Pyrotechnics, Australia, Olympics, Sydney 2000, Sydney Harbour Bridge, Events, Special, Darling Harbour, Corporate Events,NYE,celebration,celebrate,party,bang,visual,eye candy,">
		<meta name="robots" content="Follow,Index">
		<meta http-equiv="content-language" content="en" />
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<title>Wizardry Fireworks PTY LTD</title>
		<style>
			body {
				font-family: Arial, Helvetica, Sans Serif;
				font-size: 11px;
				color: #666666;
				background: #ffffff;
			}

			a:link, a:visited {
				color: #333333; 
				text-decoration: none;
				font-weight: bold;
			}

			a:hover {
				color: #666666;	
				text-decoration: none;
				font-weight: bold;
}
		</style>
		<script language="JavaScript" type="text/JavaScript">
			<!-- 
				function NewWindow(URL) {
					day = new Date();
					id = day.getTime();
					eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=400,height=420,left = 340,top = 250');");
				}
			-->
		</script>
	</head>
	<body>
		<table cellpadding="0" cellspacing="2" style="border: 1px solid #cccccc;" align="center" >
			<tr>
				<td style="background: #ffffff;">
<!---  start footprint and login //-->
					<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<tr>
							<td width="650">
<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc;">
<tr>
<td style="background: #ffffff;">
&nbsp;&nbsp;&nbsp;&nbsp;<a href="/default.asp">Home</a>
<%
IF Request.QueryString("id") = "" THEN
	%> || <a href="/news.asp">News</a><%
ELSE
	id = request.querystring("ID")
	strSQL = "SELECT * FROM publicNews WHERE id = "& id &""
	dbconnect()
%> || <a href="/news.asp">News</a> || <a href="/news.asp?id=<%= id %>"><%= dbRst("topic") %></a><%
	dbrst.Close
	SET dbrst = NOTHING
END IF
%>
</td>
</tr>
</table>
							</td>
							<td width="2"></td>
							<td>
<!--#include file="logwin.add"-->
							</td>
						</tr>
					</table>
<!---  finish footprint and login //-->

					<table cellpadding="0" cellspacing="0" height="2"><td></td></tr></table>
<!---  start main frames//-->
					<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<tr>
							<td width="650" valign="top">
<!--#include file="logolarge.add"-->
<!--- Start main content window //-->
								<table cellpadding="0" cellspacing="0" height="2"><td></td></tr></table>
								<table border="0" cellpadding="0" cellspacing="1" width="100%" style="background: url(/images/swoosh.jpg) no-repeat; background-position: bottom left;">
									<tr>
<!--- Start menu window //-->
										<td valign="top">
											<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc;">
												<tr>
													<td style="background: #ffffff; text-align: center;" valign="top">
<!--#include file="menu.add"-->
													</td>
												</tr>
											</table>
										</td>
<!--- finish menu window //-->
										<td width="1"></td>
										<td width="500" valign="top">
											<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc;">
												<tr>
													<td style="background: #ffffff; text-align: center;" valign="top">
														<table border="0" width="100%" cellpadding="1" cellspacing="0" height="680">
															<tr><td style="background: #cccccc; text-align: left; padding-left:8px;" valign="top" height="10">Wizardry Fireworks News Articals</td></tr>
															<tr>
																<td style="background: #ffffff; text-align: left; padding-left:10px; padding-right:10px;" valign="top" height="100%">
<!--- Start main page news content //-->
<%
IF Request.QueryString("id") = "" THEN
	strSQL = "SELECT * FROM publicNews ORDER BY dateposted DESC"
	dbconnect()
	icount=0
	DO WHILE NOT dbRst.EOF AND (icount < 20)
		icount = icount + 1
%>
		<x style="font-size:18px;"><% Response.Write dbRst("topic") %></x>
		<x style="font-size:12px;"> - Date <% Response.Write dbRst("dateposted") %> by - <% Response.Write dbRst("postedBy") %></x>
		<br /><br />
<%
		string = dbRst("news")
		Response.Write (left(string,76))
%>
		... <a href="/news.asp?id=<% Response.Write dbRst("id") %>"> Read More &nbsp;<img src="/images/plus.jpg" border="0" alt="Read more about <% Response.Write dbRst("topic") %>" /></a><br /><br />
<%
		dbRst.MoveNext
	loop
ELSE
	id = request.querystring("ID")
	strSQL = "SELECT * FROM publicNews WHERE id = "& id &""
	dbconnect()
%>
	<x style="font-size:18px;"><% Response.Write dbRst("topic") %></x>
	<x style="font-size:12px;"> - Date <% Response.Write dbRst("dateposted") %> by - <% Response.Write dbRst("postedBy") %></x>
	<br /><br />
	<% Response.Write dbRst("news") %>
	<br /><br /><br />
	<x style="font-size:12px;"><a onclick="history.back()">BACK</a></x>
	<br /><hr>
<%
END IF
dbrst.Close
SET dbrst = NOTHING
%>
<!--- finish main page news content //-->
																</td>
															</tr>
															<tr><td style="background: #ffffff; text-align: center;" valign="bottom" height="50">Copyright &copy; 2006 Wizardry Fireworks PTY LTD All rights Reserved</td></tr>
															<tr><td style="background: #cccccc; text-align: center;" valign="top" height="5"></td></tr>
														</table>
													</td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
<!--- finish main content window //-->
							</td>
							<td width="2"></td>
							<td align="center" style="text-align:center;" valign="top">
<!--#include file="rightstuff.add"-->
							</td>
						</tr>
					</table>
<!---  finish main frames//-->
				</td>
			</tr>
		</table>
	</body>
    		<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-7750574-1");
pageTracker._trackPageview();
} catch(err) {}</script>
</html>










