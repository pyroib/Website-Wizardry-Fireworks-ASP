<%@ Language=VBScript %>
<% Option Explicit %>
<%

DIM strSQL, dbConn, dbRst, string, icount, random_number, id, startat, b, prefix, linetotal, id2, c

FUNCTION dbconnect()
	Set dbConn = Server.CreateObject("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\014753\wizDB.mdb") & ";"
	dbConn.Open
	Set dbRst = dbConn.Execute(strSQL)
END FUNCTION

FUNCTION prefix_test()
	IF b < 10 THEN
		prefix = "00"
		ELSE IF b > 9 THEN
			prefix = "0" 
		END IF
	END IF
END FUNCTION

FUNCTION random_me()
	randomize
	random_number=int(rnd*10)+1
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
		<meta name="description" content="One of Australia's most creative fireworks companies.">
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

				function moveover(name,suffix,type){
					eval("document."+name+".src='/images/"+suffix + type+ "'")
				}
				
				function moveout(name,suffix,type){
					eval("document."+name+".src='/images/"+suffix + type+ "'")
				}
			// --> 
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
	%> || <a href="/gallery.asp">Galleries</a><%
ELSE
	id = request.querystring("ID")
	%> || <a href="/gallery.asp">Galleries</a> || <a href="/gallery.asp?id=<%= id %>"><x style="text-transform:capitalize;"><%= id %></x></a><%
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
															<tr><td style="background: #cccccc; text-align: left; padding-left:8px;" valign="top" height="10">Wizardry Fireworks Image Gallery</td></tr>
															<tr>
																<td style="background: #ffffff; text-align: left; padding-left:10px; padding-right:10px;" valign="top" height="100%">
<!--- Start Galleries content //-->
		<br />
					<table border="0" width="100%" cellpadding="5" cellspacing="0" style="border: 1px solid #cccccc; text-align:left;">
<%
			id = request.QueryString("id")
			IF Request.QueryString("id") = "" THEN
%>
						<tr>
							<td valign="top" align="center">
								<img src="/images/indoor_outdoor_choice.jpg" alt="Please Choose Indoor or OutDoor Images" border="0" usemap="#Map" />
								<map name="Map" id="Map">
									<area shape="rect" alt="" coords="0,276,431,498" href="/gallery.asp?menu=gallery&id=indoor">
									<area shape="rect" alt="" coords="0,0,431,276" href="/gallery.asp?menu=gallery&id=outdoor">
								</map>
							</td>
						</tr>
					</table>
<%
			ELSE
	
				IF id = "indoor" THEN
					id2 = "outdoor"
					c = 6
				ELSE
					IF id ="outdoor" THEN
						id2 = "indoor"
						c = 36
					END IF
				END IF
			
				IF Request.QueryString("startat") < 1 THEN
					startat = 1
					ELSE
						startat = Request.QueryString("startat")
				END IF
			
%>
				<table border="0" width="100%" cellpadding="5" cellspacing="0" style="border: 1px solid #cccccc; text-align:left;">
					<tr>
						<td colspan="3" align="center" class="gal_thumbs">
<% 
						b = startat
						prefix_test()
						response.write "<img src='/images/"& id &"/"& id & prefix & b &".jpg' name='big_one' alt='Image' width='400px'/>"
%>
						</td>
					</tr>
					<tr>
						<td width="40" class="text">
							<% IF startat > 2 THEN %>
								<a href="/gallery.asp?menu=gallery&id=<%= id %>&startat=<% response.write startat - 10 %>"><img src="/images/last.jpg" alt="Previous Page" border="0" /></a>
							<% END IF %>
						</td>
						<td align="center">
<%
		
				DO UNTIL linetotal = 3
					prefix_test()		
%>
					<a href="javascript:moveover('big_one','<%= id %>/<%= id & prefix & b %>','.jpg')"><img src="/images/<%= id %>/tn/<%= id & prefix & b %>.jpg" alt="" border="0" /></a>
<%
					linetotal = linetotal + 1
					b = b + 1
				LOOP
				
				Response.Write "<br />"
				
				DO UNTIL linetotal = 9
					prefix_test()
%>
					<a href="javascript:moveover('big_one','<%= id %>/<%= id & prefix & b %>','.jpg')"><img src="/images/<%= id %>/tn/<%= id & prefix & b %>.jpg" alt="" border="0" /></a>
<%
					linetotal = linetotal + 1
					b = b + 1
				LOOP
				
				startat = startat + linetotal
%>
						</td>
						<td width="40">						
<%
			 if b < c THEN 
%>
								<a href="/gallery.asp?menu=gallery&id=<%= id %>&startat=<%= startat %>"><img src="/images/next.jpg" alt="Next Page" border="0" /></a>
<%  			end if
%>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="3">
							<a href="/gallery.asp?menu=gallery&id=<%= id2 %>">Don't forget to take a look at our <%= id2 %> display images!</a>
						</td>
					</tr>
				</table>
<%
				'Set folder = Nothing
				'Set files = Nothing
				'Set fso = Nothing
			END IF
%>
<!--- finish galleries content //-->
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
        		<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-7750574-1");
pageTracker._trackPageview();
} catch(err) {}</script>
	</body>
</html>










