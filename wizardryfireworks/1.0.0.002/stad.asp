<%@ Language=VBScript %>
<% Option Explicit %>
<%
DIM strSQL, dbConn, dbRst, random_number

IF Session("staff_id") = "" THEN
Response.redirect ("/login.asp")
ELSE
FUNCTION dbconnect()
	Set dbConn = Server.CreateObject("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\014753\wizDB.mdb") & ";"
	dbConn.Open
	Set dbRst = dbConn.Execute(strSQL)
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
		<meta name="description" content="One of Australia's most creative fireworks companies. With impecable safety and huge attention to detail we are rated amoungst the top in the country.">
		<meta name="keywords" content="Fireworks,Sydney Fireworks,Fireworks Sydney, Sydney Pyrotechnics,Fireworks galleries,Australia,Fireworks Australia,Australian Fireworks,family fireworks,consumer fireworks,Fireworks,Fireworks Distributor,Fireworks Wholesale,Fireworks Supplier,Wholesale Fireworks in australia,wholesale fireworks,fireworks distributors,fireworks displays,wedding fireworks,firework packs,corporate firework displays,brilliant fireworks,fireworks for parties,bonfire night,mail order fireworks,low noise firework displays,finale fireworks,professional fireworks,fireworks for sale,firework retailer,discounted fireworks,quality fireworks,pyrotechnics,chinese new year fireworks,halloween fireworks,diwali fireworks,Olympic Fireworks, Olympic Flame, fireworks, pyrotechnics, pyro, New Years Eve, Indoor Pyrotechnics, Australia, Olympics, Sydney 2000, Sydney Harbour Bridge, Events, Special, Darling Harbour, Corporate Events,NYE,celebration,celebrate,party,bang,visual,eye candy,">
		<meta name="robots" content="Follow,Index">
		<meta http-equiv="content-language" content="en" />
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<title>Wizardry Fireworks Pty Ltd</title>
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
					eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=0,width=400,height=420 left = 340,top = 250');");
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
								<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc;"><tr><td style="background: #ffffff;">&nbsp;&nbsp;&nbsp;&nbsp;<a href="/default.asp" alt="Staff login">Home</a></td></tr></table>
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
															<tr><td style="background: #cccccc; text-align: left; padding-left:8px;" valign="top" height="10">Wizardry Fireworks Staff Admin Page</td></tr>
															<tr>
																<td style="background: #ffffff; text-align: left; padding-left:10px; padding-right:10px;" valign="top" height="100%">
<!--- Start Staff Admin page //-->
					<br />
					<table border="0" width="100%" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc; text-align:center;">
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center"><x style="font-size:18px;">Site Settings</x></td>
							<td width="50" valign="top" align="center"><x style="font-size:18px;">Site Data</x></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center"><a href="/siteset.asp?id=password" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Change Password</a></td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=add_news" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Add News</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=edit_news" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Edit News</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><!--a href="/dataedit.asp?id=add_gallery" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()"//-->Add Gallery Photo<!--/a//--></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><!--a href="/dataedit.asp?id=edit_gallery" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()"//-->Edit Gallery Photo<!--/a//--></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=add_team" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Add Team Member</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=edit_team" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Edit Team Member</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=add_employ" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Add Employment Details</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=edit_employ" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Edit Employment Details</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=add_affiliate" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Add Affiliate Link</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=edit_affiliate" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Edit Affiliate Link</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=add_copro" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Add Company Profile</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center"><a href="/dataedit.asp?id=edit_copro" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Edit Company Profile</a></td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
						<tr>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
							<td width="50" valign="top" align="center">&nbsp;</td>
						</tr>
					</table>
<!--- finish Staff Admin Page //-->
																</td>
															</tr>
															<tr><td style="background: #ffffff; text-align: center;" valign="bottom" height="50">Copyright &copy; 2006 Wizardry Fireworks Pty Ltd All rights Reserved</td></tr>
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
<!--- start right menu //-->


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
<% END IF %>