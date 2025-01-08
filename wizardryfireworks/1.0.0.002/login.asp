<%@ Language=VBScript %>
<% Option Explicit %>
<%

DIM strSQL, dbConn, dbRst, id, staff_id, staff_pw, former

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
								<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc;"><tr><td style="background: #ffffff;">&nbsp;&nbsp;&nbsp;&nbsp;<a href="/default.asp">Home</a> || <a href="/login.asp">Login</a></td></tr></table>
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
															<tr><td style="background: #cccccc; text-align: left; padding-left:8px;" valign="top" height="10">Wizardry Fireworks Login Page</td></tr>
															<tr>
																<td style="background: #ffffff; text-align: left; padding-left:10px; padding-right:10px;" valign="top" height="100%">
<!--- Start login page content //-->
<%
former = Request.Form("Submit")
id = request.querystring("id")

IF id = "logout" THEN
	session.abandon
	Response.redirect ("/default.asp")	
END IF
		IF former = "Login" THEN
	
			staff_id = Request.Form("txtStaff_id")
			staff_id = Lcase(staff_id)
			staff_pw = Request.Form("txtStaff_pw")
			staff_pw = Lcase(staff_pw)
			
			staff_id = Replace(staff_id,"'","")
			staff_pw = Replace(staff_pw,"'","")
			
			strSQL = "SELECT * FROM logins WHERE staff_id ='" & staff_id & "' AND staff_pw='" & staff_pw & "'"
			
			dbconnect()
			
			IF dbRst.EOF AND dbRst.BOF THEN
				'details incorrect
				dbConn.Close
				SET dbConn = NOTHING
				Response.redirect ("/login.asp")
			ELSE
				Session("staff_id") = staff_id
				dbConn.Close
				SET dbConn = NOTHING
				Response.redirect ("/login.asp")
			END IF
		END IF
			IF Session("staff_id") = "" THEN
%>
					<br />
					<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc; text-align:center;">
						<form method="post" action="/login.asp">
							<tr> 
								<td  width="50%" valign="middle" height="50" align="right" class="login">Staff ID: &nbsp;&nbsp;</td>
								<td  width="50%" valign="middle" height="50" align="left"><input type="text" name="txtstaff_Id" size="30" maxlength="25" autocomplete="no" />
							</tr>
							<tr> 
								<td valign="middle" height="50" align="right" class="login">Password: &nbsp;&nbsp;</td>
								<td valign="middle" height="50" align="left"><input type="password" name="txtstaff_pw" size="30" maxlength="25" /></td>
							</tr>
							<tr> 
								<td valign="middle" height="50" align="left"></td>
								<td valign="middle" height="50" align="left"><input type="submit" name="Submit" value="Login" /></td>
							</tr>
							<tr>
								<td colspan="2"></td>
							</tr>
						</form>
					</table>
<%
			ELSE
%>
					<br />
					<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc; text-align:center;">
						<tr> 
							<td  width="50%" valign="middle" height="50" align="center" class="login"><br /><br /><br /><a href="default.asp">You are now logged in as <x style="text-transform:capitalize;"><%= Session("staff_id") %></x>, Click here to continue</a><br /><br /><br /><a href="/login.asp?id=logout">Log Out</a></td>
						</tr>
					</table>
<%
			END IF
%>
<!--- finish login page content //-->
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










