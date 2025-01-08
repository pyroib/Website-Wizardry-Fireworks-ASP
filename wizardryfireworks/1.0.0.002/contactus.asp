<%@ Language=VBScript %>
<% Option Explicit %>
<%

DIM strSQL, dbConn, dbRst, string, icount, random_number, contact_name, email, enquiry, strHTML, howifind

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
<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc;">
<tr>
<td style="background: #ffffff;">
&nbsp;&nbsp;&nbsp;&nbsp;<a href="/default.asp">Home</a>
<%
IF Request.QueryString("id") = "" THEN
	%> || <a href="/contactus.asp">Contact Us</a><%
ELSE
	id = request.querystring("ID")
	%> || <a href="/contactus.asp">Contact Us || <a href="/contactus.asp?id=<%= id %>"><x style="text-transform:capitalize;"><%= id %></x></a><%
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
															<tr><td style="background: #cccccc; text-align: left; padding-left:8px;" valign="top" height="10">Contacting Wizardry Fireworks</td></tr>
															<tr>
																<td style="background: #ffffff; text-align: left; padding-left:10px; padding-right:10px;" valign="top" height="100%">
                                                                    <table border="0" width="100%" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc; text-align:center;">
                                                                        <tr><td height="20" colspan="2" style="text-align:center;"></td></tr>
                                                                        <tr>
                                                                            <td style="text-align:center;"><x style="font-size:18px;">Mail Address</x></td>
                                                                            <td style="text-align:center;"><x style="font-size:18px;">Phone</x></td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="text-align:center; top:0px">
                                                                                PO BOX 95<br />Baulkham Hills 1755<br />N.S.W. Sydney Australia<br />
                                                                            </td>
                                                                            <td style="text-align:center;">
                                                                                <x style="font-size:15px;">Office Phone:</x><br />+61 (02) 9686 1999<br />
                                                                                <br />
                                                                                <x style="font-size:15px;">Office Fax:</x><br />+61 (02) 9686 9191<br />
                                                                            </td>
                                                                        </tr>
                                                                        <tr><td height="20" colspan="2" style="text-align:center;"></td></tr>
                                                                    </table>
                                                                    <br /><br />

<!--- Start main contact us content //-->

<%
dim id
id = request.querystring("id")
IF id = "form" THEN
%>
				<form method="post" name="signup" action="/contactus.asp?id=submit">
					<table border="0" width="100%" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc; text-align:center;">
						<tr><td colspan="2"><x style="font-size:18px;">Contacting Wizardry Fireworks</x></td></tr>
						<tr><td colspan="2" height="20"></td></tr>

						<tr>
							<td width="200" align="right">Name : </td>
							<td width="200" align="left">
								<input name="contact_name" type="text" size="30" maxlength="40" />
							</td>
						</tr>
						<tr><td height="12" colspan="2"></td></tr>
						<tr>

							<td align="right">Email Address : </td>
							<td align="left">
								<input name="email" type="text" size="30" maxlength="50" />
							</td>
						</tr>
						<tr><td height="12" colspan="2"></td></tr>
						<tr>
							<td align="right" valign="top">How did you Hear about Wizardry? : </td>

							<td align="left">
								<select name="how" multiple width="20" size="5">
									<option value="Phone Book">Phone Book (Yellow / White Pages)</option>
									<option value="Rec. by a F.Centre">Recomended by Function Centre</option>
									<option value="Rec. by a Friend">Recomended by a Friend</option>
									<option value="Internet Search">Internet Search Engine (google etc)</option>
									<option value="Saw the Trailer">Saw your trailer on the Road</option>

									<option value="Trade show">Trade Show</option>
									<option value="Magazine">Magazine add</option>
									<option value="I saw a previous show">I saw a show you put on</option>
									<option value="Affiliate link">Affiliate Link</option>
									<option value="Other">Other</option>
								</select>

							</td>
						</tr>
						<tr><td height="12" colspan="2"></td></tr>
						<tr>
							<td align="right" valign="top">Enquiry :</td>
							<td align="left">
								<textarea name="enquiry" cols="30" rows="10" name="textarea"></textarea>
							</td>

						</tr>
						<tr><td colspan="2">&nbsp;</td><tr>
						<tr>
							<td colspan="2">
								<input name="submit" type="submit" value="submit" />
							</td>
						</tr>
						<tr>
							<td colspan="4">&nbsp;</td>

						</tr>
					</table>
				</form>
<%
				ELSE IF id = "submit" THEN
					IF Request.form("email") <> "" THEN
					
						contact_name = Request.form("contact_name")
						contact_name = Replace(contact_name,"'","")
						contact_name = Replace(contact_name,"""","")
						
						email = Request.form("email")
						email = Replace(email,"'","")
						email = Replace(email,"""","")
											
						enquiry = Request.form("enquiry")
						enquiry = Replace(enquiry,"'","")
						enquiry = Replace(enquiry,"""","")

						howifind = Request.form("how")

						Const cdoSendUsingMethod = _
						"http://schemas.microsoft.com/cdo/configuration/sendusing"
						Const cdoSendUsingPort = 2
						Const cdoSMTPServer = _
						"http://schemas.microsoft.com/cdo/configuration/smtpserver"
						Const cdoSMTPServerPort = _
						"http://schemas.microsoft.com/cdo/configuration/smtpserverport"
						Const cdoSMTPConnectionTimeout = _
						"http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
						Const cdoSMTPAuthenticate = _
						"http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
						Const cdoBasic = 1
						Const cdoSendUserName = _
						"http://schemas.microsoft.com/cdo/configuration/sendusername"
						Const cdoSendPassword = _
						"http://schemas.microsoft.com/cdo/configuration/sendpassword"


						Dim objConfig ' As CDO.Configuration
						Dim objMessage ' As CDO.Message
						Dim Fields ' As ADODB.Fields


						Set objConfig = Server.CreateObject("CDO.Configuration")
						Set Fields = objConfig.Fields
	
						With Fields
							.Item(cdoSendUsingMethod) = cdoSendUsingPort
							.Item(cdoSMTPServer) = "mail.wizardryfireworks.com"
							.Item(cdoSMTPServerPort) = 25
							.Item(cdoSMTPConnectionTimeout) = 20
							.Item(cdoSMTPAuthenticate) = cdoBasic
							.Item(cdoSendUserName) = "site@wizardryfireworks.com"
							.Item(cdoSendPassword) = "pyr0sg"
			
							.Update
						End With

						Set objMessage = Server.CreateObject("CDO.Message")
						Set objMessage.Configuration = objConfig

					
						With objMessage
							.To = "Wizardry Fireworks <info@wizardryfireworks.com>"
							.bcc = "Wizardry enquiry <webmaster@wizardryfireworks.com>; Ian Blott <ian_blott@hotmail.com>"
							.From = "Site Info Request <site@wizardryfireworks.com>"
							.Subject = "Request For Information From Site User"
							strHTML = "<x style='font-size:18px;'>Enquiry -</x><br />"& enquiry &" <br /><br />How i found the site -</x><br />"& howifind &" <br /><br /><x style='font-size:18px;'>Name -</x><br /> "& contact_name &" <br /><br /><x style='font-size:18px;'>Email Address -</x><br /> "& email &""
							.HTMLBody = strHTML
							.Send
						End With

						Set Fields = Nothing
						Set objMessage = Nothing
						Set objConfig = Nothing 
						Response.write("<table align='center'><tr><td align='center'><br /><br /><br /><br /><br /><a href='/contactus.asp?id=email'>Thank you for contacting Wizardry Fireworks<br />Your query has been emailed to us and we will endevour to answer it as quickly as possible.<br /><br /><br /><br />Click Here to return to the Contacts Page</a></td></tr></table>")
					ELSE
						Response.write("<a href='/contactus.asp?id=form'>Sorry, You need to provide an email address.</a>")
					END IF
				ELSE
%>
		<br />
		<table border="0" width="100%" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc; text-align:center;">
			<tr><td colspan="2" style="text-align:center;"><x style="font-size:18px;"> E-mail</x></td></tr>
			<tr>
				<td style="text-align:center;">
				<script type="text/javascript">
				<!--
					var crawlerbreak= new Array("info","@","wizardryfireworks",".com")
					var mymail = ''
					for (i=0;i<crawlerbreak.length;i++)
					mymail = mymail + crawlerbreak[i]
						document.write('<a href="mailto:'+mymail+'">'+mymail+'</a>')
				//-->
				</script>
				</td>
				<td style="text-align:center;">General Enquiries</td>
			</tr>
			<tr>
				<td style="text-align:center;">
				<script type="text/javascript">
				<!--
					var crawlerbreak= new Array("site","@","wizardryfireworks",".com")
					var mymail = ''
					for (i=0;i<crawlerbreak.length;i++)
					mymail = mymail + crawlerbreak[i]
						document.write('<a href="mailto:'+mymail+'">'+mymail+'</a>')
				//-->
				</script>
				</td>
				<td style="text-align:center;">Site Enquiries</td>
			</tr>
			<tr>
				<td style="text-align:center;">
				<script type="text/javascript">
				<!--
					var crawlerbreak= new Array("greg","@","wizardryfireworks",".com")
					var mymail = ''
					for (i=0;i<crawlerbreak.length;i++)
					mymail = mymail + crawlerbreak[i]
						document.write('<a href="mailto:'+mymail+'">'+mymail+'</a>')
				//-->
				</script>
				</td>
				<td style="text-align:center;">Greg Coorey - Head Pyrotechnician</td>
			</tr>
			<tr>
				<td style="text-align:center;">
				<script type="text/javascript">
				<!--
					var crawlerbreak= new Array("webmaster","@","wizardryfireworks",".com")
					var mymail = ''
					for (i=0;i<crawlerbreak.length;i++)
					mymail = mymail + crawlerbreak[i]
						document.write('<a href="mailto:'+mymail+'">'+mymail+'</a>')
				//-->
				</script>
				</td>
				<td style="text-align:center;">Ian Blott - Web Designer</td>
			</tr>
			<tr><td height="20" colspan="2" style="text-align:center;"></td></tr>
			<tr>
				<td style="text-align:center;"><x style="font-size:18px;">Mail Address</x></td>
				<td style="text-align:center;"><x style="font-size:18px;">Phone</x></td>
			</tr>
			<tr>
				<td style="text-align:center; top:0px">
					PO BOX 95<br />Baulkham Hills 1755<br />N.S.W. Sydney Australia<br />
				</td>
				<td style="text-align:center;">
					<x style="font-size:15px;">Office Phone:</x><br />+61 (02) 9686 1999<br />
					<br />
					<x style="font-size:15px;">Office Fax:</x><br />+61 (02) 9686 9191<br />
				</td>
			</tr>
			<tr><td height="20" colspan="2" style="text-align:center;"></td></tr>
			<tr><td colspan="2" style="text-align:center;"><x style="font-size:18px;">Online</x></td></tr>
			<tr><td colspan="2" style="text-align:center;"><a href="/contactus.asp?id=form">Click Here</a></td></tr>

		</table>
<%
				END IF
			END IF
%>
<!--- finish main contact us content //-->
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
							<td style="text-align:center;" valign="top">
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










