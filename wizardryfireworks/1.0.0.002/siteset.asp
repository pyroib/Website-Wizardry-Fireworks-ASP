<%@ Language=VBScript %>
<% Option Explicit %>
<%
IF Session("staff_id") = "" THEN
Response.redirect ("/login.asp")
ELSE
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
	<head>
		<!--please respect all copyright laws. all content on this site belongs to Wizardry Fireworks PTY LTD. //-->
		<!--All scripts, pages and images written and created by Ian Blott. Unless specified//-->
		<!--website created and mantained by Ian Blott. email me at |ian(a)wizardryfireworks.com|//-->
		<meta name="author" content="Ian Blott" />
		<meta name="generator" content="100% notepad" />
		<meta name="copyright" content="copyright 2003 Wizardry Fireworks PTY LTD" />
		<meta name="publisher" content="Ian Blott" />
		<meta http-equiv="content-language" content="en" />
		<title>Wizardry Fireworks</title>
	</head>
	<style>

		body {
			font-family: Arial, Helvetica, Sans Serif;
			font-size: 11px;
			color: #666666;
			background: #ffffff;
			margin-top:6px;
			margin-left:9px;
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
	<script language=javascript>
		function closewindow(){
			window.close()
		}
	</script>
<%

DIM page, page_id, strSQL, dbconn, dbRst
	
	page_id = Request.Querystring("id")
	
	FUNCTION dbconnect()
		Set dbConn = Server.CreateObject("ADODB.Connection")
		dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\014753\wizDB.mdb") & ";"
		dbConn.Open
		Set dbRst = dbConn.Execute(strSQL)
	END FUNCTION

dim staffnewpw, sess, strHTML
staffnewpw = Request.form("txtstaff_pw")
sess = Session("staff_id")

	IF Request.Querystring("id") = "password" THEN
		IF Request.form("Submit") = "Change Password" THEN
			strSQL = "UPDATE logins Set staff_PW = '"& staffnewpw & "' Where staff_ID = '"& sess & "'"
			dbconnect()
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
							.Subject = "Staff Login Password has changed for wizardryfireworks.com"
							strHTML = "<x style='font-size:18px;'>The Password for www.wizardryfireworks.com staff login has been changed to "& staffnewpw &"</x>"
							.HTMLBody = strHTML
							.Send
						End With
			Set Fields = Nothing
			Set objMessage = Nothing
			Set objConfig = Nothing
		ELSE
%>
		<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc; text-align:center;">
			<form method="post" action="/siteset.asp?id=password">
				<tr> 
					<td valign="middle" height="50" align="right">Password: &nbsp;&nbsp;</td>
					<td valign="middle" height="50" align="left"><input type="password" name="txtstaff_pw" size="30" maxlength="25" /></td>
				</tr>
				<tr> 
					<td colspan="2" valign="middle" height="50" align="center"><input type="submit" name="Submit" value="Change Password" /></td>
				</tr>
				<tr>
					<td colspan="2"></td>
				</tr>
			</form>
		</table>
<%
		END IF
	ELSE
%>
		<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #cccccc; text-align:center;">
			<tr> 
				<td valign="middle" height="150" align="center">Theres not much to do Here. You've hit something weird</td>
			</tr>
		</table>
<%
	END IF
%>
	</body>
</html>
		<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-7750574-1");
pageTracker._trackPageview();
} catch(err) {}</script>
<% END IF %>