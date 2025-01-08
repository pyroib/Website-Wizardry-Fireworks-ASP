<%@ Language=VBScript %>
<% Option Explicit %>
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
	<script language=javascript>
		function closewindow(){
			window.close()
		}
	</script>
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
	<body oncontextmenu="closewindow()" onclick="closewindow()">
<%
	DIM page, page_id, strSQL, dbconn, dbRst
	
	page_id = Request.Querystring("id")
	
	FUNCTION dbconnect()
		Set dbConn = Server.CreateObject("ADODB.Connection")
		dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\014753\wizDB.mdb") & ";"
		dbConn.Open
		Set dbRst = dbConn.Execute(strSQL)
	END FUNCTION

	strSQL = "SELECT * FROM services WHERE profileid = '"& page_id & "'"

	dbconnect()
%>
		<table border="0" cellpadding="0" cellspacing="0" width="380" height="400"  style="border: 1px solid #cccccc;">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" class="black" height="400" >
						<tr>
							<td colspan="2" class="center" valign="top" style="text-align:center;"><img src="/images/wiz_logo2.jpg" alt="Wizardry Fireworks Logo" /></td>
						</tr>
						<tr>
							<td valign="top" style="padding-left:10px;">
<% response.Write dbRst("description") %>
							</td>
<% 
			response.Write ("<td class='text' valign='top'><img src='/images/"& page_id &".jpg' align='right' hspace='10px' alt='"& page_id &"' /></td>")
			dbrst.Close
			SET dbrst = NOTHING
%>
						</tr>
						<tr>
							<td colspan="2" class="center" valign="top" style="text-align:center;">Want more information? Visit our Contact us page for details on how.</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>