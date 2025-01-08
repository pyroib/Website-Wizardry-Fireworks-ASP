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
	<body>
<%
DIM affilname, affillink, affilimage, affilbulk, affilid, aff_top, coname, cosite, cologo, codescription, jobbulk, job_id, newstop, affil_id, staff_id, staff_name, newbio, rankage, staffname, staffbio, submitme, Showit, bolFound, a, edit_function, strSQL, dbconn, dbRst, q, id, newstitle, newnews, newpost, addname, newexp, newpos, staff_top, emp_top, jobtitle, jobdescription, news_id, edittopic, editpostedby, editdateposted, editnews, editid, coprofileid, coprofiletitle, coprofiledesc

FUNCTION dbconnect()
	Set dbConn = Server.CreateObject("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\014753\wizDB.mdb") & ";"
	dbConn.Open
	Set dbRst = dbConn.Execute(strSQL)
END FUNCTION


FUNCTION staffnames()
	strSQL = "SELECT * FROM Profiles ORDER BY staff_num ASC"
	dbconnect()
	q = 0
	Do While Not dbRst.EOF
		%> <option value="<%= dbRst("staff_name") %>"><%= dbRst("staff_name") %></option><%
		q = q + 1
		dbRst.MoveNext
	loop
	dbrst.Close
	SET dbrst = NOTHING
END FUNCTION

FUNCTION position(submit)
IF submitme = submit THEN
	response.write "selected"
END IF
END FUNCTION



edit_function = request.querystring("id")




'**********************************************************************************************************************************************************
SELECT CASE edit_function
'**********************************************************************************************************************************************************
CASE "add_news"
					IF Request.Form("submit") = "submit" THEN

						strSQL = "SELECT * FROM publicnews ORDER BY id ASC"
						dbconnect()
							Do While Not dbRst.EOF
								newstop = dbRst("id")
								dbRst.MoveNext
							loop
						dbConn.Close
						SET dbConn = NOTHING
						newstop = newstop + 1
						newstop = int(newstop)

						newstitle = Request.form("newstitle")
						newstitle = Replace(newstitle,"'","")
						newstitle = Lcase(newstitle)
						newpost = Request.Form("postedby")
						newnews = Request.form("subject")
						newnews = Replace(newnews,"'","")

						strSQL = "INSERT INTO publicnews (Topic, PostedBy, dateposted, news, id) VALUES ('" & newstitle & "', '" & newpost & "', '" & date() & "', '" & newnews & "', '"& newstop &"')"
						DBconnect()
						dbConn.Close
						SET dbConn = NOTHING
						Response.redirect ("/dataedit.asp?id=edit_news")
					ELSE
%>
						<form method="post" action="/dataedit.asp?id=add_news">
							<table border="0" width="100%" cellpadding="3" cellspacing="0">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td colspan="2" style="text-align:center;">:: Add Site News ::</td>
								</tr>
								<tr>
									<td width="50%" style="text-align:right;">Posted by: </td>								
									<td width="50%" style="text-align:left;">
										<select name="postedby">
<%
	staffnames()
%>
										</select>
									</td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Title: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="newstitle" size="40" maxlength="50"></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;">News: </td>								
									<td width="50%" style="text-align:left;"><textarea name="subject" cols="30" rows="7" name="textarea" wrap=virtual></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>								
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
					END IF
'**********************************************************************************************************************************************************
CASE "edit_news"
	news_id = request.querystring("topic")
	bolFound = false
	IF (news_id) = "" THEN
		Showit = "all"	
	End if
	If showit = "all" Then
		strSQL = "SELECT * FROM publicnews"
		dbconnect()
%>
			<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<tr>
				<td colspan="2" align="center"><img src="/images/wiz_logo.jpg" alt="Wizardry Fireworks" /></td> 
			</tr>
			<tr>
				<td colspan="2" style="text-align:center;">Click the news title to Edit the topic - Click Delete to remove the topic</td> 
			</tr>
<%
			a = 0
			Do While Not dbRst.EOF
				bolFound = False
%>
			<tr>
				<td>
					<form method="post" action="/dataedit.asp?id=delete_news&topic=<%= dbRst("id") %>">
						<a href="/dataedit.asp?id=edit_news&topic=<% response.write dbRst("id")%>"><% Response.Write dbRst("topic") %></a>
				</td> 
				<td>
						<input name="Delete" type="Submit" value="Delete" />
					</form>
				</td> 
			</tr>
<%
				dbRst.MoveNext
				if (dbRst.EOF) then
					bolfound = true
				end if
			loop
		dbConn.Close
		SET dbConn = NOTHING
%>
			<tr>
				<td colspan="2" align="center"><a href="/dataedit.asp?id=add_news">Add News</a></td> 
			</tr>
		</table>
<%
	ELSE
		IF Request.Form("submit") = "submit" THEN
			newstitle = Request.form("newstopic")
			newstitle = Replace(newstitle,"'","")
			newstitle = Lcase(newstitle)
			newnews = Request.form("newsbulk")
			newnews = Replace(newnews,"'","")
			newpost = Request.Form("postedby")
			strSQL = "UPDATE publicnews SET Topic='"& newstitle &"', PostedBy='"& newpost &"', news='"& newnews &"'  WHERE id="& news_id
			DBconnect()
			dbConn.Close
			SET dbConn = NOTHING
			Response.redirect ("/dataedit.asp?id=edit_news")
		ELSE
%>
		<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<form method="post" action="/dataedit.asp?id=edit_news&topic=<% response.write news_id %>">
<%
			strSQL = "SELECT * FROM publicnews WHERE id = " & news_id &""
			dbconnect()
			Do until (dbRst.EOF or bolfound = "true")
				IF (strComp(dbRst("id"), news_id, vbTextCompare) = 0) THEN
					BolFound = true
					edittopic = dbRst("Topic")
					editpostedby = dbRst("PostedBy")
					editdateposted = dbRst("dateposted")
					editnews = dbRst("news")
					editid = dbRst("id")
				End if
				dbRst.MoveNext
			loop
			dbConn.Close
			SET dbConn = NOTHING
%>
				<tr>
					<td colspan="2" height="40"></td> 
				</tr>
				<tr>
					<td width="200" style="{font-size:16px; color:white; text-align:right;}">Posted by : </td> 
					<td style="{text-align:left;}">
						<select name="postedby">
<%
			staffnames()
%>
						</select>
					</td>

				</tr>
				<tr>
					<td style="{font-size:16px; color:white; text-align:right;}">News Topic : </td> 
					<td style="{text-align:left;}"><input type="text" name="newstopic" size="40" maxlength="50" value="<% response.write edittopic %>"></td>
				</tr>
				<tr>
					<td style="{font-size:16px; color:white; text-align:right;}">News Text : </td> 
					<td style="{text-align:left;}"><textarea name="newsbulk" cols="30" rows="7" name="textarea" wrap=virtual><% response.write editnews %></textarea></td>
				</tr>
				<tr>
					<td colspan="2"><input type="submit" name="Submit" value="submit"></td>
				</tr>
			</form>
		</table>
<%
		END IF
	END IF
'**********************************************************************************************************************************************************
CASE "delete_news"
	id = request.querystring("topic")
	strSQL = "DELETE * FROM publicNews WHERE id = "& id &""
	dbconnect()
	dbConn.Close
	SET dbConn = NOTHING
	Response.redirect ("/dataedit.asp?id=edit_news")

'**********************************************************************************************************************************************************
CASE "add_gallery"
response.write "4"
'**********************************************************************************************************************************************************
CASE "edit_gallery"
response.write "5"
'**********************************************************************************************************************************************************
CASE "delete_gallery"
response.write "6"
'**********************************************************************************************************************************************************
CASE "add_team"
					IF Request.Form("submit") = "submit" THEN
						strSQL = "SELECT * FROM Profiles ORDER BY staff_num ASC"
						dbconnect()
							Do While Not dbRst.EOF
								staff_top = dbRst("staff_num")
								dbRst.MoveNext
							loop
						DBconnect()
						dbConn.Close
						SET dbConn = NOTHING
						staff_top = staff_top + 1
						staff_top = int(staff_top)
						addname = Request.form("newname")
						addname = Replace(addname,"'","")
						newexp = Request.form("newbio")
						newexp = Replace(newexp,"'","")
						newpos = Request.form("rank")
						strSQL = "INSERT INTO profiles (staff_num, staff_name, staff_exp, staff_pos) VALUES ('" & staff_top & "', '" & addname & "', '" & newexp & "', '" & newpos & "')"
						DBconnect()
						dbConn.Close
						SET dbConn = NOTHING
						Response.redirect ("/dataedit.asp?id=edit_team")
					ELSE
%>
						<form method="post" action="/dataedit.asp?id=add_team">
							<table border="0" width="100%" cellpadding="3" cellspacing="0">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td colspan="2" style="text-align:center;">:: Add Team Member ::</td>
								</tr>
								<tr>
									<td width="50%" style="text-align:right;">Name: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="newname" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Position: </td>	
									<td width="50%" style="text-align:left;">
										<select name="rank">
											<option value="Head Pyrotechnician">Head Pyrotechnician</option>
											<option value="Pyrotechnician">Pyrotechnician</option>
											<option value="Trainee Pyrotechnician">Trainee Pyrotechnician</option>
										</select>
									</td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;">Bio: </td>								
									<td width="50%" style="text-align:left;"><textarea name="newbio" cols="30" rows="7" name="textarea" wrap=virtual></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>								
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
					END IF

'**********************************************************************************************************************************************************
CASE "edit_team"
	staff_id = request.querystring("staffid")
	bolFound = false
	IF (staff_id) = "" THEN
		Showit = "all"	
	End if
	If showit = "all" Then
		strSQL = "SELECT * FROM profiles ORDER BY staff_num ASC"
		dbconnect()
%>
			<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<tr>
				<td colspan="2" align="center"><img src="/images/wiz_logo.jpg" alt="Wizardry Fireworks" /></td> 
			</tr>
			<tr>
				<td colspan="2" style="text-align:center;">Click the name to edit profile - Click Delete to remove the profile</td> 
			</tr>
<%
			a = 0
			Do While Not dbRst.EOF
				bolFound = False
%>
			<tr>
				<td>
					<form method="post" action="/dataedit.asp?id=delete_team&staffid=<%= dbRst("staff_num") %>">
						<a href="/dataedit.asp?id=edit_team&staffid=<% response.write dbRst("staff_num")%>"><% Response.Write dbRst("staff_name") %></a>
				</td> 
				<td>
						<input name="Delete" type="Submit" value="Delete" />
					</form>
				</td> 
			</tr>
<%
				dbRst.MoveNext
				if (dbRst.EOF) then
					bolfound = true
				end if
			loop
		dbConn.Close
		SET dbConn = NOTHING
%>
			<tr>
				<td colspan="2" align="center"><a href="/dataedit.asp?id=add_team">Add Staff</a></td> 
			</tr>
		</table>
<%
	ELSE
		IF Request.Form("submit") = "submit" THEN
			staff_name = Request.form("staffname")
			staff_name = Replace(staff_name,"'","")
			newbio = Request.form("staffbulk")
			newbio = Replace(newbio,"'","")
			rankage = Request.Form("rank")
			strSQL = "UPDATE profiles SET staff_name='"& staff_name &"', staff_exp='"& newbio &"', staff_pos='"& rankage &"'  WHERE staff_num="& staff_id
			DBconnect()
			dbConn.Close
			SET dbConn = NOTHING
			Response.redirect ("/dataedit.asp?id=edit_team")
		ELSE
%>
		<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<form method="post" action="/dataedit.asp?id=edit_team&staffid=<%= staff_id %>">
<%


			strSQL = "SELECT * FROM profiles WHERE staff_num = " & staff_id &""
			dbconnect()
			Do until (dbRst.EOF or bolfound = "true")
				IF (strComp(dbRst("staff_num"), staff_id, vbTextCompare) = 0) THEN
					BolFound = true
					staffname = dbRst("staff_name")
					staffbio = dbRst("staff_exp")
					submitme = dbRst("staff_pos")
				End if
				dbRst.MoveNext
			loop
			dbConn.Close
			SET dbConn = NOTHING
%>
				<tr>
					<td colspan="2" height="40"></td> 
				</tr>
				<tr>
					<td width="200" style="{font-size:16px; color:white; text-align:right;}">Name : </td> 
					<td style="{text-align:left;}"><input type="text" name="staffname" size="40" maxlength="50" value="<%= staffname %>"></td>

				</tr>
				<tr>
					<td style="{font-size:16px; color:white; text-align:right;}">Staff Position : </td> 
					<td style="{text-align:left;}">
						<select name="rank">
							<option value="Head Pyrotechnician" "<% position("Head Pyrotechnician") %>">Head Pyrotechnician</option>
							<option value="Pyrotechnician" "<% position("Pyrotechnician") %>">Pyrotechnician</option>
							<option value="Trainee Pyrotechnician" "<% position("Trainee Pyrotechnician") %>">Trainee Pyrotechnician</option>
						</select>
					</td>
				</tr>
				<tr>
					<td style="{font-size:16px; color:white; text-align:right;}">Staff Biography : </td> 
					<td style="{text-align:left;}"><textarea name="staffbulk" cols="30" rows="7" name="textarea" wrap=virtual><%= staffbio %></textarea></td>
				</tr>
				<tr>
					<td colspan="2"><input type="submit" name="Submit" value="submit"></td>
				</tr>
			</form>
		</table>
<%
		END IF
	END IF
'**********************************************************************************************************************************************************
CASE "delete_team"
	id = request.querystring("staffid")
	strSQL = "DELETE * FROM profiles WHERE staff_num = "& id &""
	dbconnect()
	dbConn.Close
	SET dbConn = NOTHING
	Response.redirect ("/dataedit.asp?id=edit_team")
'**********************************************************************************************************************************************************
CASE "add_employ"
					IF Request.Form("submit") = "submit" THEN
						strSQL = "SELECT * FROM employment ORDER BY id ASC"
						dbconnect()
						Do While Not dbRst.EOF
							emp_top = dbRst("id")
							dbRst.MoveNext
						loop
						dbConn.Close
						SET dbConn = NOTHING
						emp_top = emp_top + 1
						emp_top = int(emp_top)
						jobtitle = Request.form("jobtitle")
						jobtitle = Replace(jobtitle,"'","")

						jobdescription = Request.form("jobdescription")
						jobdescription = Replace(jobdescription,"'","")

						strSQL = "INSERT INTO employment (id, JobDescription, title) VALUES ('" & emp_top & "', '" & jobdescription & "', '" & jobtitle & "')"
						DBconnect()
						dbConn.Close
						SET dbConn = NOTHING
						Response.redirect ("/dataedit.asp?id=edit_employ")
					ELSE
%>
						<form method="post" action="/dataedit.asp?id=add_employ">
							<table border="0" width="100%" cellpadding="3" cellspacing="0">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td colspan="2" style="text-align:center;">:: Add Job Available ::</td>
								</tr>
								<tr>
									<td width="50%" style="text-align:right;">Job Title: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="jobtitle" size="25" maxlength="50"></td>
								<tr>

								<tr>
									<td width="50%" valign="top" style="text-align:right;">Job Description: </td>								
									<td width="50%" style="text-align:left;"><textarea name="jobdescription" cols="30" rows="7" name="textarea" wrap=virtual></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>								
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
					END IF
'**********************************************************************************************************************************************************
CASE "edit_employ"
	job_id = request.querystring("jobid")
	bolFound = false
	IF (job_id) = "" THEN
		Showit = "all"	
	End if
	If showit = "all" Then
		strSQL = "SELECT * FROM employment ORDER BY id ASC"
		dbconnect()
%>
			<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<tr>
				<td colspan="2" align="center"><img src="/images/wiz_logo.jpg" alt="Wizardry Fireworks" /></td> 
			</tr>
			<tr>
				<td colspan="2" style="text-align:center;">Click the name to edit job - Click Delete to remove the job<br /><br />(Please ensure that if there is no jobs available, that there remains 1 job description explaining this)</td> 
			</tr>
<%
			a = 0
			Do While Not dbRst.EOF
				bolFound = False
%>
			<tr>
				<td>
					<form method="post" action="/dataedit.asp?id=delete_employ&empid=<%= dbRst("id") %>">
						<a href="/dataedit.asp?id=edit_employ&jobid=<%= dbRst("id")%>"><%= dbRst("title") %></a>
				</td> 
				<td>
						<input name="Delete" type="Submit" value="Delete" />
					</form>
				</td> 
			</tr>
<%
				dbRst.MoveNext
				if (dbRst.EOF) then
					bolfound = true
				end if
			loop
		dbConn.Close
		SET dbConn = NOTHING
%>
			<tr>
				<td colspan="2" align="center"><a href="/dataedit.asp?id=add_employ">Add New Job</a></td> 
			</tr>
		</table>
<%
	ELSE
		IF Request.Form("submit") = "submit" THEN
			jobtitle = Request.form("jobtitle")
			jobtitle = Replace(jobtitle,"'","")
			jobbulk = Request.form("jobbulk")
			jobbulk = Replace(jobbulk,"'","")


			strSQL = "UPDATE employment SET title='"& jobtitle &"', JobDescription='"& jobbulk &"'  WHERE id="& job_id
			DBconnect()
			dbConn.Close
			SET dbConn = NOTHING
			Response.redirect ("/dataedit.asp?id=edit_employ")
		ELSE
%>
		<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<form method="post" action="/dataedit.asp?id=edit_employ&jobid=<%= job_id %>">
<%
			strSQL = "SELECT * FROM employment WHERE id = " & job_id &""
			dbconnect()
			Do until (dbRst.EOF or bolfound = "true")
				IF (strComp(dbRst("id"), job_id, vbTextCompare) = 0) THEN
					BolFound = true
					jobtitle = dbRst("title")
					jobbulk = dbRst("JobDescription")
				End if
				dbRst.MoveNext
			loop
			dbConn.Close
			SET dbConn = NOTHING
%>
				<tr>
					<td colspan="2" align="center"><img src="/images/wiz_logo.jpg" alt="wiz logo" /></td> 
				</tr>
				<tr>
					<td width="200" style="{font-size:16px; color:white; text-align:right;}">Job Title : </td> 
					<td style="{text-align:left;}"><input type="text" name="jobtitle" size="40" maxlength="50" value="<%= jobtitle %>"></td>

				</tr>
				<tr>
					<td valign="top" style="{font-size:16px; color:white; text-align:right;}">Job Description:</td> 
					<td style="{text-align:left;}"><textarea name="jobbulk" cols="30" rows="7" name="textarea" wrap=virtual><%= jobbulk %></textarea></td>
				</tr>
				<tr>
					<td colspan="2" align="center"><input type="submit" name="Submit" value="submit"></td>
				</tr>
			</form>
		</table>
<%
		END IF
	END IF
'**********************************************************************************************************************************************************
CASE "delete_employ"
	id = request.querystring("empid")
	strSQL = "DELETE * FROM employment WHERE id = "& id &""
	dbconnect()
	dbConn.Close
	SET dbConn = NOTHING
	Response.redirect ("/dataedit.asp?id=edit_employ")
'**********************************************************************************************************************************************************
CASE "add_affiliate"
					IF Request.Form("submit") = "submit" THEN
						strSQL = "SELECT * FROM affiliates ORDER BY id ASC"
						dbconnect()
							Do While Not dbRst.EOF
								aff_top = dbRst("id")
								dbRst.MoveNext
							loop
						dbConn.Close
						SET dbConn = NOTHING

						aff_top = aff_top + 1
						aff_top = int(aff_top)
						coname = Request.form("coname")
						coname = Replace(coname,"'","")
						cosite = Request.form("cosite")
						cosite = Replace(cosite,"'","")
						cosite = Lcase(cosite)
						cologo = Request.form("cologo")
						cologo = Replace(cologo,"'","")
						cologo = Lcase(cologo)
						codescription = Request.form("codescription")
						codescription = Replace(codescription,"'","")
						codescription = Lcase(codescription)

						strSQL = "INSERT INTO affiliates (affiliatename, affiliatelink, affiliateimage, id, Description) VALUES ('" & coname & "', '" & cosite & "', '" & cologo & "', '" & aff_top & "', '" & codescription & "')"
						DBconnect()
						dbConn.Close
						SET dbConn = NOTHING
						Response.redirect ("/dataedit.asp?id=edit_affiliate")
					ELSE
%>
						<form method="post" action="/dataedit.asp?id=add_affiliate">
							<table border="0" width="100%" cellpadding="3" cellspacing="0">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td width="50%" style="text-align:right;">Company Name: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="coname" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Site Address: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="cosite" value="http://" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Logo Address: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="cologo" value="http://" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;">Description: </td>								
									<td width="50%" style="text-align:left;"><textarea name="codescription" cols="30" rows="7" name="textarea" wrap=virtual></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
					END IF
'**********************************************************************************************************************************************************
CASE "edit_affiliate"
	affil_id = request.querystring("affilnum")
	bolFound = false
	IF (affil_id) = "" THEN
		Showit = "all"	
	End if
	If showit = "all" Then
		strSQL = "SELECT * FROM affiliates"
		dbconnect()
%>
			<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<tr>
				<td colspan="2" align="center"><img src="/images/wiz_logo.jpg" alt="Wizardry Fireworks" /></td> 
			</tr>
			<tr>
				<td colspan="2" style="text-align:center;">Click the affiliate title to Edit the settings - Click Delete to remove the affiliate</td> 
			</tr>
<%
			a = 0
			Do While Not dbRst.EOF
				bolFound = False
%>
			<tr>
				<td>
					<form method="post" action="/dataedit.asp?id=delete_affiliate&affilnum=<%= dbRst("id") %>">
						<a href="/dataedit.asp?id=edit_affiliate&affilnum=<% response.write dbRst("id")%>"><% Response.Write dbRst("affiliatename") %></a>
				</td> 
				<td>
						<input name="Delete" type="Submit" value="Delete" />
					</form>
				</td> 
			</tr>
<%
				dbRst.MoveNext
				if (dbRst.EOF) then
					bolfound = true
				end if
			loop
		dbConn.Close
		SET dbConn = NOTHING
%>
			<tr>
				<td colspan="2" align="center"><a href="/dataedit.asp?id=add_affiliate">Add News</a></td> 
			</tr>
		</table>
<%
	ELSE
		IF Request.Form("submit") = "submit" THEN
			coname = Request.form("coname")
			coname = Replace(coname,"'","")
			cosite = Request.form("cosite")
			cosite = Replace(cosite,"'","")
			cosite = Lcase(cosite)
			cologo = Request.form("cologo")
			cologo = Replace(cologo,"'","")
			cologo = Lcase(cologo)
			codescription = Request.form("codescription")
			codescription = Replace(codescription,"'","")
			codescription = Lcase(codescription)
			strSQL = "UPDATE affiliates SET affiliatename='"& coname &"', affiliatelink='"& cosite &"', affiliateimage='"& cologo &"', Description='"& codescription &"' WHERE id="& affil_id
			DBconnect()
			dbConn.Close
			SET dbConn = NOTHING
			Response.redirect ("/dataedit.asp?id=edit_affiliate")
		ELSE
			strSQL = "SELECT * FROM affiliates WHERE id = " & affil_id &""
			dbconnect()
				BolFound = true
				affilname = dbRst("affiliatename")
				affillink = dbRst("affiliatelink")
				affilimage = dbRst("affiliateimage")
				affilbulk = dbRst("Description")
				affilid = dbRst("id")
			dbConn.Close
			SET dbConn = NOTHING
%>
						<form method="post" action="/dataedit.asp?id=edit_affiliate&affilnum=<% response.write affil_id %>">
							<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td width="50%" style="text-align:right;">Company Name: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="coname" size="25" value="<%= affilname %>" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Site Address: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="cosite" value="<%= affillink %>" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Logo Address: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="cologo" value="<%= affilimage %>" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;">Description: </td>								
									<td width="50%" style="text-align:left;"><textarea name="codescription" cols="30" rows="7" name="textarea" wrap=virtual><%= affilbulk %></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>								
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
		END IF
	END IF
'**********************************************************************************************************************************************************
CASE "delete_affiliate"
	id = request.querystring("affilnum")
	strSQL = "DELETE * FROM affiliates WHERE id = "& id &""
	dbconnect()
	dbConn.Close
	SET dbConn = NOTHING
	Response.redirect ("/dataedit.asp?id=edit_affiliate")
'**********************************************************************************************************************************************************
CASE "edit_copro"
dim coproid, coprotitlenew, coproidnew
	coproid = request.querystring("coproid")
	bolFound = false
	IF (coproid) = "" THEN
		Showit = "all"	
	End if
	If showit = "all" Then
		strSQL = "SELECT * FROM services"
		dbconnect()
%>
			<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
			<tr>
				<td colspan="2" align="center"><img src="/images/wiz_logo.jpg" alt="Wizardry Fireworks" /></td> 
			</tr>
			<tr>
				<td colspan="2" style="text-align:center;">Click the profile title to Edit the settings <br /> Click Delete to remove the topic</td> 
			</tr>
			<tr>
				<td colspan="2" align="center"><a href="/dataedit.asp?id=edit_coprointro">Edit the Introduction Paragraph</a></td> 
			</tr>
<%
			a = 0
			Do While Not dbRst.EOF
				bolFound = False
				IF dbRst("profileid") = "intro" THEN
					dbRst.MoveNext
					ELSE
%>
			<tr>
				<td>
					<form method="post" action="/dataedit.asp?id=delete_copro&coproid=<%= dbRst("profileid") %>">
						<a href="/dataedit.asp?id=edit_copro&coproid=<% response.write dbRst("profileid")%>"><% Response.Write dbRst("title") %></a>
				</td> 
				<td>
						<input name="Delete" type="Submit" value="Delete" />
					</form>
				</td> 
			</tr>
<%
				dbRst.MoveNext
				END IF
				if (dbRst.EOF) then
					bolfound = true
				end if
			loop
		dbConn.Close
		SET dbConn = NOTHING
%>
			<tr>
				<td colspan="2" align="center"><a href="/dataedit.asp?id=add_copro">Add Company Profile</a></td> 
			</tr>
		</table>
<%
	ELSE
		IF Request.Form("submit") = "submit" THEN
			coprotitlenew = Request.form("coprotitle")
			coprotitlenew = Replace(coprotitlenew,"'","")
			coproidnew = Request.form("coproid")
			coproidnew = Replace(coproidnew,"'","")
			coproidnew = Lcase(coproidnew)
			codescription = Request.form("codescription")
			codescription = Replace(codescription,"'","")
			codescription = Lcase(codescription)
			strSQL = "UPDATE services SET title='"& coprotitlenew &"', description='"& codescription &"', profileid='"& coproidnew &"' WHERE profileid='"& coproid &"'"
			DBconnect()
			dbConn.Close
			SET dbConn = NOTHING
			Response.redirect ("/dataedit.asp?id=edit_copro")
			response.write strsql
		ELSE
			strSQL = "SELECT * FROM services WHERE profileid = '" & coproid &"'"
			dbconnect()
				coprofileid = dbRst("profileid")
				coprofiletitle = dbRst("title")
				coprofiledesc = dbRst("description")
			dbConn.Close
			SET dbConn = NOTHING
%>
						<form method="post" action="/dataedit.asp?id=edit_copro&coproid=<% response.write coprofileid %>">
							<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td width="50%" style="text-align:right;">Profile Title: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="coprotitle" size="25" value="<%= coprofiletitle %>" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Profile id: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="coproid" value="<%= coprofileid %>" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td style="text-align:right;" colspan="2">Please ensure you leave no spaces or formating in the ID.</td>	
								<tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;">Description: </td>								
									<td width="50%" style="text-align:left;"><textarea name="codescription" cols="30" rows="7" name="textarea" wrap=virtual><%= coprofiledesc %></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>								
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
		END IF
	END IF
'**********************************************************************************************************************************************************
CASE "delete_copro"
	id = request.querystring("coproid")
	strSQL = "DELETE * FROM services WHERE profileid = '"& id &"'"
	dbconnect()
	dbConn.Close
	SET dbConn = NOTHING
	Response.redirect ("/dataedit.asp?id=edit_copro")
'**********************************************************************************************************************************************************
CASE "edit_coprointro"
dim profileid
		profileid = "intro"
		IF Request.Form("submit") = "submit" THEN
			codescription = Request.form("codescription")
			codescription = Replace(codescription,"'","")
			codescription = Lcase(codescription)
			strSQL = "UPDATE services SET description='"& codescription &"', profileid='"& profileid &"', title='"& profileid &"' WHERE profileid='intro'"
			response.write strsql
			DBconnect()
			dbConn.Close
			SET dbConn = NOTHING
			Response.redirect ("/dataedit.asp?id=edit_copro")
		ELSE
			strSQL = "SELECT * FROM services WHERE profileid = 'intro'"
			dbconnect()
%>
						<form method="post" action="/dataedit.asp?id=edit_coprointro&coproid=intro">
							<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;">Description: </td>								
									<td width="50%" style="text-align:left;"><textarea name="codescription" cols="30" rows="7" name="textarea" wrap=virtual><%= dbRst("description") %></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>								
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
			dbConn.Close
			SET dbConn = NOTHING
		END IF


'**********************************************************************************************************************************************************
CASE "add_copro"
					IF Request.Form("submit") = "submit" THEN

						coprotitlenew = Request.form("coprotitle")
						coprotitlenew = Replace(coprotitlenew,"'","")
						coproidnew = Request.form("coproid")
						coproidnew = Replace(coproidnew,"'","")
						coproidnew = Lcase(coproidnew)
						codescription = Request.form("codescription")
						codescription = Replace(codescription,"'","")
						codescription = Lcase(codescription)

						strSQL = "INSERT INTO services (profileid, description, title) VALUES ('" & coproidnew & "', '" & codescription & "', '" & coprotitlenew & "')"
						DBconnect()
						dbConn.Close
						SET dbConn = NOTHING
						Response.redirect ("/dataedit.asp?id=edit_copro")
					ELSE
%>
						<form method="post" action="/dataedit.asp?id=add_copro">
							<table border="0" width="100%" cellpadding="10" cellspacing="0" style="border: 1px solid #cccccc;">
								<tr>
									<td colspan="2" style="text-align:center;"><img src="/images/wiz_logo.jpg" alt="wizardry logo" /></td>
								</tr>
								<tr>
									<td width="50%" style="text-align:right;">Profile Title: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="coprotitle" size="25" value="" maxlength="50"></td>
								<tr>
								<tr>
									<td width="50%" style="text-align:right;">Profile id: </td>								
									<td width="50%" style="text-align:left;"><input type="text" name="coproid" value="" size="25" maxlength="50"></td>
								<tr>
								<tr>
									<td style="text-align:right;" colspan="2">Please ensure you leave no spaces or formating in the ID.</td>	
								<tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;">Description: </td>								
									<td width="50%" style="text-align:left;"><textarea name="codescription" cols="30" rows="7" name="textarea" wrap=virtual></textarea></td>
								</tr>
								<tr>
									<td width="50%" valign="top" style="text-align:right;"></td>								
									<td width="50%" style="text-align:center;"><input name="submit" type="submit" value="submit" /></td>
								</tr>
							</table>
						</form>
<%
					END IF
'**********************************************************************************************************************************************************
CASE ELSE
	Response.redirect ("/error.asp")
END SELECT
'**********************************************************************************************************************************************************

END IF
%>
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