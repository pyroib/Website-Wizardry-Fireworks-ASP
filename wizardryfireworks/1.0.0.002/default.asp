<%@ Language=VBScript %>
<% Option Explicit %>
<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"> 
<%
	DIM page, menu_title, JobDescription, EventDate, affiliate_name, show_firing_time, show_start_time, show_finish_time, counting, topic, news, affiliate_link, co_name, link, affiliate_image, affiliateDescription, title, dbconn, dbRst, strSQL, staffid, show_id, image, startat, a, b, display, postid, start, prefix, finish, counter, fs, id, id2, gallery_num, contact_name, enquiry, email, fso, folder, files, file, linetotal, staff_id, staff_pw, new_name, new_bio, permit, bolFound, accesslevel, staff_function, staffpw, client_name, amount_owing, show_date, show_location, show_client, show_type, fire_hour, fire_min, fire_am_pm, start_hour, start_min, start_am_pm, finish_hour, finish_min, finish_am_pm, money_collection, show_price, show_description, staff_needed

	page = Request.Querystring("menu")
	id = request.QueryString("id")
		
	SELECT CASE page
		CASE ""
			menu_title = "Wizardry Fireworks Home Page"
		CASE "news"
			menu_title = "Wizardry News"
		CASE "services"
			menu_title = "Fireworks Services"
		CASE "gallery"
			menu_title = "Fireworks Gallery"
		CASE "profiles"
			menu_title = "Wizardry Fireworks - Company and Staff Profiles"
		CASE "employment"
			menu_title = "Wizardry Employment Opertunities"
		CASE "affiliates"
			menu_title = "Wizardry Affiliates and Links"
		CASE "login"
			menu_title = "Staff only Login - Wizardry Fireworks"
		CASE "contactus"
			menu_title = "Contact Wizardry Fireworks"
		CASE "staff"
			menu_title = "Wizardry Fireworks - Staff Only"
		CASE ELSE
			menu_title = "Home Page"
	END SELECT
%>
	<head>
		<!--please respect all copyright laws. all content on this site belongs to Wizardry Fireworks PTY LTD. //-->
		<!--All scripts, pages and images written and created by Ian Blott. Unless specified//-->
		<!--website created and Maintained by Ian Blott. email me at |ian(a)wizardryfireworks.com|//-->
		<meta name="author" content="Ian Blott" />
		<meta name="generator" content="100% notepad" />
		<meta name="copyright" content="copyright 2003 Wizardry Fireworks PTY LTD" />
		<meta name="publisher" content="Ian Blott" />
		<meta name="description" content="One of Australia's most creative fireworks companies. With impecable safety and huge attention to detail we are rated amoungst the top in the country.">
		<meta name="keywords" content="Fireworks,Sydney Fireworks,Fireworks Sydney, Sydney Pyrotechnics,Fireworks galleries,Australia,Fireworks Australia,Australian Fireworks,family fireworks,consumer fireworks,Fireworks,Fireworks Distributor,Fireworks Wholesale,Fireworks Supplier,Wholesale Fireworks in australia,wholesale fireworks,fireworks distributors,fireworks displays,wedding fireworks,firework packs,corporate firework displays,brilliant fireworks,fireworks for parties,bonfire night,mail order fireworks,low noise firework displays,finale fireworks,professional fireworks,fireworks for sale,firework retailer,discounted fireworks,quality fireworks,pyrotechnics,chinese new year fireworks,halloween fireworks,diwali fireworks,Olympic Fireworks, Olympic Flame, fireworks, pyrotechnics, pyro, New Years Eve, Indoor Pyrotechnics, Australia, Olympics, Sydney 2000, Sydney Harbour Bridge, Events, Special, Darling Harbour, Corporate Events,NYE,celebration,celebrate,party,bang,visual,eye candy,">
		<meta name="robots" content="Follow,Index">
		<meta http-equiv="content-language" content="en" />
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<title><%= menu_title %></title>
<%
		IF page = "" THEN
			RESPONSE.WRITE ("		<link rel='stylesheet' href='openstyle_1024.css' type='text/css' />")
		ELSE
			RESPONSE.WRITE ("		<link rel='stylesheet' href='style_1024.css' type='text/css' />")
		END IF
%>
		<script language="JavaScript" type="text/JavaScript">
			<!--
				function moveover(name,suffix,type){
					eval("document."+name+".src='images/"+suffix + type+ "'")
				}
				
				function moveout(name,suffix,type){
					eval("document."+name+".src='images/"+suffix + type+ "'")
				}
			// --> 
		</script>
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
<%
		FUNCTION dbconnect()
			Set dbConn = Server.CreateObject("ADODB.Connection")
			dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("\portfolio\web\wiz\014753\wizDB.mdb") & ";"
			dbConn.Open
			Set dbRst = dbConn.Execute(strSQL)
		END FUNCTION


		FUNCTION home_page()
%>
			<table border="0" width="600">
				<tr>
					<td align="right">
						<img src="images/guild.gif" alt="Proud Members of the International Pyrotechnicians Guild" />
						<br />
						<a href='/2gb.asp' onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">
							<img src="images/2gb.jpg" border="0" />
						</a>
					</td>
				</tr>
			</table>
<%
		END FUNCTION

		FUNCTION prefix_test()
			IF b < 10 THEN
				prefix = "00"
				ELSE IF b > 9 THEN
					prefix = "0" 
				END IF
			END IF
		END FUNCTION
				
		FUNCTION news_page()
				IF request.QueryString("id") = "" THEN
%>
					<table border="0" width="500" align="center" cellpadding="3" cellspacing="0">
						<tr>
							<td colspan="3">
								<p class="center"><img src="images/logosmall.gif" alt="Wizardry Logo" <br /></p>
								<p class="news">This page Is sorted by Time of posting<br /></p>
								<p class="news">The newest Posts are at the top, Feel free to browse the Archives at your own will<br /></p>
							</td>
						</tr>
<% 
					IF Request.QueryString("startat") = "" THEN
						start = -1
					ELSE
						IF Request.QueryString("startat") < 0 THEN
							start = -1
						ELSE
							start = Request.QueryString("startat")
						END IF
					END IF
					strSQL = "SELECT * FROM publicNews ORDER BY date DESC"
					dbconnect()
%>
						<tr>
							<td width="250" class="news_head"> Topic : </td>
							<td width="125" class="news_head"> Posted by : </td>
							<td width="125" class="news_head"> On : </td>
						</tr>
<%
					IF (start = "all") THEN
						DO WHILE NOT dbRst.EOF
%>
						<tr>
							<td  height="22"></a> </td>
							<td class="news_post"><% Response.Write dbRst("PostedBy") %> </td>
							<td class="news_date"><% Response.Write dbRst("Date") %> </td>
						</tr>
<%
							dbRst.MoveNext
						LOOP 
					ELSE
						a = 0
						b = 0
						display = 10
						
						DO WHILE NOT dbRst.EOF AND (b < display)
							IF (a < start + 1) THEN
								a = a + 1
								dbRSt.MoveNext
							ELSE
%>
						<tr>
							<td  class="news_topic" height="22"><a href="default.asp?menu=news&id=<% Response.Write dbRst("id")%>"><% Response.Write dbRst("Topic")%></a></td>
							<td class="news_post"><% Response.Write dbRst("PostedBy")%> </td>
							<td class="news_date"><% Response.Write dbRst("Date")%> </td>
						</tr>
<%
								a = a + 1
								b = b + 1
								dbRst.MoveNext
							END IF
						LOOP
%>
						<tr>
							<td class="news_nav">
<% 
						IF start > 0 THEN
%>
								<a href="default.asp?menu=news&startat=<%response.write start - 10 %>">View Previous page</a>
<%
						END IF
%>
							</td>
							<td colspan="2" class="news_nav">
<%			
						IF NOT dbRst.EOF THEN
%>
								<a href="default.asp?menu=news&startat=<%response.write start + 10 %>">View Next page</a>
<% 
						ELSE
						 response.write("&nbsp;")
						END IF
%>
							</td>
						</tr>
					</table>
<%
					END IF
					dbrst.Close
					SET dbrst = NOTHING
				ELSE

				strSQL = "SELECT * FROM PublicNews WHERE id = "& Request.QueryString("ID") & ""
				
				dbconnect()
%>
					<table border="0" width="500" align="center" cellpadding="5" cellspacing="0">
						<tr>
							<td colspan="3">
								<SCRIPT LANGUAGE="Javascript">
									<!--  
									function image() { };  
										image = new image(); 
										number = 0;  
						
										image[number++] = "<img src='images/outdoor/outdoor001.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor002.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor004.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor005.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor006.jpg' border='0'>"  
										image[number++] = "<img src='images/outdoor/outdoor007.jpg' border='0'>"  
										image[number++] = "<img src='images/outdoor/outdoor008.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor009.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor010.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor011.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor012.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor013.jpg' border='0'>" 
										image[number++] = "<img src='images/outdoor/outdoor014.jpg' border='0'>"  
						
										increment = Math.floor(Math.random() * number);  document.write(image[increment]);  
									//-->
								</SCRIPT>
<%
			Do While Not dbRst.EOF
%>
							</td>
						</tr>
						<tr> 
							<td colspan="2" align="left" class="news_topic"><div class="bold">Topic:</div><% Response.Write dbRst("Topic")%></td>
							<td width="350" rowspan="2" class="news_content" align="left"><% Response.Write dbRst("Topic") &" -<br />"& dbRst("News")%></td>
						</tr>
						<tr>
							<td width="100" class="news_post" valign="top"><div class="bold">Posted By:</div><% Response.Write dbRst("PostedBy")%></td>
							<td class="news_date" valign="top"><div class="bold">Posted On:</div><% Response.Write dbRst("Date")%></td>
						</tr>
						<tr><td colspan="3" class="news_nav"><a href="javascript:history.back();">Back</a></td></tr>
					</table>
		
<%
			dbRst.MoveNext
			LOOP
			END IF
		END FUNCTION	
				
				
				
		FUNCTION services_page()
%>
					<table border="0" width="550" align="center" colspan="0" cellpadding="0">
						<tr>
							<td colspan="3" class="text">
								<div class="center"><img src="images/logosmall.gif" alt="" /></div>
								The Wizardry team make it possible to have pyrotechnics or special effects  at your event.<br />
								Our services are quite different to what has been offered in the past, They include everything from 
								the magical colours and sounds of  ‘traditional fireworks displays’, to romantic additions to a Bridal 
								watlz<br /><br />
								We are offering now, what we believe will be the future of pyrotechnics. Fireworks that are geared 
								toward being more spectacular and safer than ever before, with a focus on minimal impact to the 
								environment, and the community.<br /><br />
								Wizardry Fireworks Guarentee that on  each display there is at least ONE (1) senior pyrotechnician 
								to ensure audience safety. We look after all approvals, safety equipment, communication equipment, 
								and insurance costs allowing you to plan your big event.<br />
							</td>
						</tr>
						<tr>
							<td width="238" height="27" class="service_1">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=comunitydisplay" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Community Fireworks Displays</a></td>
							<td width="74"></td>
							<td width="238" height="27" class="service_2">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=xmascarol" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Christmas Carol nights</a></td>
						</tr><tr><td height="5"></td></tr><tr>
							<td width="238" height="27" class="service_3">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=customwedding" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Custom designed Wedding effects</a></td>
							<td width="24"></td>
							<td width="238" height="27" class="service_4">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=private" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Private Parties</a></td>
						</tr><tr><td height="5"></td></tr><tr>
							<td width="238" height="27" class="service_5">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=corperate" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Corporate Events</a></td>
							<td width="24"></td>
							<td width="238" height="27" class="service_6">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=nye" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">New Years Eve celebrations</a></td>
						</tr><tr><td height="5"></td></tr><tr>
							<td width="238" height="27" class="service_7">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=beach" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Beach Displays</a></td>
							<td width="20"></td>
							<td width="238" height="27" class="service_1">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=barge" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Barge Displays</a></td>
						</tr><tr><td height="5"></td></tr><tr>
							<td width="238" height="27" class="service_2">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=birthday" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Birthday milestones</a></td>
							<td width="24"></td>
							<td width="238" height="27" class="service_3">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=religion" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Religious festivals</a></td>
						</tr><tr><td height="5"></td></tr><tr>
							<td width="238" height="27" class="service_4">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=grandopening" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Grand Openings</a></td>
							<td width="24"></td>
							<td width="238" height="27" class="service_7">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=crackernight" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Firecracker / Bonfire nights</a></td>
						</tr><tr><td height="5"></td></tr><tr>
							<td width="238" height="27" class="service_6">&nbsp;&nbsp;<a href="breakout.asp?page=services&id=indoor" onClick="NewWindow(this.href,'Wizardry Fireworks','400','400','no','center');return false" onFocus="this.blur()">Indoor Pyrotechnics</a></td>
							<td width="24"></td>
							<td width="238" height="27" class="service_5">&nbsp;&nbsp;And any other celebration.</td>
						</tr>
					</table>
<%
		END FUNCTION
		
		FUNCTION gallery_page()

			IF Request.QueryString("id") = "" THEN
			%>
					<table border="0" width="590">
						<tr>
							<td valign="top" align="center">
								<img src="images/indoor_outdoor_choice.jpg" alt="Please Choose Indoor or OutDoor images" border="0" usemap="#Map" />
								<map name="Map" id="Map">
									<area shape="poly" coords="0,0,0,280,220,280,220,204,260,203,260,0" href="default.asp?menu=gallery&id=indoor" />
									<area shape="rect" coords="223,205,479,477" href="default.asp?menu=gallery&id=outdoor" />
								</map>
							</td>
						</tr>
					</table>
				<%
			ELSE
	
	dim c
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
			
				'Set fso = CreateObject("Scripting.FileSystemObject")
				'Set folder = fso.GetFolder(Server.Mappath("\wiz\images\"& id &"\tn\"))
				'Set files = folder.Files

%>
				<table border="0"  width="590" align="center" cellpadding="0" cellspacing="0">
					<tr>
						<td colspan="3" align="center" class="gal_thumbs">
<% 
						b = startat
						prefix_test()
						response.write "<img src='images/"& id &"/"& id & prefix & b &".jpg' name='big_one' alt='Image' />"
%>
						</td>
					</tr>
					<tr>
						<td width="40" class="text">
							<% IF startat > 2 THEN %>
								<a href="default.asp?menu=gallery&id=<%= id %>&startat=<% response.write startat - 10%>"><img src="images/last.gif" alt="Previous Page" border="0" /></a>
							<% END IF %>
						</td>
						<td align="center">
<%
		
				DO UNTIL linetotal = 5
					prefix_test()		
%>
					<a href="javascript:moveover('big_one','<%= id %>/<%= id & prefix & b %>','.jpg')"><img src="images/<%= id %>/tn/<%= id & prefix & b %>.jpg" alt="" border="0" /></a>
<%
					linetotal = linetotal + 1
					b = b + 1
				LOOP
				
				Response.Write "<br />"
				
				DO UNTIL linetotal = 10
					prefix_test()
%>
					<a href="javascript:moveover('big_one','<%= id %>/<%= id & prefix & b %>','.jpg')"><img src="images/<%= id %>/tn/<%= id & prefix & b %>.jpg" alt="" border="0" /></a>
<%
					linetotal = linetotal + 1
					b = b + 1
				LOOP
				
				startat = startat + linetotal
%>
						</td>
						<td width="40">
						<% if b < c THEN %>
								<a href="default.asp?menu=gallery&id=<%= id %>&startat=<%= startat %>"><img src="images/next.gif" alt="Next Page" border="0" /></a>
						<% 
						response.write startat
						end if%>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="3">
							<a href="default.asp?menu=gallery&id=<%= id2 %>">Don't forget to take a look at our <%= id2 %> display images!</a>
						</td>
					</tr>
				</table>
<%
				'Set folder = Nothing
				'Set files = Nothing
				'Set fso = Nothing
			END IF
		END FUNCTION
		
		FUNCTION profiles_page()
%>
					<table border="0" width="500" align="center" cellpadding="3" cellspacing="0">
						<tr>
							<td colspan="4">
								<p class="center"><img src="images/logosmall.gif" alt="Wizardry Logo" <br /></p>
								<p class="news">Wizardry Fireworks Pty Ltd is a team of trained licensed pyrotechnicians performing fireworks effects that are truly different.</p>
								<p class="news">A team that bring in a multitude of skills in many industries and many years of experience.</p>
								<p class="news">With this skill set working together all clients can expect the safest displays and the most unique.</p>
								<p class="news">We our proud of what we have achieved in developing a specialised team and a loyal clientele to date and have the biggest pyrotechnic accolades in our sights.</p>
								<p class="news">We look forward in making your next event even more special and challenge ourselves to bring out our best.</p>
							</td>
						</tr>
						<tr>
							<td class="news_head" width="120">Name: </td>
							<td class="news_head" width="130">Position & profile: </td>
							<td class="news_head" width="120">Name: </td>
							<td class="news_head" width="130">Position & profile: </td>
						</tr>
<%
			strSQL = "SELECT * FROM Profiles WHERE status < 2 order BY staff_num ASC"
			dbconnect()
			DO WHILE NOT (dbRst.EOF)
				IF dbRst.EOF THEN
					return false
				ELSE
%>
						<tr>
							<td class="profile_name" width="120"><%= dbRst("Name") %></td>
							<td class="profile_years" width="130">
<%
%>
							<A HREF="javascript:NewWindow('breakout.asp?page=profiles&id=<%= dbRst("staff_id") %>')"><%= dbRst("position") %></a></td>
<%
						dbRst.MoveNext
				END IF
				IF dbRst.EOF THEN
%>
							<td class="profile_name" width="120"></td>
							<td class="profile_years" width="130"></td>
						</tr>
<%
					ELSE
%>
							<td class="profile_name" width="120"><%= dbRst("Name") %></td>
							<td class="profile_years" width="130"><A HREF="javascript:NewWindow('breakout.asp?page=profiles&id=<%= dbRst("staff_id") %>')"><%= dbRst("position") %></a></td>
						</tr>
<%
				dbRst.MoveNext
				END IF
			LOOP
%>
					</table>
<%
		END FUNCTION
		
		
		
		FUNCTION employment_page()

		strSQL = "SELECT * FROM employment"

		dbconnect()
%>
		<table border="0" width="500" align="center" cellpadding="3" cellspacing="0">
			<tr><td height="30" colspan="2" align="center"></td></tr>
			<tr><td colspan="2" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks" /></td></tr>
			<tr><td height="20" colspan="2" align="center"></td></tr>
			<tr>
				<td class="emp_top">Positions Available</td>
				<td class="emp_top">Duties Include :</td>
			</tr>
			<tr>
				<td class="emp_title" height="15"></td>
				<td class="emp_detail" height="15"></td>
			</tr>
<%
		DO WHILE NOT (dbRst.EOF)
			IF NOT (dbRst.EOF) THEN
%>
			<tr>
				<td class="emp_title"><%= dbRst("title") %></td>
				<td class="emp_detail"><%= dbRst("JobDescription") %></td>
			</tr>
			<tr>
				<td class="emp_title" height="15"></td>
				<td class="emp_detail" height="15"></td>
			</tr>
<%
			dbRst.MoveNext
			END IF
		LOOP
%>
			<tr>
				<td class="emp_thanks" colspan="2">We thank-you for your interest in Working at Wizardry Fireworks.</td>
			</tr>
		</table>
<%

%>

<%
		END FUNCTION
		
		
		
		FUNCTION affiliates_page()
%>
		<table border="0" width="500" cellpadding="3" cellspacing="0" align="center">
			<tr><td colspan="2" height="20" align="center"></td></tr>
			<tr><td colspan="2" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks Logo" /></td></tr>
			<tr><td colspan="2" height="20" align="center"></td></tr>
			<tr><td colspan="2" height="10" align="center" class="affil_border"></td></tr>
			<tr>
				
<%
	strSQL = "SELECT * FROM affiliates"

	dbconnect()

	Do While Not (dbRst.EOF)

		Response.Write "<td class='affil_logo' align='center'>"
		Response.Write "<a href='"& dbRst("affiliatelink") &"' target='_blank'><img border='0' src='"& dbRst("affiliateimage") &"' alt='Link to "& dbRst("affiliatename") &"'></a>"
%>
				</td>
				<td class="affil_desc" valign="top">
					<a href="<%= dbRst("affiliatelink") %>" target="_blank"><%= dbRst("affiliatename") %></a><br /><br />
<%
		Response.Write dbRst("description")
		dbRst.MoveNext
%>
</td>
</tr>

<tr><td colspan="2" height="5" align="center" class="affil_border"></td></tr>
<%
	loop
%>		
				
			
			<tr><td colspan="2" height="10" align="center" class="affil_border"></td></tr>
		</table>

<%
		END FUNCTION
		




		FUNCTION login_page()
%>
		<table border="0" width="500" cellpadding="3" cellspacing="0" align="center">
			<tr><td colspan="2" height="10" align="center"></td></tr>
			<tr><td colspan="2" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks Logo" /></td></tr>
			<tr><td colspan="2" height="10" align="center"></td></tr>
			<tr><td colspan="2" height="20" align="center" class="staffonly">Wizardry Fireworks - Staff Access</td></tr>
			<tr>
				<td class="affil_logo">
<%

	IF request.querystring("logout") = "yes" THEN
		session.abandon
		Response.redirect ("default.asp?menu=login")	
	END IF
	
	IF Session("staff_id") = "" THEN
	
		IF Request.Form("Submit") = "Login" THEN
	
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
				Response.redirect ("default.asp?menu=login")
			ELSE
				Session("staff_id") = staff_id
				Session("permit") = dbRst("permit")
				dbConn.Close
				SET dbConn = NOTHING
				Response.redirect ("default.asp?menu=login")
			END IF
			
		ELSE
%>
					<table width="550" border="0" cellpadding="0" cellspacing="0">
						<script language="Javascript">
							<!--
							alert ("For this Demo user name = 'test' password = 'test' ")
							//-->
						</script>
						<form method="post" action="default.asp?menu=login">
							<tr> 
								<td width="250" valign="middle" height="50" align="right" class="login">Staff ID: &nbsp;&nbsp;</td>
								<td width="250" valign="middle" height="50" align="left"><input type="text" name="txtstaff_Id" size="30" maxlength="25" autocomplete="no" />
							</tr>
							<tr> 
								<td width="250" valign="middle" height="50" align="right" class="login">Password: &nbsp;&nbsp;</td>
								<td width="250" valign="middle" height="50" align="left"><input type="password" name="txtstaff_pw" size="30" maxlength="25" /></td>
							</tr>
							<tr> 
								<td width="250" valign="middle" height="50" align="left"></td>
								<td width="250" valign="middle" height="50" align="left"><input type="submit" name="Submit" value="Login" /></td>
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
					<table width="550" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td valign="top">
								<table width="125" border="1" cellpadding="0" cellspacing="0">
									<tr><td class="staffbrowse">Personal Data</td></tr>
									<!--<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=personal_details">Edit My Details</a></td></tr>//-->
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=view_all_shows">View Upcoming Shows</a></td></tr>
									<tr><td height="15"></td></tr>
<%
		IF Session("permit") < 2 THEN
%>
									<tr><td class="staffbrowse">Staff Data</td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=add_login">Add Login User</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_login">Edit Login User</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=delete_login">Delete Login User</a></td></tr>
<% 
		END IF
%>
								</table>
								<table><tr><td width="250" valign="middle" height="50" align="left"></td></tr></table>
								<table><tr><td width="250" valign="middle" height="50" align="left"><a href="default.asp?menu=login&logout=yes"> Log Out</a></td></tr></table>
							</td>
							<td valign="top" class="staffbrowse">
								<table width="300" border="1" cellpadding="0" cellspacing="0">
<%
		IF Session("permit") < 2 THEN
%>
									<tr><td class="staffbrowse">Staff only Content</td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=add_show">Add show Information</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_a_show">Edit show Information</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_a_show">Delete show Information</a></td></tr>
									<tr><td height="15"></td></tr>
<%
		END IF
%>
									<tr><td class="staffbrowse">Staff News</td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=view_s_news">View Staff News</a></td></tr>									
								</table>
							</td>
							<td valign="top" class="staffbrowse">
								<table width="125" border="1" cellpadding="0" cellspacing="0">
<%
		IF Session("permit") < 2 THEN
%>
									<tr><td class="staffbrowse">Site Details</td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=add_jobs">Add Employment</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_jobs">Edit Employment</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_jobs">Delete Employment</a></td></tr>
									<tr><td height="15"></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=add_affiliate">Add Affiliate</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_affiliates">Edit Affiliate</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_affiliates">Delete Affiliate</a></td></tr>
									<tr><td height="15"></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=add_p_news">Add Public News</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_p_news">Edit Public News</a></td></tr>
									<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_p_news">Delete Public News</a></td></tr>
									<tr><td height="15"></td></tr>
									<!--<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=add_s_news">Add Staff News</a></td></tr>//-->
									<!--<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_s_news">Edit Staff News</a></td></tr>//-->
									<!--<tr><td class="staffbrowse"><a href="default.asp?menu=staff&data=edit_s_news">Delete Staff News</a></td></tr>//-->
<%
		END IF
%>
								</table>
							</td>
						</tr>
					</table>
<%
	END IF


'END OF PAGE 

		END FUNCTION

		
		
		FUNCTION Staff_only()
			staff_function = request.querystring("data")
			IF Session("staff_id") = "" THEN
				Response.redirect ("default.asp?menu=login")	
			END IF
			SELECT CASE staff_function

				CASE "view_all_shows"
					id = request.querystring("id")
						bolFound = FALSE
						strSQL = "SELECT * FROM Showdetails ORDER BY show_date DESC"
						dbconnect()
%>
						<table border="0" width="500" align="center" cellpadding="3" cellspacing="0">
							<tr><td height="30" colspan="3" align="center"></td></tr>
							<tr><td colspan="3" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks" /></td></tr>
							<tr><td height="20" colspan="3" align="center"></td></tr>
							<tr>
								<td class="emp_top">Client Name:</td>
								<td class="emp_top">Show Date:</td>
								<td class="emp_top">Show Location:</td>
							</tr>
<%
								a = 0
								DO WHILE NOT (dbRst.EOF OR a = 15)				
%>
									<tr>
										<td class="news_head"><a href="default.asp?menu=staff&data=view_a_show&id=<%= dbRst("id") %>"><%= dbRst("show_client") %></a></td>
										<td class="emp_title"><%= dbRst("show_date") %></td>
										<td class="news_head"><%= dbRst("show_Location") %></td> 
									</tr>
<%
									a= a + 1
									dbRst.MoveNext
									IF (dbRst.EOF) THEN
										bolfound = TRUE
									END IF
								LOOP
						dbConn.Close
						SET dbConn = NOTHING
%>
							<tr>
								<td class="emp_thanks" colspan="4" height="20"><a href="default.asp?menu=login">Staff only Page</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:history.back();">Back</a></td>
							</tr>
						</table>
<%
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************

				CASE "add_show"
%>
						<form method="post" action="default.asp?menu=staff&data=add_show">
							<table border="0" width="500" align="center" cellpadding="3" cellspacing="0">
								<tr><td height="10" colspan="2" align="center"></td></tr>
								<tr><td colspan="2" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks" /></td></tr>
								<tr><td height="20" colspan="2" align="center"></td></tr>
								<tr>
									<td class="emp_top" colspan="2">:: Add Show Details ::</td>
								</tr>
								<tr>
									<td class="emp_title">Date:</td>
									<td class="emp_detail">
										<select name="show_day">
											<option value="1" >1</option>
											<option value="2" >2</option>
											<option value="3" >3</option>
											<option value="4" >4</option>
											<option value="5" >5</option>
											<option value="6" >6</option>
											<option value="7" >7</option>
											<option value="8" >8</option>
											<option value="9" >9</option>
											<option value="10" >10</option>
											<option value="11" >11</option>
											<option value="12" >12</option>
											<option value="13" >13</option>
											<option value="14" >14</option>
											<option value="15" >15</option>
											<option value="16" >16</option>
											<option value="17" >17</option>
											<option value="18" >18</option>
											<option value="19" >19</option>
											<option value="20" >20</option>
											<option value="21" >21</option>
											<option value="22" >22</option>
											<option value="23" >23</option>
											<option value="24" >24</option>
											<option value="25" >25</option>
											<option value="26" >26</option>
											<option value="27" >27</option>
											<option value="28" >28</option>
											<option value="29" >29</option>
											<option value="30" >30</option>
											<option value="31" >31</option>
										</select>
										<select name="show_month">
											<option value="01" >January</option>
											<option value="02" >Feburary</option>
											<option value="03" >March</option>
											<option value="04" >April</option>
											<option value="05" >May</option>
											<option value="06" >June</option>
											<option value="07" >July</option>
											<option value="08" >August</option>
											<option value="09" >September</option>
											<option value="10" >October</option>
											<option value="11" >November</option>
											<option value="12" >December</option>
										</select>
										<select name="show_year">
											<option value="2003" >2003</option>
											<option value="2004" >2004</option>
											<option value="2005" selected>2005</option>
											<option value="2006" >2006</option>
											<option value="2007" >2007</option>
											<option value="2008" >2008</option>
										</select>
									</td>
								</tr>
								<tr>
									<td class="emp_title">Location:</td>								
									<td class="emp_detail"><input type="text" name="show_Location" size="30" maxlength="50"></td>
								<tr>
								<tr>
									<td class="emp_title">Clients Name:</td>								
									<td class="emp_detail"><input type="text" name="client_name" size="30" maxlength="50"></td>
								<tr> 
								<tr>
									<td class="emp_title">Type:</td>								
									<td class="emp_detail">
										<select name="show type">
											<option value="indoor" >Indoor</option>
											<option value="outdoor" >Outdoor</option>
										</select>
									</td>
								<tr> 
								<tr>
									<td class="emp_title">Firing Time :</td>								
									<td class="emp_detail">
										<select name="fire_hour">
											<option value="1" >1</option>
											<option value="2" >2</option>
											<option value="3" >3</option>
											<option value="4" >4</option>
											<option value="5" >5</option>
											<option value="6" >6</option>
											<option value="7" >7</option>
											<option value="8" >8</option>
											<option value="9" >9</option>
											<option value="10" >10</option>
											<option value="11" >11</option>
											<option value="12" >12</option>
										</select>
										<select name="fire_min">
											<option value="00" >00</option>
											<option value="05" >05</option>
											<option value="10" >10</option>
											<option value="15" >15</option>
											<option value="20" >20</option>
											<option value="25" >25</option>
											<option value="30" >30</option>
											<option value="35" >35</option>
											<option value="40" >40</option>
											<option value="45" >45</option>
											<option value="50" >50</option>
											<option value="55" >55</option>
										</select>
										<select name="fire_am_pm">
											<option value="AM" >AM</option>
											<option value="PM" >PM</option>
										</select>
									</td>
								<tr> 
								<tr>
									<td class="emp_title">Work start :</td>								
									<td class="emp_detail">
										<select name="start_hour">
											<option value="1" >1</option>
											<option value="2" >2</option>
											<option value="3" >3</option>
											<option value="4" >4</option>
											<option value="5" >5</option>
											<option value="6" >6</option>
											<option value="7" >7</option>
											<option value="8" >8</option>
											<option value="9" >9</option>
											<option value="10" >10</option>
											<option value="11" >11</option>
											<option value="12" >12</option>
										</select>
										<select name="start_min">
											<option value="00" >00</option>
											<option value="05" >05</option>
											<option value="10" >10</option>
											<option value="15" >15</option>
											<option value="20" >20</option>
											<option value="25" >25</option>
											<option value="30" >30</option>
											<option value="35" >35</option>
											<option value="40" >40</option>
											<option value="45" >45</option>
											<option value="50" >50</option>
											<option value="55" >55</option>
										</select>
										<select name="start_am_pm">
											<option value="AM" >AM</option>
											<option value="PM" >PM</option>
										</select>								
									</td>
								<tr> 
								<tr>
									<td class="emp_title">Work Finish :</td>								
									<td class="emp_detail">
										<select name="finish_hour">
											<option value="1" >1</option>
											<option value="2" >2</option>
											<option value="3" >3</option>
											<option value="4" >4</option>
											<option value="5" >5</option>
											<option value="6" >6</option>
											<option value="7" >7</option>
											<option value="8" >8</option>
											<option value="9" >9</option>
											<option value="10" >10</option>
											<option value="11" >11</option>
											<option value="12" >12</option>
										</select>
										<select name="finish_min">
											<option value="00" >00</option>
											<option value="05" >05</option>
											<option value="10" >10</option>
											<option value="15" >15</option>
											<option value="20" >20</option>
											<option value="25" >25</option>
											<option value="30" >30</option>
											<option value="35" >35</option>
											<option value="40" >40</option>
											<option value="45" >45</option>
											<option value="50" >50</option>
											<option value="55" >55</option>
										</select>
										<select name="finish_am_pm">
											<option value="AM" >AM</option>
											<option value="PM" >PM</option>
										</select>
									</td>
								<tr> 
								<tr>
									<td class="emp_title">Collect Money :</td>								
									<td class="emp_detail">
										<select name="amount_owing">
											<option value="yes" >Yes</option>
											<option value="no" >No</option>
										</select>
									</td>
								<tr> 
								<tr>
									<td class="emp_title">Price :</td>								
									<td class="emp_detail"><input type="text" name="show_Price" size="30" maxlength="25"></td>
								<tr> 
								<tr>
									<td class="emp_title">Description :</td>								
									<td class="emp_detail"><input type="text" name="show_description" size="30" maxlength="255"></td>
								<tr>
								<tr>
									<td class="emp_title">Staff Needed :</td>								
									<td class="emp_detail">
										<select name="staff_needed">
											<option value="1" >1</option>
											<option value="2" >2</option>
											<option value="3" >3</option>
											<option value="4" >4</option>
											<option value="5" >5</option>
											<option value="6" >6</option>
										</select>
									</td>
								</tr>
								<tr> 
									<td colspan="2" class="emp_thanks" valign="middle" align="center" height="20"><input type="Submit" name="Submit" value="Submit"></td>
								</tr>
								<tr> 
									<td colspan="2" class="emp_thanks" valign="middle" align="center" height="20"><a href="javascript:history.back();">Back</a></td>
								</tr>
							</table>
						</form>
<%


'**************************************************************************************************************************************			
'**************************************************************************************************************************************

'**************************************************************************************************************************************
				CASE "delete_show"
					IF request.querystring("data") = "delete_show" THEN
						id = request.querystring("ID")
					END IF
'**************************************************************************************************************************************	
'**************************************************************************************************************************************

'**************************************************************************************************************************************
			


				CASE "edit_a_show"
					id = request.querystring("id")
						bolFound = FALSE
						strSQL = "SELECT * FROM Showdetails ORDER BY show_date DESC"
						dbconnect()
%>
						<table border="0" width="500" align="center" cellpadding="3" cellspacing="0">
							<tr><td height="30" colspan="2" align="center"></td></tr>
							<tr><td colspan="2" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks" /></td></tr>
							<tr><td height="20" colspan="2" align="center"></td></tr>
							<tr>
								<td class="emp_top">Client Name</td>
								<td colspan="3" class="emp_top">Show Date:</td>
							</tr>
<%
								a = 0
								DO WHILE NOT (dbRst.EOF OR a = 15)				
%>
									<tr>
										<td class="emp_detail"><%= dbRst("show_client") %></td>
										<td class="emp_title"><%= dbRst("show_date") %></td>
										<td class="news_head"><a href="default.asp?menu=staff&data=edit_a_show">Edit</a></td> 
										<td class="news_head"><a href="default.asp?menu=staff&data=edit_a_show">Delete</a></td> 
									</tr>
<%
									a= a + 1
									dbRst.MoveNext
									IF (dbRst.EOF) THEN
										bolfound = TRUE
									END IF
								LOOP
						dbConn.Close
						SET dbConn = NOTHING
%>
							<tr>
								<td class="emp_thanks" colspan="4" height="20"><a href="default.asp?menu=login">Staff only Page</a></td>
							</tr>
						</table>
<%
'**************************************************************************************************************************************

'**************************************************************************************************************************************




			CASE ELSE
				Response.redirect ("default.asp?menu=login")
			END SELECT
		END FUNCTION
'**************************************************************************************************************************************

'**************************************************************************************************************************************

	
		FUNCTION contactus_page()
		
			IF id = "form" THEN 
			
%>
				<form method="post" name="signup" action="#">
					<table  width="580" border="0" cellpadding="0" cellspacing="0" align="center">
						<tr><td colspan="2" height="20"></td></tr>
						<tr><td colspan="2" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks Logo" /></td></tr>
						<tr><td colspan="2" height="20"></td></tr>
						<tr><td class="contactus_head" colspan="2">Contact Wizardry Fireworks</td></tr>
						<tr><td colspan="2" height="20" class="contactus_email"></td></tr>
						<tr>
							<td width="144" align="right" class="contactus_head">Name : </td>
							<td width="243" align="left" class="contactus_email">
								<input name="contact_name" type="text" size="20" maxlength="10" />
							</td>
						</tr>
						<tr><td height="12" colspan="2" class="contactus_email"></td></tr>
						<tr>
							<td align="right" class="contactus_head">* Email Address : </td>
							<td align="left" class="contactus_email">
								<input name="email" type="text" size="20" maxlength="50" />
							</td>
						</tr>
						<tr><td height="12" colspan="2" class="contactus_email"></td></tr>
						<tr>
							<td align="right" class="contactus_head">* Enquiry :</td>
							<td align="left" class="contactus_email">
								<textarea name="enquiry" cols="50" rows="10" name="textarea"></textarea>
							</td>
						</tr>
						<tr><td colspan="2" class="contactus_email">&nbsp;</td><tr>
						<tr>
							<td class="contactus_head" colspan="2">
								<input name="submit" type="submit" value="submit" />
							</td>
						</tr>
						<tr>
							<td colspan="4"></td>
						</tr>
					</table>
				</form>
<%
			ELSE
%>
		<table border="0" width="500" cellpadding="3" cellspacing="0" align="center">
			<tr><td height="20" colspan="2" align="center"></td></tr>
			<tr><td colspan="2" align="center"><img src="images/logosmall.gif" alt="Wizardry Fireworks" /></td></tr>
			<tr><td height="10" colspan="2" align="center"></td></tr>
			<tr><td colspan="2" align="center" class="contactus_head"> E-mail</td></tr>
			<tr>
				<td align="center" class="contactus_email">
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
				<td align="center" class="contactus_name">General Enquiries</td>
			</tr>
			<tr>
				<td align="center" class="contactus_email">
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
				<td align="center" class="contactus_name">Site Enquiries</td>
			</tr>
			<tr>
				<td align="center" class="contactus_email">
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
				<td align="center" class="contactus_name">Greg Coorey - Head Pyrotechnician</td>
			</tr>
			<tr>
				<td align="center" class="contactus_email">
				<script type="text/javascript">
				<!--
					var crawlerbreak= new Array("ian","@","wizardryfireworks",".com")
					var mymail = ''
					for (i=0;i<crawlerbreak.length;i++)
					mymail = mymail + crawlerbreak[i]
						document.write('<a href="mailto:'+mymail+'">'+mymail+'</a>')
				//-->
				</script>
				</td>
				<td align="center" class="contactus_name">Ian Blott - Web Designer</td>
			</tr>
			<tr><td height="20" colspan="2" align="center"></td></tr>
			<tr>
				<td align="center" class="contactus_head">Mail Address</td>
				<td align="center" class="contactus_head">Phone</td>
			</tr>
			<tr>
				<td align="center" class="contactus_name">
					PO BOX 95<br />Baulkham Hills 1755<br />N.S.W. Sydney Australia<br />
				</td>
				<td align="center" class="contactus_name">
					Office Phone:<br />+61 (02) 9686 1999<br />
					<br />
					Office Fax:<br />+61 (02) 9686 9191<br />
				</td>
			</tr>
			<tr><td height="20" colspan="2" align="center"></td></tr>
			<tr><td colspan="2" align="center" class="contactus_head"> Online</td></tr>
			<tr><td colspan="2" align="center" class="contactus_email"><a href="default.asp?menu=contactus&id=form">Click Here</a></td></tr>

		</table>
<%
			END IF
			
		END FUNCTION








		FUNCTION profiles_page()
%>
					<table border="0" width="500" align="center" cellpadding="3" cellspacing="0">
						<tr>
							<td colspan="4">
								<p class="center"><img src="images/logosmall.gif" alt="Wizardry Logo" <br /></p>
								<p class="news">Wizardry Fireworks Pty Ltd is a team of trained licensed pyrotechnicians performing fireworks effects that are truly different.</p>
								<p class="news">A team that bring in a multitude of skills in many industries and many years of experience.</p>
								<p class="news">With this skill set working together all clients can expect the safest displays and the most unique.</p>
								<p class="news">We our proud of what we have achieved in developing a specialised team and a loyal clientele to date and have the biggest pyrotechnic accolades in our sights.</p>
								<p class="news">We look forward in making your next event even more special and challenge ourselves to bring out our best.</p>
							</td>
						</tr>
						<tr>
							<td class="news_head" width="120">Name: </td>
							<td class="news_head" width="130">Position & profile: </td>
							<td class="news_head" width="120">Name: </td>
							<td class="news_head" width="130">Position & profile: </td>
						</tr>
<%
			strSQL = "SELECT * FROM Profiles WHERE status < 2 order BY staff_num ASC"
			dbconnect()
			DO WHILE NOT (dbRst.EOF)
				IF dbRst.EOF THEN
					return false
				ELSE
%>
						<tr>
							<td class="profile_name" width="120"><%= dbRst("Name") %></td>
							<td class="profile_years" width="130">
<%
%>
							<A HREF="javascript:NewWindow('breakout.asp?page=profiles&id=<%= dbRst("staff_id") %>')"><%= dbRst("position") %></a></td>
<%
						dbRst.MoveNext
				END IF
				IF dbRst.EOF THEN
%>
							<td class="profile_name" width="120"></td>
							<td class="profile_years" width="130"></td>
						</tr>
<%
					ELSE
%>
							<td class="profile_name" width="120"><%= dbRst("Name") %></td>
							<td class="profile_years" width="130"><A HREF="javascript:NewWindow('breakout.asp?page=profiles&id=<%= dbRst("staff_id") %>')"><%= dbRst("position") %></a></td>
						</tr>
<%
				dbRst.MoveNext
				END IF
			LOOP
%>
					</table>
<%
		END FUNCTION			
			








		FUNCTION error_page()
%>
			<table width="590" height="100%">
				<tr>
					<td>
						<div class="news_nav">
							<br /><br /><br />
							<a href="default.asp">Sorry, You have accesed a page that does not exist on this server.</a>
							<br /><br /><br /><br />
						</div>
					</td>
				</tr>
			</table>			
<%
		END FUNCTION
%>
	<body bgcolor="black">
		<table border="0" cellpadding="0" cellspacing="0">
			<tr> 
				<td width="250" valign="top" class="center">
					<img src="images/logosmall.gif" alt="Wizardry Fireworks Logo" /><br /><br />		
					<a href="default.asp?menu="><img src="images/home_1.jpg" alt="Home" name="home" border="0" id="home" onMouseOver="moveover('home','home_2','.jpg')" onMouseOut="moveout('home','home_1','.jpg')" /></a><br />
					<a href="default.asp?menu=news"><img src="images/news_1.jpg" alt="News" name="news" border="0" id="news" onMouseOver="moveover('news','news_2','.jpg')" onMouseOut="moveout('news','news_1','.jpg')" /></a><br />
					<a href="default.asp?menu=services"><img src="images/services_1.jpg" alt="Services" name="services" border="0" id="services" onMouseOver="moveover('services','services_2','.jpg')" onMouseOut="moveout('services','services_1','.jpg')" /></a><br />
					<a href="default.asp?menu=gallery"><img src="images/Wiz_magic_1.jpg" alt="Gallery" name="gallery" border="0" id="gallery" onMouseOver="moveover('gallery','wiz_magic_2','.jpg')" onMouseOut="moveout('gallery','wiz_magic_1','.jpg')" /></a><br />
					<a href="default.asp?menu=profiles"><img src="images/profiles_1.jpg" alt="Profiles" name="profiles" border="0" id="profiles" onMouseOver="moveover('profiles','profiles_2','.jpg')" onMouseOut="moveout('profiles','profiles_1','.jpg')" /></a><br />
					<a href="default.asp?menu=employment"><img src="images/employment_1.jpg" alt="Employment" name="employment" border="0" id="employment" onMouseOver="moveover('employment','employment_2','.jpg')" onMouseOut="moveout('employment','employment_1','.jpg')" /></a><br />
					<a href="default.asp?menu=affiliates"><img src="images/affiliates_1.jpg" alt="Affiliates" name="affiliates" border="0" id="affiliates" onMouseOver="moveover('affiliates','affiliates_2','.jpg')" onMouseOut="moveout('affiliates','affiliates_1','.jpg')" /></a><br />
					<a href="default.asp?menu=login"><img src="images/staff_only_1.jpg" alt="Staff Only" name="staff_only" border="0" id="staff_only" onMouseOver="moveover('staff_only','staff_only_2','.jpg')" onMouseOut="moveout('staff_only','staff_only_1','.jpg')" /></a><br />
					<a href="default.asp?menu=contactus"><img src="images/contact_us_1.jpg" alt="Contact Us" name="contact_us" border="0" id="contact_us" onMouseOver="moveover('contact_us','contact_us_2','.jpg')" onMouseOut="moveout('contact_us','contact_us_1','.jpg')" /></a><br />
				</td>
				<td width="610" height="540" align="center">
					<table border="0" cellpadding="0" cellspacing="0" height="100%">
						<tr>
<%
	IF page <> "" THEN
%>
							<td background="images/pageheader.jpg" height="16" width="600"></td>
						</tr>
						<tr>
							<td  valign="top">
								<table border="0" cellpadding="0" cellspacing="0" class="tableborder">
									<tr>
										<td>
											<table border="0" width="598" height="520" cellpadding="0" cellspacing="0" class="largetable">
												<tr>
													<td valign="top" class="main_table">
	<%	ELSE %>
                     		 <td></td>
						</tr>
							<td align="right" valign="bottom">
<%
	END IF
					
					
	SELECT CASE page
		CASE ""
			home_page()
		CASE "news"
			news_page()
		CASE "services"
			services_page()
		CASE "gallery"
			gallery_page()
		CASE "profiles"
			profiles_page()
		CASE "employment"
			employment_page()
		CASE "affiliates"
			affiliates_page()
		CASE "login"
			login_page()
		CASE "contactus"
			contactus_page()
		CASE "staff"
			staff_only()
		CASE ELSE
			error_page()
	END SELECT
	
	
	IF page <> "" THEN
%>
													</td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
<% END IF %>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>