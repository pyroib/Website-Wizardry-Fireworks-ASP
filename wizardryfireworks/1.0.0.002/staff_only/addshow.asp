<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if
%>
<!-- #include virtual="/header.wiz"//-->
	<td width="100%" height="5" valign="bottom" class="maintble"></td>
</tr>
<tr> 
	<td height="10" valign="bottom" align="left" bgcolor="black">
		<div class="pageheader">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Add new show details</div>
	</td>
</tr>
<tr>
	<td valign="top" height="485" align="center" style="{padding: 10px; border: 1px ridge #FFFF66;}">
		<table width="550" border="0" cellpadding="0" cellspacing="0" class="text">

<%
		Dim strSQL, staff_id, month, show_date, show_id, show_location, client_name, show_type, fire_hour, fire_min, fire_am_pm, start_hour, start_min, start_am_pm, finish_hour, finish_min, finish_am_pm, money_collection, show_price, show_description, staff_needed, amount_owing

	If Request.Form("Submit") = "submit" Then

		staff_id = Session("staff_id")

	show_location = Request.form("show_location")
	show_location = Replace(show_location,"'","")
	client_name = Request.form("client_name")
	client_name = Replace(client_name,"'","")
	show_type = Request.form("show_type")
	show_type = Replace(show_type,"'","")
	fire_hour = Request.form("fire_hour")
	fire_hour = Replace(fire_hour,"'","")
	fire_min = Request.form("fire_min")
	fire_min = Replace(fire_min,"'","")
	fire_am_pm = Request.form("fire_am_pm")
	fire_am_pm = Replace(fire_am_pm,"'","")
	start_hour = Request.form("start_hour")
	start_hour = Replace(start_hour,"'","")
	start_min = Request.form("start_min")
	start_min = Replace(start_min,"'","")
	start_am_pm = Request.form("start_am_pm")
	start_am_pm = Replace(start_am_pm,"'","")
	finish_hour = Request.form("finish_hour")
	finish_hour = Replace(finish_hour,"'","")
	finish_min = Request.form("finish_min")
	finish_min = Replace(finish_min,"'","")
	finish_am_pm = Request.form("finish_am_pm")
	finish_am_pm = Replace(finish_am_pm,"'","")
	money_collection = Request.form("money_collection")
	money_collection = Replace(money_collection,"'","")
	show_price = Request.form("show_price")
	show_price = Replace(show_price,"'","")
	show_description = Request.form("show_description")
	show_description = Replace(show_description,"'","")
	staff_needed = Request.form("staff_needed")
	staff_needed = Replace(staff_needed,"'","")
	amount_owing = Request.form("amount_owing")
	amount_owing = Replace(amount_owing,"'","")

		show_date = Request.form("show_day") &"/"& Request.form("show_month") &"/"&Request.form("show_year")
		strSQL = "INSERT INTO showdetails (show_date, show_location, show_client, show_type, fire_hour, fire_min, fire_am_pm, start_hour, start_min, start_am_pm, finish_hour, finish_min, finish_am_pm, money_collection, show_price, show_description, staff_needed) VALUES ('" & show_date & "', '" & show_location & "', '" & client_name & "', '" & show_type & "', '" & fire_hour & "', '" & fire_min & "', '" & fire_am_pm & "', '" & start_hour & "', '" & start_min & "', '" & start_am_pm & "', '" & finish_hour & "', '" & finish_min & "', '" & finish_am_pm & "', '" & amount_owing & "', '" & show_price & "', '" & show_description & "', '" & staff_needed & "')"
		%><!-- #include virtual="/DBconnect.wiz"//--><%
		Response.redirect ("../staff_only/shows.asp")

	else
%>


			<form method="post" action="/staff_only/addshow.asp">
				<tr> 
					<td colspan="4" valign="top" height="30">&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="4" valign="top"><img src="/images/logosmall.gif" alt="wizardry Logo" /><br />:: Add new show ::<br /><br /></td>
				</tr>
				<tr> 
					<td valign="middle" height="50" align="right"> Date :</td>
					<td valign="middle" height="50"   align="left">
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
							<option value="2005" >2005</option>
							<option value="2006" >2006</option>
						</select>
					</td>
					<td valign="middle" height="50"   align="right">Location :</td>
					<td valign="middle" height="50"   align="left"><input type="text" name="show_Location" size="30" maxlength="50"></td>
				</tr>
				<tr> 
					<td valign="middle" height="50"   align="right">Clients Name :</td>
					<td valign="middle" height="50"   align="left"><input type="text" name="client_name" size="30" maxlength="50"></td>
					<td valign="middle" height="50"   align="right">Type :</td>
					<td valign="middle" height="50"   align="left"><input type="text" name="show_type" size="30" maxlength="30"></td>
				</tr>
				<tr> 
					<td valign="middle" height="50"   align="right">Firing Time :</td>
					<td valign="middle" height="50"   align="left">
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
					<td valign="middle" height="50"   align="right">Work start :</td>
					<td valign="middle" height="50"   align="left">
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
				</tr>
				<tr> 
					<td valign="middle" height="50"   align="right">Work Finish :</td>
					<td valign="middle" height="50"   align="left">
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
					<td valign="middle" height="50" align="right">Collect Money :</td>
					<td valign="middle" height="50" align="left">
						<select name="amount_owing">
							<option value="yes" >Yes</option>
							<option value="no" >No</option>
						</select>
					</td>
				</tr>
				<tr> 
					<td valign="middle" height="50" align="right">Price :</td>
					<td valign="middle" height="50" align="left"><input type="text" name="show_Price" size="30" maxlength="25"></td>
					<td valign="middle" height="50" align="right">Description :</td>
					<td valign="middle" height="50" align="left"><input type="text" name="show_description" size="30" maxlength="255"></td>
				</tr>
				<tr> 
					<td valign="middle" height="50" align="right">Staff Needed :</td>
					<td valign="middle" height="50" align="left">
						<select name="staff_needed">
							<option value="1" >1</option>
							<option value="2" >2</option>
							<option value="3" >3</option>
							<option value="4" >4</option>
							<option value="5" >5</option>
							<option value="6" >6</option>
						</select>
					</td>
					<td colspan="2" valign="middle" align="center" height="50"><input type="submit" name="Submit" value="submit"></td>
				</tr>
			</form>
<% end if %>
			<tr>
				<td colspan="4"><a href="/staff_only"> Back to Staff Page </a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td valign="bottom" align="center">
		<div style="{font-size:10px; color:white;}">&copy; Wizardry Fireworks Pty Ltd</div>
	</td>
<!-- #include virtual="/footer.wiz"//-->
