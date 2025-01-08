<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if

	Dim strSQL, StrDate, show_id, show_date, show_location, show_client, show_type, fire_hour, fire_min, fire_am_pm, start_hour, start_min, start_am_pm, finish_hour, finish_min, finish_am_pm, money_collection, show_price, show_description, staff_needed

	show_id = request.querystring("ID")

	show_date = Request.form("show_date")
	show_date = Replace(show_date,"'","")
	show_location = Request.form("show_location")
	show_location = Replace(show_location,"'","")
	show_client = Request.form("show_client")
	show_client = Replace(show_client,"'","")
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

	strSQL = "UPDATE showdetails SET show_date = '" & show_date & "', show_location = '" & show_location & "', show_client = '" & show_client & "', show_type = '" & show_type & "', fire_hour = '" & fire_hour & "', fire_min = '" & fire_min & "', fire_am_pm = '" & fire_am_pm & "', start_hour = '" & start_hour & "', start_min = '" & start_min & "', start_am_pm = '" & start_am_pm & "', finish_hour = '" & finish_hour & "', finish_min = '" & finish_min & "', finish_am_pm = '" & finish_am_pm & "', money_collection = '" & money_collection & "', show_price = '" & show_price & "', show_description = '" & show_description & "', staff_needed = '" & staff_needed & "' WHERE id = "& show_id &""
	response.write strSQL

	If Request.Form("Submit") = "submit" Then
%>
		<!-- #include virtual='/DBconnect.wiz'//-->
<%
	End If
	Response.redirect "/staff_only/shows.asp?id="& show_id
%>

