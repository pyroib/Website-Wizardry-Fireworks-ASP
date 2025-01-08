<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if

	Dim strSQL, StrDate, id, staffid, staffpw, permit

	id = request.querystring("ID")

	staffid = Request.form("staffid")
	staffid = Replace(staffid,"'","")
	staffpw = Request.form("staffpw")
	staffpw = Replace(staffpw,"'","")
	permit = Request.form("permit")
	permit = Replace(permit,"'","")

	strSQL = "UPDATE logins SET staff_ID = '" & staffid & "', staff_pw = '" & staffpw & "', permit = '" & permit & "' WHERE staff_ID= '"& id &"'"
	response.write strSQL

	If Request.Form("Submit") = "submit" Then
%>
		<!-- #include virtual='/DBconnect.wiz'//-->
<%
	End If
	Response.redirect "/staff_only/editlogin.asp"
%>

