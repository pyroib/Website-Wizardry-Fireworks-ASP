<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if

	Dim strSQL, staffid

	staffid = request.querystring("ID")

	strSQL = "delete * FROM logins WHERE staff_ID = '"& staffid &"'"

	%><!-- #include virtual='/DBconnect.wiz'//--><%

	Response.redirect ("/staff_only/editlogin.asp")
%>

