<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if

	Dim strSQL, show_id

	show_id = request.querystring("ID")

	strSQL = "delete * from employment WHERE id = "& show_id &""

%>
		<!-- #include virtual='/DBconnect.wiz'//-->
<%
	Response.redirect "/staff_only/editjob.asp"
%>

