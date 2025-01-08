<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if

	Dim strSQL, StrDate, show_id, jobdescription, requirements, wage, needed

	show_id = request.querystring("ID")


	jobdescription = Request.form("jobdescription")
	jobdescription = Replace(jobdescription,"'","")
	requirements = Request.form("requirements")
	requirements = Replace(requirements,"'","")
	wage = Request.form("wage")
	wage = Replace(wage,"'","")
	needed = Request.form("needed")
	needed = Replace(needed,"'","")
 
	strSQL = "UPDATE employment SET jobdescription = '" & jobdescription & "', requirements = '" & requirements & "', wage = '" & wage & "', needed = '" & needed & "' WHERE id = "& show_id &""
	response.write strSQL

	If Request.Form("Submit") = "submit" Then
		%><!-- #include virtual='/DBconnect.wiz'//--><%
	End If

	Response.redirect "/employment"
%>

