<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if

	Dim strSQL, StrDate, show_id, topic, news

	show_id = request.querystring("ID")

	topic = Request.form("topic")
	topic = Replace(topic,"'","")
	news = Request.form("news")
	news = Replace(news,"'","")

	strSQL = "UPDATE publicnews SET topic = '" & topic & "', news = '" & news & "' WHERE id = "& show_id &""

	If Request.Form("submit") = "submit" Then
%>
		<!-- #include virtual='/DBconnect.wiz'//-->
<%
	End If
	Response.redirect "/news/default.asp?id="& show_id
%>

