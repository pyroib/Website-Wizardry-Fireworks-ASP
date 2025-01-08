<%@ Language=VBScript %>
<% Option Explicit %>
<%
	if Session("permit") = "" then
		Response.redirect ("login.asp")
	else if Session("permit") > 1 then
		Response.redirect ("login.asp")
	end if
	End if

	Dim strSQL, StrDate, show_id, affiliate_name, affiliate_link, affiliate_image

	show_id = request.querystring("ID")

	affiliate_name = Request.form("affiliate_name")
	affiliate_name = Replace(affiliate_name,"'","")
	affiliate_link = Request.form("affiliate_link")
	affiliate_link = Replace(affiliate_link,"'","")
	affiliate_image = Request.form("affiliate_image")
	affiliate_image = Replace(affiliate_image,"'","")

	strSQL = "UPDATE affiliates SET affiliatename = '" & affiliate_name & "', affiliatelink = '" & affiliate_link & "', affiliateimage = '" & affiliate_image & "' WHERE id="& show_id &""
	response.write strSQL

	If Request.Form("Submit") = "submit" Then
		%><!-- #include virtual='/DBconnect.wiz'//--><%
	End If

	Response.redirect "/affiliates"
%>

