<%@ Language=VBScript %>
<% Option Explicit %>
<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%
DIM contact_name, enquiry, email

			IF Request.Form("submit") = "submit" THEN

				IF Request.form("email") <> "" THEN
					contact_name = Request.form("contact_name")
					contact_name = Replace(contact_name,"'","")
					contact_name = Replace(contact_name,"""","")
					
					email = Request.form("email")
					email = Replace(email,"'","")
					email = Replace(email,"""","")
										
					enquiry = Request.form("enquiry")
					enquiry = Replace(enquiry,"'","")
					enquiry = Replace(enquiry,"""","")
				END IF
				

				Const cdoSendUsingMethod = _
				"http://schemas.microsoft.com/cdo/configuration/sendusing"
				Const cdoSendUsingPort = 2
				Const cdoSMTPServer = _
				"http://schemas.microsoft.com/cdo/configuration/smtpserver"
				Const cdoSMTPServerPort = _
				"http://schemas.microsoft.com/cdo/configuration/smtpserverport"
				Const cdoSMTPConnectionTimeout = _
				"http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
				Const cdoSMTPAuthenticate = _
				"http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
				Const cdoBasic = 1
				Const cdoSendUserName = _
				"http://schemas.microsoft.com/cdo/configuration/sendusername"
				Const cdoSendPassword = _
				"http://schemas.microsoft.com/cdo/configuration/sendpassword"


				Dim objConfig ' As CDO.Configuration
				Dim objMessage ' As CDO.Message
				Dim Fields ' As ADODB.Fields


' Then define the CDO configuration ...
'-----------------------------------------------------
				Set objConfig = Server.CreateObject("CDO.Configuration")
				Set Fields = objConfig.Fields

				With Fields
					.Item(cdoSendUsingMethod) = cdoSendUsingPort
					.Item(cdoSMTPServer) = "mail.wizardryfireworks.com"
					.Item(cdoSMTPServerPort) = 25
					.Item(cdoSMTPConnectionTimeout) = 20
					.Item(cdoSMTPAuthenticate) = cdoBasic
					.Item(cdoSendUserName) = "site@wizardryfireworks.com"
					.Item(cdoSendPassword) = "pyr0sg"
	
					.Update
				End With


' ... so You can create and send an email.
'-----------------------------------------------------
				Set objMessage = Server.CreateObject("CDO.Message")
				Set objMessage.Configuration = objConfig
				
				
				With objMessage
					.To = "Site Info Request <info@wizardryfireworks.com>"
					.bcc = "Wizardry enquiry <webmaster@wizardryfireworks.com>)"
					.From = "Site Info Request <site@wizardryfireworks.com>"
					.Subject = "Request For Information From Site User"
					.TextBody = ""& enquiry & contact_name & email &""
					.Send
				End With


' Don't forget to remove the objects once You're done
'-----------------------------------------------------
				Set Fields = Nothing
				Set objMessage = Nothing
				Set objConfig = Nothing 









			END IF


	%>
<form method="post" name="signup" action="/contactus_form.asp">
	<table  width="590" border="1" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<td width="144" align="right" class="Signup">Name : </td>
			<td width="243" align="left">
				<input name="contact_name" type="text" size="20" maxlength="10" />
			</td>
		</tr>
		<tr><td height="12" colspan="2"></td></tr>
		<tr>
			<td align="right" class="Signup">* Email Address : </td>
			<td align="left">
				<input name="email" type="text" size="20" maxlength="50" />
			</td>
		</tr>
		<tr><td height="12" colspan="2"></td></tr>
		<tr>
			<td align="right" class="Signup">* Enquiry :</td>
			<td align="left">
				<textarea name="enquiry" cols="50" rows="10" name="textarea"></textarea>
			</td>
		</tr>
		<tr><td colspan="2" class="conditions">&nbsp;</td><tr>
		<tr>
			<td>&nbsp;</td>
			<td class="submit">
				<input name="submit" type="submit" value="submit" />
			</td>
		</tr>
		<tr>
			<td colspan="4"></td>
		</tr>
	</table>
</form>