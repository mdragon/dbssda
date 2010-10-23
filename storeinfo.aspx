<%@ import Namespace="System.Web.Mail" %>
<%@ Page Language="vb" debug="true"%>
<%
		Dim formdata As NameValueCollection = Request.Form()
		
		Dim name	as String
		Dim company as String
		Dim email	as String
		Dim phone	as String
		Dim address1	as String
		Dim address2	as String
		Dim city	as String
		Dim state	as String
		Dim zip		as String
		
		Dim strBody as String
		
		name	= formdata.item("name")
		company	= formdata.item("company")
		email	= formdata.item("email")
		phone	= formdata.item("phone")
		address1 = formdata.item("address1")
		address2 = formdata.item("address2")
		city	= formdata.item("city")
		state	= formdata.item("state")
		zip		= formdata.item("zip")
		
		If name <> "" And (email <> "" Or phone <> "" OR (address1 <> "" AND city <> "" AND state <> "")) Then

			'Create an instance of the MailMessage class
			Dim objMM as New MailMessage()


			'Set the subject
			objMM.Subject = "Web Site Info Submission"

			'Set the properties
			If email.tolower.indexof("subject:") > -1 _
				Or name.tolower.indexof("subject:") > -1 _
				Or phone.tolower.indexof("subject:") > -1 _
				Or email.tolower.indexof("dbssda.com") > -1 _
				Or company.tolower.indexof("subject:") > -1 Then
				
				objMM.To = "mdragon@gmail.com"
				objMM.Subject = "Possible Submission Spam"
			Else
				objMM.To = "sdragon@dbssda.com, mdragon@jhu.edu, mdragon@gmail.com, matt.dragon@jhu.edu"
			End IF
			objMM.From = "infomailer@dbssda.com"

			'Send the email in text format
			objMM.BodyFormat = MailFormat.Text
			'(to send HTML format, change MailFormat.Text to MailFormat.Html)

			'Set the priority - options are High, Low, and Normal
			objMM.Priority = MailPriority.Normal

			strBody = "Name:    " & name & vbCR
			strBody = strBody & "Company: " & company & vbCR
			strBody = strBody & "Email:   " & email & vbCR
			strBody = strBody & "Phone:   " & phone & vbCR
			strBody = strBody & "Address: " & vbCR
			strBody = strBody & address1 & vbCR
			if( address2 <> "" ) Then strBody = strBody & address2 & vbCR
			strBody = strBody & city & ", " & state & " " & zip & vbCR
						
			strBody = strBody & "Submitted: " & date.now.toString()

			'Set the body - use VbCrLf to insert a carriage return
			objMM.Body = strBody
			
			Try
				Dim out As IO.StreamWriter
				out = System.IO.File.CreateText(server.mappath("info_data_files\" & Date.Now.ToString("yyyy-MM-dd_hh-mm-ss") & ".txt"))
				out.WriteLine(strBody)
				out.Close()
			Catch ex as Exception
				response.write(ex.message)
			End Try
			
			Dim oldServ as String
			oldServ = SmtpMail.SmtpServer
			'SmtpMail.SmtpServer = "bizmail.cnjnet.com"
			SmtpMail.SmtpServer = "smtp.1and1.com"
			try
				SmtpMail.Send(objMM)
			catch ex as Exception
				try
					Dim out2 As IO.StreamWriter
					out2 = System.IO.File.CreateText(server.mappath("info_data_files\error" & Date.Now.ToString("yyyy-MM-dd_hh-mm-ss") & ".txt"))
					out2.WriteLine(ex.Message)
					out2.Close()
				catch ex2 as Exception
								
				end try
				'response.write(ex.message)
				Response.Redirect("thanks.asp?error-on-send=1")
			end try
			SmtpMail.SmtpServer = oldServ
			
			Response.Redirect("thanks.asp")
		Else
			Response.Redirect("default.asp?msg=1&name="& name & "&company=" & company & "&email=" & email & "&phone=" & phone & "&address1=" & address1 & "&address2=" &address2 & "&city=" &city & "&state=" &state & "&zip=" &zip)
		End If
%>
