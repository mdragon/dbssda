<%@ EnableSessionState=False LANGUAGE="VBSCRIPT"%>
<%
title="Submission Successful"
headcolor="#800000"
color="#800000"
'headcolor="#339933"
%>
<!--#include file="inc_head.asp"-->
<!--#include file="inc_nav.asp"-->
<% if Request.QueryString("error-on-send") = 1 then %>
There was an error while emailing your information.  Even with the error, it has been stored and we should be recieving it, 
but if you haven't heard from us in a few days, you may want to try again, or <a href="contact.asp">send us an email</a>.
<% 
else
%>
Your contact information has been submitted, we will be contacting you soon.
<%
end if
%>
<br/><br/><br/><br/><br/>
<!--#include file="inc_footer.asp"-->
