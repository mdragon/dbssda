<%@ EnableSessionState=False LANGUAGE="VBSCRIPT"%>
<%
title="Contact Me"
headcolor="#800000"
color="#800000"
'headcolor="#339933"

name	= Request.QueryString("name")
company = Request.QueryString("company")
email	= Request.QueryString("email")
phone	= Request.QueryString("phone")
msg		= Request.QueryString("msg")
address1 = Request.QueryString("address1")
address2 = Request.QueryString("address2")
city = Request.QueryString("city")
state = Request.QueryString("state")
zip = Request.QueryString("zip")
%>
<!--#include file="inc_head.asp"-->
<!--#include file="inc_nav.asp"-->
<%
If msg = 1 Then Response.Write("<span style=""color: red;"">We need at least your name and one way to contact you.</span>")
%>
<form action="storeinfo.aspx" method="post">
<table>
<tr><td><b>Name:</b></td><td><input name="name" value="<%= name%>"></td></tr>
<tr><td><b>Company:</b></td><td><input name="company" value="<%= company%>"></td></tr>
<tr><td><b>Email:</b></td><td><input name="email" value="<%= email%>"></td></tr>
<tr><td><b>Phone:</b></td><td><input name="phone" value="<%= phone%>" size="12"></td></tr>
<tr><td><b>Address 1:</b></td><td><input name="address1" value="<%= address2%>" size="50" ID="Text1"></td></tr>
<tr><td><b>Address 2:</b></td><td><input name="address2" value="<%= address2%>" size="50" ID="Text2"></td></tr>
<tr><td><b>City:</b></td><td><input name="city" value="<%= city%>" size="30" ID="Text3"></td></tr>
<tr><td><b>State:</b></td><td><input name="state" value="<%= state%>" size="2" ID="Text4"></td></tr>
<tr><td><b>Zip:</b></td><td><input name="zip" value="<%= zip%>" size="9" ID="Text5"></td></tr>
</table>
<input type="submit" value="Submit">
</form>

<!--#include file="inc_footer.asp"-->
