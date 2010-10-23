<%@ Page Language="vb" debug="true"%>
<%
    Dim name	as String
    Dim company as String
    Dim email	as String
    Dim phone	as String
    Dim address1	as String
    Dim address2	as String
    Dim city	as String
    Dim state	as String
    Dim zip		as String
    Dim msg as Int32
        
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
<html>
<head>
    <title>Stephen Dragon &amp; Associates</title>
    <link rel="stylesheet" href="screen.css" type="text/css" media="screen,print">
    <link rel="stylesheet" href="ie.css" type="text/css" media="screen,print">
    <link rel="stylesheet" href="main.css" type="text/css" media="screen,print">
</head>
<body>
    <div id="main">
        <div class="color-bar">
            &nbsp;
        </div>
        <div class="content-wrapper">
            <div class="header">
                <div class="header-contact">
                    <img src="images/logo.gif" alt="SDA Logo"/>
                </div>
                <div class="header-contact">
                    <h2>Stephen Dragon &amp; Associates</h2>
                    12 Highmount Ave.<br/>
                    Warren, New Jersey 07059-5454
                </div>
                <div class="header-contact" style="margin-top: 30px;">
                    Phone:<b>(908) 757-7382</b> <br/>
                    Fax:<b>(908) 753-0872</b> <br/>
                    <b><a href="mailto:sdragon@NOSPAMdbssda.com">sdragon@NOSPAMdbssda.com</a></b><br/>
                </div>
                <div style="float: left; clear: both;">&nbsp;</div>
            </div>
            <div class="content">
                <div class="section">
                    <% if Request.QueryString("error-on-send") = 1 then %>
                    There was an error while emailing your information.  Even with the error, it has been stored and we should be recieving it, 
                    but if you haven't heard from us in a few days, you may want to try again, or <a href="mailto:sdragon@dbssda.com">send us an email</a>.
                    <% 
                    else
                    %>
                    Your contact information has been submitted, we will be contacting you soon.
                    <%
                    end if
                    %>
                </div>
            </div>
        </div>
        <div style="clear: left;"></div>
    </div>
</body>
</html>