<%@ Page Language="vb" Debug="true" AutoEventWireup="false"%>
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
    <!--[if lt IE 8]><link rel="stylesheet" href="css/blueprint/ie.css" type="text/css" media="screen, projection"><![endif]-->
    <link rel="stylesheet" href="main.css" type="text/css" media="screen,print">
</head>
<body>
    <div id="main">
        <table class="main-table">
        <td class="color-bar"></td>
        <td>
            <div class="content-wrapper">
                <div class="header">
                    <div class="header-contact">
                        <img src="images/logo.gif" alt="SDA Logo"/>
                    </div>
                    <div class="header-contact">
                        <h2>Stephen Dragon &amp; Associates</h2>
                        12 Highmount Avenue<br/>
                        Warren, New Jersey 07059-5454
                    </div>
                    <div class="header-contact" style="margin-top: 30px;">
                        Phone:<b>(908) 757-7382</b> <br/>
                        Fax:<b>(908) 753-0872</b> <br/>
                        <b><a href="mailto:sdragon@dbssda.com">sdragon@dbssda.com</a></b><br/>
                    </div>
                    <div style="float: left; clear: both;">&nbsp;</div>
                </div>
                <div class="content">
                    <div class="section">
                        <p>Since 1984, <span class="sda">Stephen Dragon &amp; Associates</span> has provided complete service in the area of custom software design as well as modifications to existing financial software products.</p>
                        <p>We are located in Warren, New Jersey, and serve the New York Metropolitan area. We currently provide sales and support for <a href="http://www.cyma.com">CYMA</a> Not-For-Profit and Financial Management products. </p>
                        <p>We were involved in founding the New York Fox User Group and the New Jersey Fox User Group.  We have also been actively involved in many other professional organizations.</p>
                    </div>
                    <div class="section">
                        <h3>CYMA Products</h3>
                        <div class="indent">                    
                            <div style="float: right;">
                                <img src="images/cyma_logo.gif" alt="CYMA logo"/>
                            </div>
                            <div style="float: left; width: 580; margin-right: 5px;">
                                We provide sales and service for <a href="http://www.cyma.com">CYMA</a> Not-For-Profit and Financial Management products.
                                <h4>Not-For-Profit</h4>
                                <div class="indent">
                                    "Not-For-Profit organizations need an accounting software package designed for the specific nature of the nonprofit industry. Since 1980, thousands of nonprofit clients have depended on CYMA Not-For-Profit Accounting Software (NFP) for their financial needs."
                                    <br/><a href="http://www.cyma.com/business-accounting-software/nfp.asp">Read more about CYMA Not-For-Profit</a>
                                </div>
                                <h4>Financial Management</h4>
                                <div class="indent">
                                    "CYMA Financial Management System (FMS) represents the most sophisticated accounting software on the market today for growing mid-sized businesses. CYMA FMS provides accounting functionality powerful enough for the most demanding user and straightforward for new users."
                                    <br/><a href="http://www.cyma.com/business-accounting-software/fms.asp">Read more about CYMA Financial Management</a>
                                </div>
                            </div>
                            <div style="float: left; clear: both;">&nbsp;</div>
                        </div>
                    </div>
                    <div class="section">
                        <h3>Contact Us</h3>
                        <div class="indent">                    
                            Provide your contact information below and we will contact you.<br/>
                            <%
                            If msg = 1 Then Response.Write("<span style=""color: red;"">We need at least your name and one way to contact you.</span>")
                            %>
                            <form action="storeinfo.aspx" method="post" class="indent">
                            <table>
                            <tr><td><b>Name:</b></td><td><input name="name" value="<%= name%>" size="30"></td></tr>
                            <tr><td><b>Company:</b></td><td><input name="company" value="<%= company%>" size="40"></td></tr>
                            <tr><td><b>Email:</b></td><td><input name="email" value="<%= email%>" size="30"></td></tr>
                            <tr><td><b>Phone:</b></td><td><input name="phone" value="<%= phone%>" size="20"></td></tr>
                            <tr><td><b>Address 1:</b></td><td><input name="address1" value="<%= address2%>" size="50"></td></tr>
                            <tr><td><b>Address 2:</b></td><td><input name="address2" value="<%= address2%>" size="50"></td></tr>
                            <tr><td><b>City:</b></td><td><input name="city" value="<%= city%>" size="30"></td></tr>
                            <tr><td><b>State:</b></td><td><input name="state" value="<%= state%>" size="2" ></td></tr>
                            <tr><td><b>Zip:</b></td><td><input name="zip" value="<%= zip%>" size="9"></td></tr>
                            </table>
                            <input type="submit" value="Submit">
                            </form>
                        </div>
                    </div>
                    <div class="section">
                        <h3>About Us</h3>
                        <div class="indent">
                            <p><span class="sda">Stephen Dragon &amp; Associates</span> has provided complete systems analysis, programming, installation, training, and support services to a wide variety 
                            of business and service organizations in the New York City metropolitan area. We are a member of Microsoft's Solutions Channels program. Staff members have conducted seminars at the Raritan Valley 
                            Community College and other locations on Networking PC's as well as database technology and various other topics.</p>
                        </div>
                    </div>
                    <div class="section" style="text-align: center; margin: auto !important; width: 200px;">
                        <!-- BEGIN FINDACCONTINGSOFTWARE.COM BADGE -->
                        <style type="text/css"><!--@import 'http://findaccountingsoftware.com/pp-19709.css';--></style>
                        <div class="FasBadge Standard" style="display:none;"><div class="FasBadgeBody">
                        <p><a class="CompanyName" href="http://findaccountingsoftware.com/reseller/stephen-dragon-and-associates/">Stephen Dragon & Associates</a> has been helping Find <a href="http://findaccountingsoftware.com/">Accounting Software</a> visitors solve their business needs since&nbsp;2004, with products such as:</p>
                        <ul><li><a href="http://findaccountingsoftware.com/directory/cyma/cyma-iv-not-for-profit-edition/">CYMA IV Not-for-Profit Edition </a></li><li><a href="http://findaccountingsoftware.com/directory/cyma/cyma-iv-financial-management-system/">CYMA IV Financial Management System (FMS)</a></li></ul>
                        </div><div class="FasBadgeFooter"></div></div><script type="text/javascript" src="http://findaccountingsoftware.com/pp-19709.js"></script>
                        <!-- END FINDACCONTINGSOFTWARE.COM BADGE -->
                    </div>
                </div>
                <div class="thinline">&nbsp;</div>
                <div class="copyright">Copyright 1999-<%= year(now()) %> Stephen Dragon &amp; Associates</div>
                <div style="float: none; clear: both;"></div>
            </div>
        </td>
        </table>
    </div>
</body>
</html>