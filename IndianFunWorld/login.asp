<%@language=VBScript%>
<%Option Explicit%>
<html>
<head>
<title>IndianFunWorld</title>
<link href="style.css" rel="stylesheet" type="text/css" />
<!--[if gte IE 5.5]>
<script language="JavaScript" src="ie.js" type="text/JavaScript"></script>
<![endif]-->
</head>
<body>
<div id="header">
      <img src="Images/banner.gif" width="1007" height="140" alt="" /></div>
      <!-- end header -->
      <br>


<div id="wrapper">
      <dl id="browse">
      <ul id="navmenu">
  <li><a href="indexmain.html">Home</a></li>
  <li><a href="login.asp">Category +</a>
    <ul>
      <li><a href="login.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="login.asp">DVD+</a>
        <ul>
          <li><a href="dvd_movie.asp">Movie</a></li>
          <li><a href="dvd_app.asp">Application</a></li>
          <li><a href="dvd_game.asp">Games</a></li>
          <li><a href="dvd_music.asp">Music</a></li>
        </ul>
      </li>
    </ul>
  </li>
      <li><a href="login.html">Login</a></li>
      <li><a href="register.html">Register</a></li>
      <li><a href="feedback.html">FeedBack!!!</a></li>
        
  <li><a href="Aboutus.html">About Us</a></li>
  <li><a href="Contact.html">Contact Us</a></li>
</ul>
</dl>
<div id="inner">
    <div id="body">
      <div class="inner">
<%
dim rec,conn
dim uname,pass
dim query
set conn=Server.CreateObject("ADODB.Connection")
set rec=Server.CreateObject("ADODB.RecordSet")
conn.Open"Provider=Microsoft.Jet.Oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
uname=request.form("uname")
pass=request.form("pass")
query="Select * from [user] where uname='"& uname &"' AND pass='"& pass &"'"
rec.open query,conn
if (rec.EOF)then
   response.write"<h3>Username or Password is incorrect<br><br><a href=login.html>Click here to go to login Page</a></h3>"
  response.write""
else
    response.cookies("uname")=uname
    response.cookies("uname").Expires = #May 1, 2010# 
    response.write"<br><h3>You Successfully Logged in.<br><br><a href=indexmain.asp>Click Here to return on index page</a></h3>"
  End If
set conn=nothing
set rec=nothing
'rec.close
'conn.close
%>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
</div>
    <!-- end inner -->
    </div>
    <!-- end body -->
    <div class="clear"></div>
    <div id="footer"> &nbsp;Best View on Mozilla FireFox at 1024X786 resolution.</div>
    <!-- end footer -->
    </div>
  <!-- end inner -->
</div>
<!-- end wrapper -->
</body>
</html>