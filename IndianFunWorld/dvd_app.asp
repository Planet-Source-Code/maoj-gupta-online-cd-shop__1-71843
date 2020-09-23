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
      <%
      if not(request.cookies("uname")="")then
        response.write "<h4>Welcome to IndianFunWorld "& request.cookies("uname")  &"&nbsp;&nbsp;&nbsp;&nbsp;<a href=Logout.asp>Logout</a> &nbsp;&nbsp;&nbsp;&nbsp;<a href=setting.asp>Setting</a>"
        response.write ""
        response.write ""
      else
        response.write "<h4>Welcome to IndianFunWorld Guest</h4>"
      End If
      If not(request.cookies("uname")="")then
          response.write"&nbsp;&nbsp;&nbsp;&nbsp;<a href=cart.asp>View Cart</a>"
      End If
      If (request.cookies("uname")="admin")then
          response.write"&nbsp;&nbsp;&nbsp;&nbsp;<a href=admin.asp>Administration Tools</a></h3>"
      End If
      %>
<form method=post action=search.asp>
<input type=text name="searchname">
<select name="type">
<option>All</option>
<option>CD</option>
<option>DVD</option>
</select>
<select name="category">
<option>All</option>
<option>Movie</option>
<option>Application</option>
<option>Game</option>
<option>Music</option>
</select>
<input type=Submit value="Search">
</form>

<div id="wrapper">
      <dl id="browse">
      <ul id="navmenu">
  <li><a href="indexmain.asp">Home</a></li>
  <li><a href="dvd_app.asp">Category +</a>
    <ul>
      <li><a href="dvd_app.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="dvd_app.asp">DVD+</a>
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
           dim conn,rec
           set conn=Server.CreateObject("ADODB.Connection")
           set rec=Server.CreateObject("ADODB.RecordSet")
           conn.Open="Provider=Microsoft.Jet.Oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
           rec.Open "Select * from [Disc] where Category='Application' AND Type='DVD'",conn
           
           do while not rec.EOF 
          response.write "<h3>" & rec.Fields(1) & "</h3>"
          response.write "<img src='"& rec.Fields(2) &"' width=91 height=99 alt=photo6 class=left />"
          response.write"<p><b>Price:</b> <b>"& rec.Fields(3) & "</b> &amp; eligible for FREE Super Saver Shipping on orders over <b>150 Rs</b>.</p>"
          response.write"<p><b>Availability:</b> Usually ships within 24 hours</p>"
          response.write"<form method=post action=buy.asp><br><p align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=Submit value=Buy name=b1>"
          response.write"&nbsp;&nbsp;<input type=Submit value='Add To Cart' name=b1>"
          response.write"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='"& rec.Fields(4) &"'>More Information</a>"
		  response.write"&nbsp;&nbsp;&nbsp;<b>Available Quantity = "& rec.fields(7) &"</b></p>"
		  response.write"<input type=hidden name=cd value='"& rec.fields(1) &"'></form>"
          response.write"<div class=clear></div>"
          response.write"<div class='clear br'></div>"
          rec.movenext
          loop
          set conn=nothing
          set rec=nothing
          %>
        <div class="clear"></div>
      </div>
      <!-- end .inner -->
    </div>
    <!-- end body -->
    <div class="clear"></div>
    <div id="footer"> &nbsp;Best View on Mozilla FireFox at 1024X786 resolution.
      
      <!-- end footnav -->
    </div>
    <!-- end footer -->
  </div>
  <!-- end inner -->
</div>
<!-- end wrapper -->

</body>
</html>
