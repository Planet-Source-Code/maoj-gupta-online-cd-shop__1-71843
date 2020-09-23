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
  <li><a href="cart.asp">Category +</a>
    <ul>
      <li><a href="cart.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="cart.asp">DVD+</a>
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
<!-- end navmenu -->
</dl>
<!-- end browse -->
  <div id="inner">
    <div id="body">
      <div class="inner">
      <h1 align=center>Welcome <%response.write(request.cookies("uname"))%> to Your Cart</h1><br><br>

      <%
          
           dim conn,rec
           set conn=Server.CreateObject("ADODB.Connection")
           set rec=Server.CreateObject("ADODB.RecordSet")
           conn.Open="Provider=Microsoft.Jet.Oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
           rec.Open "Select * from [Cart] where Uname='"& request.cookies("uname") &"'",conn
           if rec.EOF then
              response.write"<br><h3>No Item in Cart<br><br><a href=indexmain.asp>Click Here to return on Index Page</a></h3>"
              response.write"<br><br><br><br><br><br><br><br>"
           else
               response.write"<form name=cart method=post action=cartbuy.asp>" 
               response.write"<table border=2 align=center><tr>"
               response.write"<td><h3>Select</h3></td>"
               response.write"<td><h3>Disc Name</h3></td>"
               response.write"<td><h3>Category</h3></td>"
               response.write"<td><h3>Type</h3></td>"
               response.write"<td><h3>Quantity</h3></td></tr>"
              do while not rec.EOF 
                  response.write"<tr>"
                  response.write"<td align=center><input type=radio name=id value="& rec.fields(0) &"></td>"
                  response.write"<td><h3>"& rec.fields(2) &"</h3></td>"
                  response.write"<td><h3>"& rec.fields(3) &"</h3></td>"
                  response.write"<td><h3>"& rec.fields(4) &"</h3></td>"
                  response.write"<td><h3>"& rec.fields(5) &"</h3></td></tr>"
                  rec.movenext
              loop
              response.write"</table>"
              response.write"<br><br><br><br><br><br><input type=Submit name=action value='Buy Now'>"
              response.write"&nbsp;&nbsp;&nbsp;&nbsp;<input type=Submit name=action value='Remove From Cart'>"
              response.write"</form>"
           End If

      %>
<br><br><br><br><br><br><br><br><br>
      
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
