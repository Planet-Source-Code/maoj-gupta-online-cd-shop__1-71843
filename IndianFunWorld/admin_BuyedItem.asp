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
  <li><a href="admin_BuyedItem.asp">Category +</a>
    <ul>
      <li><a href="admin_BuyedItem.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="admin_BuyedItem.asp">DVD+</a>
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
      <br><br>
      <%
      dim conn,rec
      dim query
       set conn=Server.CreateObject("ADODB.Connection")
       set rec=Server.CreateObject("ADODB.RecordSet")
       conn.Open="Provider=Microsoft.Jet.oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
       query="Select * from [BuyedItem]"
       rec.open query,conn
       if rec.EOF then
           response.write"<h3>Nothing Buyed Till Now<br><br><a href=admin.asp>Click Here to Return on Administration Tools</a></h3>"
           response.write"<br><br><br><br><br><br><br><br><br><br><br>"
       else
           response.write"<form method=post action=delete_BuyedItem.asp><table border=2><tr>"
           response.write"<td><h3>Select</h3></td>"
           response.write"<td><h3>Disc Title</h3></td>"
           response.write"<td><h3>Category</h3></td>"
           response.write"<td><h3>Type</h3></td>"
           response.write"<td><h3>Purchased By</h3></td>"
           response.write"<td><h3>Purchase Date</h3></td>"
           response.write"<td><h3>Delete</h3></td></tr>"
           do while not rec.EOF
                response.write"<tr>"
                response.write"<td><input type=radio name=id value="& rec.Fields(0) &" </td>"
                response.write"<td><h3>"& rec.Fields(1) &"</h3></td>"
                response.write"<td><h3>"& rec.Fields(2) &"</h3></td>"
                response.write"<td><h3>"& rec.Fields(3) &"</h3></td>"
                response.write"<td><h3>"& rec.Fields(4) &"</h3></td>"
                response.write"<td><h3>"& rec.Fields(5) &"</h3></td>"
                response.write"<td><input type=Submit value=Delete></td></tr>"
                rec.movenext
           loop
           response.write"</table><br><br>"
           response.write"</form>"
       End If
      %>
      
      <br><br><br><br><br><br><br><br>
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
