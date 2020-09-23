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
<body>
<div id="header">
      <img src="Images/banner.gif" width="1007" height="140" alt="" /></div>
      <!-- end header -->
      <br>
<div id="wrapper">
      <dl id="browse">
      <ul id="navmenu">
  <li><a href="indexmain.asp">Home</a></li>
  <li><a href="buy.asp">Category +</a>
    <ul>
      <li><a href="buy.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="buy.asp">DVD+</a>
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
dim conn,rec,rec1
dim query
dim discname,qty,qty1
dim action
dim category,t,discid
action=request.form("b1")
set conn=Server.CreateObject("ADODB.Connection")
set rec=Server.CreateObject("ADODB.RecordSet")
set rec1=Server.CreateObject("ADODB.RecordSet")
conn.Open"Provider=Microsoft.Jet.Oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
if(request.cookies("uname")="")then
    response.write"<br><h3>You are not Logged in to Buy you must logged in<br><br><a href=login.html>Click here to login</a></h3>"
    response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
else
  If(action="Buy")then
      discname=request.form("cd")
      query="Select * from [Disc] where Title='"& discname &"'"
      rec.Open query,conn
      discid=rec.fields(0)
      category=rec.fields(5)
      t=rec.fields(6)
      qty=cint(rec.fields(7))
      if(qty<1)then
        response.write"<h3>Sorry Disc is out of Stock You can buy this later<br><br><a href=indexmain.asp>Click Here to Return on Index page</a></h3>"
        response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
      else
        qty=qty-1 
        query="UPDATE [Disc] SET Disc.Quantity ="& qty &" where Title='"& discname &"'"
        conn.Execute(query)
        query="Insert Into [BuyedItem] values('"& discid &"','"& discname &"','"& category &"','"& t &"','"& request.cookies("uname") &"','"& Now() &"')"
        conn.Execute(query)
        response.write"<br><br><br><h3>You Buy the "& discname &"<br><br><a href=indexmain.asp>Click Here to Return on Index Page</h3>"
        response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
      End If
  Else
      discname=request.form("cd")
      query="Select * from [Disc] where Title='"& discname &"'"
      rec.Open query,conn
      rec1.Open"Select Quantity from [Cart] where Uname='"& request.cookies("uname") & "' AND DiscTitle='"& discname & "'",conn
      qty=cint(rec.fields(7))
      if(qty<1)then
        response.write"<h3>Sorry Disc is out of Stock You can buy this later<br><br><a href=indexmain.asp>Click Here to Return on Index page</a></h3>"
        response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
      else
        if rec1.EOF then
            query="Insert into [Cart] values("& rec.fields(0) &",'"& request.cookies("uname") &"','"& rec.fields(1) &"','"& rec.fields(5) &"','"& rec.fields(6) &"',1)"
            conn.Execute(query)
            response.write"<br><br><br><h3>"& discname &" is Added to your cart<br><br><a href=indexmain.asp>Click Here to Return on Index Page</h3>"        
            response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
        else
            qty1=cint(rec1.fields(0))+1
            query="Update [cart] Set Quantity="& qty1 &" where uname='"& request.cookies("uname") &"' AND DiscTitle='"& discname &"'"
            conn.Execute(query)
            response.write"<br><br><br><h3>"& discname &" is Added to your cart<br><br><a href=indexmain.asp>Click Here to Return on Index Page</h3>"     
            response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"   
        End If
      End If
	End If
End If
%>

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