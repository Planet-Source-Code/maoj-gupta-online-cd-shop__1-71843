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
  <li><a href="indexmain.asp">Home</a></li>
  <li><a href="cartbuy.asp">Category +</a>
    <ul>
      <li><a href="cartbuy.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="cartbuy.asp">DVD+</a>
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

<%
    dim conn,rec,rec1
    dim query,qty,qty1
    dim id,action,qty2
    dim discid,category,t,discname
    set conn=Server.CreateObject("ADODB.Connection")
    set rec=Server.CreateObject("ADODB.RecordSet")
    set rec1=Server.CreateObject("ADODB.RecordSet")
    conn.Open="Provider=Microsoft.Jet.Oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
    id=request.form("id")
    if(id="")Then
        response.write"<h3>Please Select any Disc to buy or Remove<br><a href=cart.asp>Click Here to return on your cart</a><h3>"
    Else
          query="Select * from [Cart] where DiscId="& id
          rec.Open query,conn
          rec1.Open"Select * from [Disc] where ID="& id,conn
          action=request.form("action")
          discid=rec1.fields(0)
          category=rec1.fields(5)
          t=rec1.fields(6)
          discname=rec1.fields(1)
          qty=cint(rec.fields(5))
          qty1=cint(rec1.fields(7))
          if (action="Buy Now") then
               qty2=qty1-qty
               if(qty2<1)then
                  response.write"<br><h3>Dont have enough stock for "& rec1.fields(1) &"You Can Buy This later<br><br><a href=cart.asp>Click here to return on your Cart</a></h3>"
                  response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
               else
                  conn.Execute"Update [Disc] Set Quantity="& qty2 &" where ID="& id
                  conn.Execute "Delete * from [Cart] where DiscID="& id
                  response.write"<br><h3>You Buy the "& rec1.fields(1) &"<br><br><a href=cart.asp>Click here to return on your Cart</a><h3>"
                  query="Insert Into [BuyedItem] values('"& discid &"','"& discname &"','"& category &"','"& t &"','"& request.cookies("uname") &"','"& Now() &"')"
                  conn.Execute(query)
                  response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
               End If
           else
                conn.Execute "Delete * from [Cart] where DiscID="& id
                response.write"<h3>"& rec1.fields(1) &" Deleted From your Cart<br><br><a href=cart.asp>Click here to return on your Cart</a></h3>"
                response.write"<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
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
