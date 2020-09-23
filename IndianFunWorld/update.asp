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
  <li><a href="update.asp">Category +</a>
    <ul>
      <li><a href="update.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="update.asp">DVD+</a>
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
dim conn,rec,query
dim name,pass,repass
dim address,ctno,ques,ans
dim email,state,action
name=trim(request.form("name"))
pass=trim(request.form("pass"))
repass=trim(request.form("repass"))
address=trim(request.form("address"))
ctno=trim(request.form("ctno"))
ques=trim(request.form("question"))
ans=trim(request.form("ans"))
email=trim(request.form("email"))
state=trim(request.form("state"))
action=request.form("action")
set conn=Server.CreateObject("ADODB.Connection")
set rec=Server.CreateObject("ADODB.RecordSet")
conn.Open="Provider=Microsoft.Jet.oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
rec.Open "select * from [user] where uname='"& request.cookies("uname") &"'",conn
if(pass=repass)then
    if(action="Update My Account")then
          if(name="" or pass="" or repass="" or address="" or email="" or state="" or ans="" or ctno="")then
              response.write "<h3>Please fill all fields</h3>"
              response.write "<br><h3><a href=setting.asp>Click here to return on your Control Panel</a></h3>"
          else
              If (email <> "" AND inStr(email,"@") <> 0 AND inStr(email,".") <> 0) THEN
                   if(pass=repass) Then
                      query="UPDATE [user] SET Name ='"& name &"' , pass ='"& pass &"' , address ='"& address &"' , ctno ='"& ctno &"' , question ='"& ques &"' , answer ='"& ans &"' , email ='"& email &"' , state ='"& state &"' where uname='"& request.cookies("uname") &"'"
                      conn.Execute(query)
                      response.write "<br><h3>Account updated Successfully<br><br><a href=indexmain.asp>Click here to return on index page</a></h3>"
                   Else
                   response.write "<br><h3>Passoword is not matching<br><br><a href=setting.asp>Click here to you Control Panel</a></h3>"
                   End If
             else
                 response.write "<br><h3>Email Address is Not Valid<br><br><a href=setting.asp>Click here to you Control Panel</a></h3>"   
              End If
          End If
    elseif(action="Delete Account")then
         if request.cookies("uname")="admin" then
              response.write"<br><h3>Admin Account Cannot be Deleted.<br><br><a href=setting.asp>Click Here to Return on Setting Page.</a></h3>"
         Else
              query="DELETE * FROM [user] where uname='"& request.cookies("uname") &"'"
              conn.Execute(query)
              response.cookies("uname")=""
              response.write "<br><h3>Account Deleted Successfully<br><br><a href=indexmain.asp>Click here to return on index page</a></h3>"
         End If
    else          
        rec.CancelUpdate
        response.write "<br><h3>Account update Canceled by User<br><br><a href=indexmain.asp>Click here to return on index page</a></h3>"
    End if
else
    response.write "<h3>Password is not Matching<br><a href=setting.asp>Click here to return on Setting Page</a></h3>"
End If
set conn=nothing
set rec=nothing
%>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
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