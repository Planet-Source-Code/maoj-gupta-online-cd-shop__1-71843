<%@language=VBScript%>
<%Option Explicit%>
<html>
<body>
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
  <li><a href="signup.asp">Category +</a>
    <ul>
      <li><a href="signup.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="signup.asp">DVD+</a>
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
dim name,uname,pass,repass
dim address,email,state
dim ques,ans,ctno
dim query
set conn=Server.CreateObject("ADODB.Connection")
set rec=Server.CreateObject("ADODB.RecordSet")
conn.Open"Provider=Microsoft.Jet.Oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
name=trim(request.form("name"))
uname=trim(request.form("uname"))
pass=trim(request.form("pass"))
repass=trim(request.form("repass"))
address=trim(request.form("address"))
email=trim(request.form("email"))
ques=trim(request.form("question"))
ans=trim(request.form("ans"))
state=trim(request.form("state"))
ctno=trim(request.form("ctno"))
query="Select * from [user] where uname='"& uname &"'"

rec.Open query,conn
if(name="" or uname="" or pass="" or repass="" or address="" or email="" or state="" or ans="" or ctno="")then
    response.write "<br><h3>Please fill all fields</h3>"
    response.write "<br><h3><a href=javascript:history.back(1);>Click here to register an account</a></h3>"
else
    If (email <> "" AND inStr(email,"@") <> 0 AND inStr(email,".") <> 0) THEN
        if (rec.EOF)then
            if(pass=repass)then
            query="Insert into [user] values('"& uname &"','"& pass &"','"& name &"','"& address &"',"& ctno &",'"& ques &"','"& ans &"','"& email &"','"& state &"')"
            response.write "<br><h3>Account Created Successfully<br><br><a href=login.html>Click here to go on Login Page</a></h3>"
            conn.Execute(query)
            else
            response.write "<br><h3>Passoword is not matching<br><br><a href=register.html>Click here to go back<a></h3>"
            End if
        else
            response.write"<br><h3>User Already Exists<br><br><a href=register.html>Click Here to Return on SignUp Page</a><h3>"
        End If
    else
         response.Write "<br><h3>Email Address is not valid<br><br><a href=register.html>Click here to go back<a></h3>"
    End If
end if
%>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
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
