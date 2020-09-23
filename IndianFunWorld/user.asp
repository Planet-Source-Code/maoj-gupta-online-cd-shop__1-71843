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
  <li><a href="user.asp">Category +</a>
    <ul>
      <li><a href="user.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="user.asp">DVD+</a>
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
       dim query,uname
       dim action
       action=request.form("action")
       uname=request.form("user")
       if(uname="")Then
            response.write"<h3><br>Please Select Any user<br><br><a href=admin_user.asp>Click Here to Return on User Editing Page</a></h3>"
              response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
       else
             set conn=Server.CreateObject("ADODB.Connection")
             set rec=Server.CreateObject("ADODB.RecordSet")
             conn.Open="Provider=Microsoft.Jet.oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
             query="Select * from [user] where uname='"& uname &"'"
             rec.open query,conn
              
             if (action="Delete") then
                 if not(uname="admin")then
                     query="DELETE * FROM [user] where uname='"& uname &"'"
                     conn.Execute(query)
                     response.write("<br><br><h3>Account Deleted Successfully<br><a href=admin_user.asp>Click here to Return on Administration Tools page</a></h3>")
                 else
                     response.write"<br><br><h3>Admin Account cannot be Deleted<br><a href=admin_user.asp>Click here to Return on Administration Tools page</a></h3>"
                 End If
             else
       %>
             <table cellspacing=10 >
            <tr><td><strong>Name:-</strong></td>
           <td><Input type=Text name=name maxlength=15 value='<%response.write rec.fields(2) %>' readonly></td></tr>

            <tr>
            <td><strong>Username:-</strong></td>
            <td><Input type=Text name=uname maxlength=15 value='<%response.write rec.fields(0) %>' readonly></td></tr>
            <tr></tr>

            <tr>
            <td><strong>Address:</strong></td>
            <td><textarea cols=20 rows=5 name=address readonly><%response.write rec.fields(3) %></textarea></td>
            </tr>

            <tr>
            <td><strong>Contact Number:</strong></td>
            <td><font size=3>+91</font>&nbsp;<input type=text name=ctno value='<%response.write rec.fields(4) %>' readonly></td>
            </tr>

            <tr>
            <td><strong>Email Address:-</strong></td>
            <td><input type=text name=email value='<%response.write rec.fields(7) %>' readonly></td></tr>

            <tr>
            <td><strong>State:-</strong></td>
            <td><input type=text value='<%response.write rec.fields(8) %>' readonly></td>
            </tr>
            <tr></tr>
            </table>
       <%
        response.write"<br><br><br><br><h3><a href=admin_user.asp>Click here to Return on Administration Tools page</a></h3><br>"
        End If
        End If
      %>
<br><br>

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


