<%@language=VBScript%>
<%Option Explicit%>

<html>
<head>
<title>IndianFunWorld</title>
<link href="style.css" rel="stylesheet" type="text/css" />
<!--[if gte IE 5.5]>
<script language="JavaScript" src="ie.js" type="text/JavaScript"></script>
<![endif]-->
<SCRIPT language=Javascript>
      <!--
      function isNumberKey(evt)
      {
         var charCode = (evt.which) ? evt.which : event.keyCode
         if (charCode > 31 && (charCode < 48 || charCode > 57))
            return false;

         return true;
      }
      function isChar(evt)
      {
         var charCode = (evt.which) ? evt.which : event.keyCode
         if ((charCode > 0 && charCode < 65 && charCode!=8 && charCode!=32) || (charCode > 90 && charCode < 97) || charCode > 122)
             return false;
            
         return true;
      }
      
       function isSpace(evt)
      {
         var charCode = (evt.which) ? evt.which : event.keyCode
         if (charCode == 32)
            return false;

         return true;
      }
      //-->
   </SCRIPT>

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
  <li><a href="setting.asp">Category +</a>
    <ul>
      <li><a href="setting.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="setting.asp">DVD+</a>
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
dim query
set conn=Server.CreateObject("ADODB.Connection")
set rec=Server.CreateObject("ADODB.RecordSet")
conn.Open="Provider=Microsoft.Jet.oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
query="Select * from [user] where uname='"& request.cookies("uname") &"'"
rec.open query,conn

%>
<form method=post action=update.asp>
<table cellspacing=10 >
<tr><td><strong>Name:-</strong></td>
<td><Input type=Text name=name maxlength=15 value="<%response.write rec.fields(2)%>" onkeypress="return isChar(event)"></td></tr>

<tr>
<td><strong>Username:-</strong></td>
<td><Input type=Text name=uname maxlength=15 value="<% response.write rec.fields(0) %>" readonly="readonly"></td></tr>
<tr><td></td><td><font color="#6f6f6f" size="-1" face="Arial, sans-serif"><strong>Example:manojmc2,raj_india etc.</strong></font></td></tr>
<tr></tr>

<tr>
<td><strong>Password:-</strong></td>
<td><input type=Password name=pass maxlength=15 value=""></td></tr>
<tr><td></td><td><font color="#6f6f6f" size="-1" face="Arial, sans-serif"><strong>Minimum of 6 characters in length.</strong></font></td></tr>

<tr>
<td><strong>Re-Type Password:-</strong></td>
<td><input type=Password name=repass maxlength=15 value=""></td>
</tr>

<tr>
<td><strong>Address:</strong></td>
<td><textarea cols=20 rows=5 name=address><%response.write rec.fields(3)%></textarea></td>
</tr>

<tr>
<td><strong>Contact Number:</strong></td>
<td><font size=3>+91</font>&nbsp;<input type=text name=ctno value="<%response.write rec.fields(4)%>" onkeypress="return isNumberKey(event)"></td>
</tr>

<tr>
<td><strong>Security Question:-</strong></td>
<td><select name=question>
   <option><%response.write rec.fields(5)%></option> 
  <option value="What is your Pet name">What is your Pet name</option>
  <option value="What is your library card number">What is your library card number</option>
  <option value="What was your first phone number">What was your first phone number</option>
  <option value="What was your first teacher's name">What was your first teacher's name</option>
</select></td></tr>

<tr>
<td><strong>Answer:-</strong></td>
<td><input type=text name=ans value="<%response.write rec.fields(6)%>"></td></tr>

<tr>
<td><strong>Email Address:-</strong></td>
<td><input type=text name=email value="<%response.write rec.fields(7)%>"></td></tr>

<tr>
<td><strong>State:-</strong></td>
<td><select name=state>
<option><%response.write rec.fields(8)%></option>
<option>Maharastra</option>
<option>Delhi</option>
<option>Panjab</option>
<option>Jammu &amp Kasmir</option>
<option>Uttar Pradesh</option>
<option>Bihar</option>
<option>Madhya Pradesh</option>
<option>Kalkatta</option>
<option>Bengal</option>
<option>Chennai</option>
<option>Tamilnadu</option>
<option>Mizoram</option>
<option>Sikkim</option>
</select></td>
</tr>
<tr></tr>
<tr>
<td><input type=Submit value="Update My Account" name=action></td>
<td><input type=Submit value="Cancel" name=action>&nbsp;&nbsp;&nbsp;&nbsp;
<input type=Submit value="Delete Account" name=action></td>
</tr>
</table>
</form>
<%
set rec=nothing
set conn=nothing
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


