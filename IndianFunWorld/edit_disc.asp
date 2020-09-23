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
  <li><a href="edit_disc.asp">Category +</a>
    <ul>
      <li><a href="edit_disc.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="edit_disc.asp">DVD+</a>
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
         dim conn,rec
         dim query,action
         dim id
         action=request.form("action")
         set conn=Server.CreateObject("ADODB.Connection")
         set rec=Server.CreateObject("ADODB.RecordSet")
         conn.Open="Provider=Microsoft.Jet.oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"
         if(action="Add Disc" or action="Edit") Then
            if(action="Add Disc")then
                response.write"<form method=Post action=add_Disc.asp ENCTYPE='multipart/form-data'>"
                response.write"<table>"
                response.write"<tr>"
                response.write"<td><h3>Title:-</h3></td>"
                response.write"<td><input type=text name=title></td>"
                response.write"</tr>"
                response.write"<tr>"
                response.write"<td><h3>Image Path:-</h3></td>"
                response.write"<td><input type=File name=image></td>"
                response.write"</tr>"
                response.write"<tr>"
                response.write"<td><h3>Price:-</h3></td>"
                response.write"<td><input type=Text name=price></td>"
                response.write"</tr>"
                response.write"<tr>"
                response.write"<td><h3>Information:-</h3></td>"
                response.write"<td><input type=text name=info></td>"
                response.write"</tr>"
                response.write"<tr>"
                response.write"<td><h3>Category:-</h3></td>"
                response.write"<td><select name=category>"
                response.write"<option>Application</option>"
                response.write"<option>Movie</option>"
                response.write"<option>Game</option>"
                response.write"<option>Music</option>"
                response.write"</select></td>"
                response.write"</tr>"
                response.write"<tr>"
                response.write"<td><h3>Disc Type</h3></td>"
                response.write"<td><select name=t>"
                response.write"<option>CD</option>"
                response.write"<option>DVD</option>"
                response.write"</select></td>"
                response.write"</tr>"
                response.write"<tr>"
                response.write"<td><h3>Quantity</h3></td>"
                response.write"<td><input type=text name=qty onkeypress='return isNumberKey(event)'></td>"
                response.write"</tr>"
                response.write"<tr><tr>"
                response.write"<tr>"
                response.write"<td>&nbsp;&nbsp;<input type=Submit value='Add Disc' name=action></td>"
                response.write"<td><input type=Reset value=Clear></td>"
                response.write"</tr>"
                response.write"</table>"
                response.write"</form>"
            else
                id=request.form("id")    
                if not(id="")then   
                      query="Select * from [Disc] where ID="& id
                      rec.open query,conn     
                      response.write"<form method=post action=add_Disc.asp ENCTYPE='multipart/form-data'>"
                      response.write"<table>"
                      response.write"<input type=hidden name=id value="& id &">"
                      response.write"<tr>"
                      response.write"<td><h3>Title:-</h3></td>"
                      response.write"<td><input type=text name=title value='"& rec.fields(1) &"'></td>"
                      response.write"</tr>"
                      response.write"<tr>"
                      response.write"<td><h3>Image Path:-</h3></td>"
                      response.write"<td><input type=text name=image value='"& rec.fields(2) &"'></td>"
                      response.write"</tr>"
                      response.write"<tr>"
                      response.write"<td><h3>Price:-</h3></td>"
                      response.write"<td><input type=text name=price value='"& rec.fields(3) &"'></td>"
                      response.write"</tr>"
                      response.write"<tr>"
                      response.write"<td><h3>Information:-</h3></td>"
                      response.write"<td><input type=text name=info value='"& rec.fields(4) &"'></td>"
                      response.write"</tr>"
                      response.write"<tr>"
                      response.write"<td><h3>Category:-</h3></td>"
                      response.write"<td><select name=category>"
                      response.write"<option>"& rec.fields(5) &"</option>"
                      response.write"<option>Application</option>"
                      response.write"<option>Movie</option>"
                      response.write"<option>Game</option>"
                      response.write"<option>Music</option>"
                      response.write"</select></td>"
                      response.write"</tr>"
                      response.write"<tr>"
                      response.write"<td><h3>Disc Type</h3></td>"
                      response.write"<td><select name=t>"
                      response.write"<option>"& rec.fields(6) &"</option>"
                      response.write"<option>CD</option>"
                      response.write"<option>DVD</option>"
                      response.write"</select></td>"
                      response.write"</tr>"
                      response.write"<tr>"
                      response.write"<td><h3>Quantity</h3></td>"
                      response.write"<td><input type=text name=qty value="& rec.fields(7) &"></td>"
                      response.write"</tr>"
                      response.write"<tr><tr>"
                      response.write"<tr>"
                      response.write"<td>&nbsp;&nbsp;<input type=Submit value='Update' name=action></td>"
                      response.write"<td><input type=Reset value='Clear'></td>"
                      response.write"</tr>"
                      response.write"</table>"
                      response.write"</form>"
                 else
                      response.write"<h3>Please Select Any Disc First<br><br><a href=admin_cd.asp>Click Here to return on Admin Disc Edit Page</a></h3>"
                 End If
            End If 
         elseif(action="Delete") then
            id=request.form("id")       
            if not id="" then
                conn.Execute"Delete * From [Disc] where ID="& id
                response.write"<br><h3>Disc Delete Successfully<br><br><a href=admin_cd.asp>Click Here to return on Disc Edit page</a></h3>"
            else
                response.write"<h3>Please Select Any Disc First<br><br><a href=admin_cd.asp>Click Here to return on Admin Disc Edit Page</a></h3>"
            End If
         End If
      %>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
</div>
    <!-- end inner -->
    </div>
    <!-- end body -->
    <div class="clear"></div>
    <div id="footer"> &nbsp;</div>
    <!-- end footer -->
  </div>
  <!-- end inner -->
</div>
<!-- end wrapper -->

</body>
</html>
