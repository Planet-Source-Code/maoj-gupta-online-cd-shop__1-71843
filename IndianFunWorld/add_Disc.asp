<%@language=VBScript%>
<%Option Explicit%>
<!--#INCLUDE FILE="clsUploadRequest.asp" -->
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
  <li><a href="add_Disc.asp">Category +</a>
    <ul>
      <li><a href="add_Disc.asp">CD+</a>
        <ul>
          <li><a href="cd_movie.asp">Movie</a></li>
          <li><a href="cd_app.asp">Application</a></li>
          <li><a href="cd_game.asp">Games</a></li>
          <li><a href="cd_music.asp">Music</a></li>
        </ul>
      </li>
      <li><a href="add_Disc.asp">DVD+</a>
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
          dim query
          dim title,image,price
          dim info,category,t
          dim qty,id,action
          set conn=Server.CreateObject("ADODB.Connection")
          set rec=Server.CreateObject("ADODB.RecordSet")
          conn.Open="Provider=Microsoft.Jet.oledb.4.0;Data Source=C:\Inetpub\wwwroot\IndianFunWorld\db1.mdb"

          'Uploading Files
          	'Upload Class Step 1/4 ------------------------------
                Response.Expires = 0
                Server.ScriptTimeout = 10000		'Time-out limit for the file upload in seconds, increase if you expect and allow large files

              'Upload Class Step 2/4 ------------------------------
              Dim objUploadRequest : Set objUploadRequest = New UploadRequest

              'Initialize and configurate
              With objUploadRequest
                .UploadLimit = 30000000				'File size limit in Bytes
                .UploadFolder = "Images/"
                .AllowedExtensions = "jpg gif" 		'Very important to avoid hijacking, otherwise an executable file could be uploaded
                .Verbose = 1						'Display info about uploaded file(s): 0 = No / 1 = Yes
                .RandomName = 0						'Generate a unique name for uploaded file, extension is kept: 0 = No / 1 = Yes
                .UploadPrefix = ""					'Prefix for value names of uploaded files, example: "Uploaded/"
              End With

              'Upload Class Step 3/4 ------------------------------
              If objUploadRequest.GetTotalBytes() > 0 Then 'Data has been received
                'Process upload request (save files and store POST variables)
                If objUploadRequest.Process() Then
                  'Show request value, use instead of Request("File1") or Request.Form("File1")
                Else 'Upload limit exceeded
                  Response.Write("<DIV ALIGN=""CENTER""><B>Sorry, your request cannot be completed because:<BR><BR>Maximum allowed size for the file is " & objUploadRequest.GetRoundedKB(objUploadRequest.UploadLimit) & " KB. Your request: " & objUploadRequest.GetRoundedKB(objUploadRequest.GetTotalBytes()) & " KB</DIV>")
                  Response.Write("<META HTTP-EQUIV=""REFRESH"" CONTENT=""5; URL=javascript:history.go(-1)"">")

                  'Clean up
                  Set objUploadRequest = Nothing

                  Response.End
                End If
              End If
          
          'Inserting Into Database
          title=trim(objUploadRequest.GetValue("title"))
          image="Images\" & trim(objUploadRequest.GetValue("image"))
          price=trim(objUploadRequest.GetValue("price"))
          info=trim(objUploadRequest.GetValue("info"))
          category=trim(objUploadRequest.GetValue("category"))
          t=trim(objUploadRequest.GetValue("t"))
          qty=trim(objUploadRequest.GetValue("qty"))
          id=objUploadRequest.GetValue("id")
          action=objUploadRequest.GetValue("action")
          
          if(action="Update") Then
              query="UPDATE Disc SET Disc.[Title] = '"& title &"', Disc.[Image] = '"& image &"', Disc.Price = '"& price &"', Disc.Info = '"& info &"', Disc.Category = '"& category &"', Disc.Type = '"& t &"', Disc.Quantity = "& qty &" where ID="& id &" ;"
              conn.Execute(query)
              response.write"<br><h3>Disc Updated Successfully<br><br><a href=admin.asp>Click Here to return on Administrative tool</a></h3>"
          Else
             if(title="" or image="" or price="" or category="" or t="" or qty="") then
                 response.write"<br><h3>Please Fill all fields<br><br><a href=javascript:history.back(1);>Click here to return on admin Tools</a></h3>" 
             else
                   id=1 
                   qty=cint(qty) 
                   rec.open"Select * from [Disc]",conn
                   do while not rec.EOF
                      id=rec.fields(0) 
                      rec.movenext
                   loop
                   id=id+1
                  query="Insert into [Disc] values("& id &",'"& title &"','"& image &"','"& price &"','"& info &"','"& category &"','"& t &"',"& qty &")"
                  conn.Execute(query)
                  response.write"<h3>Disc Added Successfully<br><br><a href=admin.asp>Click Here to return on Administrative tool</a></h3>"
              End If
          End If       
          
          'Upload Class Step 4/4 ------------------------------
          'Clean up
          Set objUploadRequest = Nothing
          %> 

      
 <br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
      
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
