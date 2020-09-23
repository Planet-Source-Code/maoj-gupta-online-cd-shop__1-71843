<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'© 2000-2005 L.C. Enterprises
'http://LCen.com
%>
<%
	Option Explicit

	'Upload Class Step 1/4 ------------------------------
	Response.Expires = 0
	Server.ScriptTimeout = 10000		'Time-out limit for the file upload in seconds, increase if you expect and allow large files
%>
<!--#INCLUDE FILE="clsUploadRequest.asp" -->
<HTML>
<%

Response.Write "<BR>Upload started: " & Now() & "<BR><BR>"

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
		Response.Write("objUploadRequest.GetValue(""File1"") = " & objUploadRequest.GetValue("File1") & "<BR>")
	Else 'Upload limit exceeded
		Response.Write("<DIV ALIGN=""CENTER""><B>Sorry, your request cannot be completed because:<BR><BR>Maximum allowed size for the file is " & objUploadRequest.GetRoundedKB(objUploadRequest.UploadLimit) & " KB. Your request: " & objUploadRequest.GetRoundedKB(objUploadRequest.GetTotalBytes()) & " KB</DIV>")
		Response.Write("<META HTTP-EQUIV=""REFRESH"" CONTENT=""5; URL=javascript:history.go(-1)"">")

		'Clean up
		Set objUploadRequest = Nothing

		Response.End
	End If
End If

'All done
Response.Write "<BR><BR>Upload complete: " & Now()
Response.Write "<BR>Speed: " & objUploadRequest.GetUploadSpeed() & " KB/s"

'Upload Class Step 4/4 ------------------------------
'Clean up
Set objUploadRequest = Nothing

dim i
i=Upload.Form("name")
if i="" then
  response.write"<br>nothing"
else
response.write "<br>" & i
End if


%>
<!-- <META HTTP-EQUIV="REFRESH" CONTENT="5; URL=forwardurl.htm"> -->
</HTML>