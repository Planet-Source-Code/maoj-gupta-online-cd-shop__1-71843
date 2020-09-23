<%
'© 2000-2005 L.C. Enterprises
'http://LCen.com
%>
<%
'Upload Request Class
'Author: Luis Cantero
'Updated: 30/DEC/2005

'INITIALIZE:	Set objUploadRequest = New UploadRequest, .UploadLimit, .UploadFolder, .AllowedExtensions, .Verbose, .RandomName, .UploadPrefix
'PROPERTIES:	.UploadLimit, .ValueCount
'METHODS:		.GetRoundedKB(lngByteAmount)
'FUNCTIONS:		.Process(), .GetValue(strName), .GetTotalBytes(), .GetUploadSpeed()
Class UploadRequest '-------------------------- CLASS BEGIN --------------------------

Private lngUploadLimit	'File size limit in Bytes
Private strUploadFolder
Private arrAllowedExtensions
Private intVerbose
Private intRandomName 'Generate random name for uploaded files
Private strUploadPrefix

Private lngBytesReceived, arrMimeData, lngItemIndex
Private strActualFileOrValue, strSaveName, strName
Private dicNamesValues, datDate

Public Property Let UploadLimit(lngInput)
	lngUploadLimit = lngInput
End Property

Public Property Let UploadFolder(strInput)
	strUploadFolder = strInput
End Property

'Extensions in string separated by spaces, case insensitive, example "jpg gif"
Public Property Let AllowedExtensions(strInput)
	arrAllowedExtensions = Split(strInput)
End Property

Public Property Let Verbose(intInput)
	intVerbose = intInput
End Property

Public Property Let RandomName(intInput)
	intRandomName = intInput
End Property

Public Property Let UploadPrefix(strInput)
	strUploadPrefix = strInput
End Property

Public Property Get UploadLimit
	UploadLimit = lngUploadLimit
End Property

Public Property Get ValueCount
	ValueCount = dicNamesValues.Count
End Property

'OUTPUT:	Boolean: (True) Upload within limits and processed ok | (False) Upload over limit
Public Function Process() '--------------------------
	'Upload limit check
	If lngBytesReceived > lngUploadLimit Then
		Process = False
		Exit Function
	End If

	'Put files into a String Array
	arrMimeData = GetMimeArray()

	'Loop for every file in the Array
	For lngItemIndex = 0 To UBound(arrMimeData)
		'Get the name of the variable
		strName = GetMimeName(arrMimeData(lngItemIndex))
		'Get file or value
		strActualFileOrValue = GetMimeValue(arrMimeData(lngItemIndex))

		If GetMimeContentType(arrMimeData(lngItemIndex)) = "" Or GetMimeFilename(arrMimeData(lngItemIndex)) = "" Then 'No Content-type Or filename found
			'Write form name and value
			Call dicNamesValues.Add(strName, strActualFileOrValue)
		Else 'Content-type and filename found
			strSaveName = GetFilenameFromPath(GetMimeFilename(arrMimeData(lngItemIndex))) 'Use original filename
			strSaveName = Replace(strSaveName, ",", "") 'Remove , to avoid MapPath errors

			'Generate unique name, extension is kept
			If intRandomName Then strSaveName = GenerateRandomName(strSaveName)

			'Compare to allowed extensions, if not allowed append first allowed extension
			strSaveName = GetAllowedExtension(strSaveName)

			Call dicNamesValues.Add(strName, strUploadPrefix & strSaveName)

			Call WriteFile(strActualFileOrValue, Server.MapPath(strUploadFolder & strSaveName))
			If intVerbose = 1 Then Response.Write("<FONT COLOR=""White"">" & strSaveName & " (" & GetRoundedKB(Len(strActualFileOrValue)) & " KB)<BR></FONT>")
		End If
	Next

	'Return
	Process = True
End Function

Private Sub Class_Initialize() '--------------------------
	datDate = Now()
	Set dicNamesValues = CreateObject("Scripting.Dictionary")
	dicNamesValues.CompareMode = vbTextCompare 'Case insensitive
	lngBytesReceived = Request.TotalBytes
End Sub

Private Sub Class_Terminate() '--------------------------
	'Clean up
	Set arrMimeData = Nothing
	Set dicNamesValues = Nothing
End Sub

Public Function GetValue(strName) '--------------------------
	'Return
	GetValue = dicNamesValues(strName)
End Function

Public Function GetTotalBytes() '--------------------------

	'Return
	GetTotalBytes = lngBytesReceived

End Function

Public Function GetUploadSpeed() '--------------------------

	'Return
	GetUploadSpeed = GetRoundedKBperS(lngBytesReceived, DateDiff("s", datDate, Now()))

End Function

'PURPOSE:	Returns the HEADER of a Mime entry
'INPUT:		String: Mime entry
'OUTPUT:	String: Found header | Empty
Private Function GetMimeHeader(strMime) '--------------------------

	Dim intDataStart, strHeader

	'Find header boundary
	intDataStart = InStr(strMime, vbCrLf & vbCrLf)
	If intDataStart > 0 Then 'Header boundary found
		'Parse header and return
		GetMimeHeader = Left(strMime, intDataStart - 1)
	End If

End Function

'PURPOSE:	Returns an array containing the received multipart data in Mime format
'OUTPUT:	String Array | Empty
Private Function GetMimeArray() '--------------------------

	Dim strSignature, strCompleteData, lngTotalBytes, lngBytesRead, lngChunkSize, objSourceData

	'Initialize
	lngChunkSize = 5242880 '5 MB
	lngBytesRead = 0

	If Request.ServerVariables("REQUEST_METHOD") = "POST" And LCase(Left(Request.ServerVariables("HTTP_Content_Type"), 19)) = "multipart/form-data" Then
		lngTotalBytes = Request.TotalBytes

		Set objSourceData = CreateObject("ADODB.Stream")
		objSourceData.Open
		objSourceData.Type = 1 'Binary

		Do While lngBytesRead < lngTotalBytes And Response.IsClientConnected
			'Adjust chunk's length before reading
			If lngChunkSize + lngBytesRead > lngTotalBytes Then lngChunkSize = lngTotalBytes - lngBytesRead

			'Read chunk of data
			objSourceData.Write(Request.BinaryRead(lngChunkSize))

			'Increase read counter
			lngBytesRead = lngBytesRead + lngChunkSize
		Loop

		'Convert binary string to real string
		objSourceData.Position = 0
		strCompleteData = GetStringFromBinary(objSourceData.Read)

		Set objSourceData = Nothing

		'Parse Signature (file separator)
		strSignature = Left(strCompleteData, InStr(strCompleteData, vbCrLf) + 1)

		strCompleteData = Mid(strCompleteData, InStr(strCompleteData, vbCrLf) + 2, Len(strCompleteData) - 2 - 2 * Len(strSignature))

		'Put files into a string Array and return
		GetMimeArray = Split(strCompleteData, strSignature)
	Else 'Return empty array
		GetMimeArray = Split("")
	End If

End Function

'PURPOSE:	Converts a multibyte or binary string (VT_UI1 | VT_ARRAY) to a real string (BSTR) using an ADO Recordset
'INPUT:		String: MultiByte or Binary string
'OUTPUT:	String: Real string
Private Function GetStringFromBinary(strBinaryData) '--------------------------

	Dim Rs, lngBinaryLength
	Const adLongVarChar = 201

	lngBinaryLength = LenB(strBinaryData)

	'Error check
	If lngBinaryLength = 0 Then Exit Function

	'MultiByte data must be converted To VT_UI1 | VT_ARRAY first
	If VarType(strBinaryData) = 8 Then strBinaryData = MultiByteToBinary(strBinaryData)

	Set Rs = CreateObject("ADODB.Recordset")

	Rs.Fields.Append "tmpBinField", adLongVarChar, lngBinaryLength 'Create temp field
	Rs.Open
	Rs.AddNew
	Rs("tmpBinField").AppendChunk strBinaryData	'Add binary data to temp table
	Rs.Update

	GetStringFromBinary = Rs("tmpBinField") 'Get string and return it

	Set Rs = Nothing

End Function

'PURPOSE:	Converts a multibyte string to real binary data (VT_UI1 | VT_ARRAY) using an ADO Recordset
'INPUT:		String: MultiByte string
'OUTPUT:	String: Real binary string
Private Function GetBinaryFromMultiByte(strMultiByte) '--------------------------

	Dim Rs, lngMultiByteLength
	Const adLongVarBinary = 205

	lngMultiByteLength = LenB(strMultiByte)

	'Error check
	If lngMultiByteLength = 0 Then Exit Function

	Set Rs = CreateObject("ADODB.Recordset")

	Rs.Fields.Append "tmpBinField", adLongVarBinary, lngMultiByteLength 'Create temp field
	Rs.Open
	Rs.AddNew
	Rs("tmpBinField").AppendChunk strMultiByte & ChrB(0) 'Add multibyte data to temp table
	Rs.Update

	GetBinaryFromMultiByte = Rs("tmpBinField").GetChunk(lngMultiByteLength) 'Get binary data and return it

	Set Rs = Nothing

End Function

'PURPOSE:	Returns the "Content-Disposition" of a Mime entry
'INPUT:		String: Mime entry
'OUTPUT:	String: Found content-disposition | Empty
Private Function GetMimeContentDisposition(strMime) '--------------------------

	GetMimeContentDisposition = GetMimeValueByCoord(strMime, "Content-Disposition:", ";")

End Function

'PURPOSE:	Returns the "Content-Type" of a Mime entry
'INPUT:		String: Mime entry
'OUTPUT:	String: Found content-type | Empty
Private Function GetMimeContentType(strMime) '--------------------------

	GetMimeContentType = GetMimeValueByCoord(strMime, "Content-Type:", vbCrLf)

End Function

'PURPOSE:	Returns the "filename" of a Mime entry
'INPUT:		String: Mime entry
'OUTPUT:	String: Found filename | Empty
Private Function GetMimeFilename(strMime) '--------------------------

	GetMimeFilename = GetMimeValueByCoord(strMime, "filename=""", """")

End Function

'PURPOSE:	Returns the "name" of a Mime entry
'INPUT:		String: Mime entry
'OUTPUT:	String: Found name | Empty
Private Function GetMimeName(strMime) '--------------------------

	GetMimeName = GetMimeValueByCoord(strMime, "name=""", """")

End Function

'PURPOSE:	Returns the value of an ENTRY of a Mime entry
'INPUT:		String: Mime entry
'OUTPUT:	String: Found value | Empty
Private Function GetMimeValue(strMime) '--------------------------

	Dim intBeg

	'Search for left boundary
	intBeg = InStr(1, strMime, vbCrLf & vbCrLf) + Len(vbCrLf & vbCrLf)

	If intBeg > 0 Then
		'Return found value
		GetMimeValue = Mid(strMime, intBeg, Len(strMime) - intBeg - 1)
	End If

End Function

'PURPOSE:	Returns the value of a VARIABLE of a Mime entry
'INPUT:		String: Mime entry, String: Left and right bounds
'OUTPUT:	String: Found value | Empty
Private Function GetMimeValueByCoord(strMime, strLeftBound, strRightBound) '--------------------------

	Dim intBeg, intEnd, strHeader

	'Get header
	strHeader = GetMimeHeader(strMime)

	'Search for value name in header
	intBeg = InStr(1, strHeader, strLeftBound)

	If intBeg > 0 Then	 'Value name found in header, parse value
		intBeg = intBeg + Len(strLeftBound)
		intEnd = InStr(intBeg, strMime, strRightBound)

		If intEnd = 0 Then intEnd = Len(strMime)

		'Return found value
		GetMimeValueByCoord = Trim(Mid(strMime, intBeg, intEnd - intBeg))
	End If

End Function

'PURPOSE:	Returns the value of an ENTRY in the Mime array, according to it's NAME
'INPUT:		Array: Mime array, String: Name of the value to be retrieved
'OUTPUT:	String: Found value | Empty
Private Function GetMimeValueByName(arrMime, strValueName) '--------------------------

	Dim intI, intBeg, strHeader

	'Search all items of array
	For intI = 0 To UBound(arrMime)

		'Get header
		strHeader = GetMimeHeader(arrMime(intI))
		If strHeader <> "" Then 'Header found
			'Search for value name in header
			intBeg = InStr(1, strHeader, "name=""" & strValueName & """")
			If intBeg > 0 Then 'Value name found in header, parse value
				intBeg = intBeg + Len("name=""" & strValueName & """") + Len(vbCrLf & vbCrLf)

				'Return found value
				GetMimeValueByName = Mid(arrMime(intI), intBeg, Len(arrMime(intI)) - intBeg - 1)
				Exit For
			End If
		End If
	Next

End Function

'PURPOSE:	Extracts the name of a file from its path
'INPUT:		String: Path
'OUTPUT:	String: Name
Private Function GetFilenameFromPath(strPath) '--------------------------

	Dim intI

	intI = InStrRev(strPath, "\")
	If intI > 0 Then
		GetFilenameFromPath = Mid(strPath, intI + 1)
	Else
		GetFilenameFromPath = strPath
	End If

End Function

'PURPOSE:	Saves a file using the FileSystemObject
'INPUT:		String: files's contents, String: save path
Private Sub WriteFile(strFileData, strSavePath) '--------------------------

	Dim objFSO, objTextStream

	If Len(strFileData) = 0 Then Exit Sub

	'Create objects
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.CreateTextFile(strSavePath, True, False)

	'Write file and close it
	objTextStream.Write strFileData
	objTextStream.Close

	'Delete objects
	Set objTextStream = Nothing
	Set objFSO = Nothing

End Sub

'PURPOSE:	Deletes a file using the FileSystemObject
'INPUT:		String: files's path
Private Sub DeleteFile(strPath) '--------------------------

	Dim objFSO, tmpFileHandle

	'Create object
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'Delete zoom image
	If objFSO.FileExists(strPath) Then
		Set tmpFileHandle = objFSO.GetFile(strPath)
	 	tmpFileHandle.Delete
	End If

	'Delete objects
	Set tmpFileHandle = Nothing
	Set objFSO = Nothing

End Sub

'PURPOSE:	Checks if the passed filename's extension is allowed, if not it appends the first allowed extension
'INPUT:		String: Filename
'OUTPUT:	String: Filename with an allowed extension
Private Function GetAllowedExtension(strInput)

	Dim i
	Dim strFileExtension, intTemp

	If InStr(strInput, ".") > 0 Then
		strFileExtension = Mid(strInput, InStrRev(strInput, ".") + 1)

		For i = 0 To UBound(arrAllowedExtensions)
			If StrComp(strFileExtension, arrAllowedExtensions(i), vbTextCompare) = 0 Then
				'Filename has allowed extension, case insensitive
				GetAllowedExtension = strInput
				Exit Function
			End If
		Next

		'Return filename with first allowed extension appended
		GetAllowedExtension = strInput & "." & arrAllowedExtensions(0)
	Else 'No extension
		'Return
		GetAllowedExtension = strInput
	End If

End Function

'PURPOSE:	Generates a "Unique Random Filename" to save uploaded files
'INPUT:		String: Name of uploaded file
'OUTPUT:	String: Unique random filename in this format: YYYY_MM_DD_RandomNumber.EXT
'NOTES:		RandomNumber = Round(RND * 1000000), Extension is kept
Private Function GenerateRandomName(strFileName) '--------------------------

	Dim strDate

	'Remove illegal characters for a filename
	strDate = Year(Now()) & "_" & String(2 - Len(Month(Now())), "0") & Month(Now()) & "_" & String(2 - Len(Day(Now())), "0") & Day(Now())

	'Return
	GenerateRandomName = strDate & "_" & CStr(GetRandomNumber(1, 1000000)) & Mid(strFileName, InStrRev(strFileName, ".")) 'YYYY_MM_DD_RandomNumber.EXT

End Function

'PURPOSE:	Generates a random number whithin the specified upper and lower bounds
'INPUT:		Integer: Upper and lower bounds
'OUTPUT:	Integer: Random number
Private Function GetRandomNumber(intLBound, intUBound) '--------------------------

	Randomize Timer / Rnd()

	'Return
	GetRandomNumber = Int((intUBound - intLBound + 1) * Rnd() + intLBound)

End Function

'PURPOSE:	Rounds a Byte amount and returns KB with 2 decimal places
'INPUT:		Long: Byte amount
'OUTPUT:	String: Rounded KB amount
Public Function GetRoundedKB(lngByteAmount) '--------------------------

	GetRoundedKB = FormatNumber(Int(lngByteAmount / 1024 * 100 + 0.5) / 100, 2)

End Function

'PURPOSE:	Rounds a Byte amount and returns, according to an elapsed time in seconds, KB/s with 2 decimal places
'INPUT:		Long: Byte amount
'OUTPUT:	String: Rounded KB/s amount
Private Function GetRoundedKBperS(lngByteAmount, lngSecondsElapsed) '--------------------------

	'Error check
	If lngSecondsElapsed <= 0 Then lngSecondsElapsed = 1

	GetRoundedKBperS = FormatNumber(Int(lngByteAmount / 1024 / lngSecondsElapsed * 100 + 0.5) / 100, 2)

End Function

End Class '-------------------------- CLASS END --------------------------
%>