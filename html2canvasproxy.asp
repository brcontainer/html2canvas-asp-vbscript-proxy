<%@ Language=vbScript Debug=true EnableSessionState=false %><%
	'proxy.js 0.0.5
	'Copyright (c) 2014 Guilherme Nascimento (brcontainer@yahoo.com.br)
	'
	'Released under the MIT license

	'usage variables
	Dim serr, xmlHttp, url, callback, absolutePath, tmpName, fileName, mime, oas, objFSO, PATH, CCACHE, SECPREFIX

	'Setup
	PATH = "images"
	SECPREFIX = "h2c_"
	CCACHE = 60 * 5 * 1000 'Limit access-control, cache-control, delete old files

	Function FULL_URL()
		Dim a, b, c, d
		a = ""
		c = Request.ServerVariables("SERVER_PORT")
		d = Request.ServerVariables("HTTP_HOST")

		If c = "443" Then
			a = a & "https://"
		Else
			a = a & "http://"
		End If
		a = a & d

		If c <> "80" AND c <> "443" AND INSTR(d, ":") = 0 Then
			a = a & ":" & c
		End If
		
		b = Split(Request.ServerVariables("URL"), "/")
		b(Ubound(b)) = ""

		FULL_URL = a & Join(b, "/")
	End Function

	Function ERROR_HANDLE(s)
		Response.Write callback & "(""error:" & JSENCODE(s) & """);"
		Response.End()
	End Function

	Function DELETE_OLD_FILES(a)
		Dim folder, files, folderIdx, timer, tmpTimer

		Set folder = objFSO.GetFolder(a)
		Set files = folder.Files

		timer = DateDiff("s", "01/01/1970 00:00:00", Now())

		If files.count <> 0 Then
			For each folderIdx In files
				tmpTimer = DateDiff("s", "01/01/1970 00:00:00", folderIdx.DateLastModified)
				IF (timer - tmpTimer) > (CCACHE * 2) Then
					If objFSO.Fileexists(a & "\" & folderIdx.Name) Then
						objFSO.DeleteFile a & "\" & folderIdx.Name
					End If
				End If
			Next
		End If
	End Function

	Function FILE_EXISTS(s, e)
		If s = 0 Then
			FILE_EXISTS = FILE_EXISTS(DateDiff("s", "01/01/1970 00:00:00", Now()), e)
		Else
			If objFSO.Fileexists(SECPREFIX & s & "." & e) Then
				FILE_EXISTS = FILE_EXISTS(s + 1, e)
			Else
				FILE_EXISTS = SECPREFIX & s & "." & e
			End If
		End If
	End Function

    Function EXTRACT_URL(url)
		Dim a(1), b(1), auth

		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(^[A-Z0-9]+[:]\/\/)+([A-Z0-9_\-:]+|)[@]([\s\S]+)$"

		a(0) = re.Replace(url, "$1$3")

		auth = re.Replace(url, "$2")

		If auth <> "" Then
			If INSTR(auth, ":") <> 0 Then
				a(1) = Split(auth, ":")
			Else
				b(0) = auth
				b(1) = ""
				a(1) = b
			End If
		Else
			a(1) = Null
		End If

		EXTRACT_URL = a
    End Function

	Function JSENCODE(s)
		'Based in VBS JSON 2.0.3, Copyright (c) 2009 Tuðrul Topuz, Under the MIT license.

		Dim a(127), b()
		a(8)  = "\b"
		a(9)  = "\t"
		a(10) = "\n"
		a(12) = "\f"
		a(13) = "\r"
		a(34) = "\"""
		a(47) = "\/"
		a(92) = "\\"

		Dim z : z = Len(s) - 1
		ReDim b(z)

		Dim i, c
		For i = 0 To z
			b(i) = Mid(s, i + 1, 1)

			c = AscW(b(i))
			If c < 127 Then
				If Not IsEmpty(a(c)) Then
					b(i) = a(c)
				ElseIf c < 32 Then
					b(i) = "\u" & Right("000" & Hex(c), 4)
				End If
			Else
				b(i) = "\u" & Right("000" & Hex(c), 4)
			End If
		Next

		JSENCODE = Join(b, "")
	End Function

	absolutePath = Replace(Server.MapPath("./"), "\.", "")

	Response.AddHeader "Access-Control-Max-Age", CCACHE
	Response.AddHeader "Access-Control-Allow-Origin", "*"
	Response.AddHeader "Access-Control-Request-Method", "*"
	Response.AddHeader "Access-Control-Allow-Methods", "OPTIONS, GET"
	Response.AddHeader "Access-Control-Allow-Headers", "*"

	Response.ContentType = "application/javascript"

	url = Request.QueryString("url")
	callback = Request.QueryString("callback")

	If url <> "" AND callback <> "" Then
		On Error Resume Next

			set filesys = CreateObject("Scripting.FileSystemObject")
			If Not filesys.FolderExists(absolutePath & "\" & PATH) Then
				filesys.CreateFolder(absolutePath & "\" & PATH)
			End If

			If Not filesys.FolderExists(absolutePath & "\" & PATH) Then
				ERROR_HANDLE("Failed to create the folder " & absolutePath & "\" & PATH)
			End If

			set xmlHttp = server.Createobject("MSXML2.ServerXMLHttp")

			eUrl = EXTRACT_URL(url)

			If IsNull(eUrl(1)) Then
				xmlHttp.Open "GET", eUrl(0), false
			Else
				xmlHttp.Open "GET", eUrl(0), false, eUrl(1)(0), eUrl(1)(1)
			End If

			xmlHttp.setRequestHeader "User-Agent", Request.ServerVariables("HTTP_USER_AGENT")
			xmlHttp.send ""

			If err = 0 Then
				If xmlHttp.Status = 200 Then
					tmpName = absolutePath & "\" & PATH

					Set oas = CreateObject("ADODB.Stream")
					oas.Open
					oas.Type = 1 'adTypeBinary

					oas.Write xmlHttp.ResponseBody
					oas.Position = 0 'Set the stream position to the start

					Set objFSO = Createobject("Scripting.FileSystemObject")

					DELETE_OLD_FILES(absolutePath & "\" & PATH)

					mime = Trim(xmlHttp.getAllResponseHeaders)
					mime = Replace(Replace(LCASE(mime), CHR(13), CHR(10)), CHR(10), "|")

					Dim counter, myArray
					myArray = Split(mime, "|")
					mime = ""

					For counter = 0 To UBound(myArray)
						If INSTR(myArray(counter), "content-type:") = 1 Then
							mime = Trim(Replace(Replace(myArray(counter), "content-type:", ""), "/x-", "/"))
						End If
					Next

					If mime = "" Then
						serr = "No such mime-type"
					ElseIf INSTR("|image/jpeg|image/jpg|image/png|image/gif|text/html|application/xhtml|application/xhtml+xml|", "|" & mime & "|") = 0 Then
						serr = "Invalid mime-type"
					Else
						mime = Replace(Replace(Replace(mime, "image/", ""), "text/", ""), "application/", "")
						mime = Replace(mime, "xhtml+xml", "xhtml")

						fileName = FILE_EXISTS(0,mime)
						tmpName = tmpName & "\" & fileName

						'save responseBody to path
						oas.SaveToFile tmpName
						oas.Close
						Set oas = Nothing

						If objFSO.Fileexists(tmpName) Then
							Response.AddHeader "Cache-control", "public, max-age = " & CCACHE
							Response.AddHeader "Pragma", "max-age = " & CCACHE
							Response.Expires = Round(CCACHE/60)

							Set objFSO = Nothing
							Response.Write callback & "(""" & JSENCODE(FULL_URL() & PATH & "/" & fileName) & """);"
							Response.End()
						Else
							serr = "Can not create file " & tmpName
						End If
					End If

					Set objFSO = Nothing
				Else
					serr = "http error: " & xmlHttp.status
				End If
			End If
		On Error GoTo 0
	ElseIf url = "" Then
		serr = "url variable is undefined"
	ElseIf callback = "" Then
		serr = "callback variable is undefined"
	End If

	If err <> 0 Then
		serr = err.Description
	ElseIf serr = "" Then
		serr = "unknown error, maybe the server from url is not available"
	End If

	If serr <> "" Then
		Response.AddHeader "Pragma", "no-cache"
		Response.AddHeader "Cache-control", "no-cache"
		Response.Expires = 0

		ERROR_HANDLE(serr)
	End If

	set xmlHttp = nothing
%>
