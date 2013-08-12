<%@ Language=vbScript Debug=true EnableSessionState=false %><%
	'html2canvas-csharp-proxy 0.0.1
	'Copyright (c) 2013 Guilherme Nascimento (brcontainer@yahoo.com.br)
	'
	'Released under the MIT license

	'usage variables
	Dim serr, xmlHttp, url, callback, absolutePath, tmpName, fileName, oas, objFSO

	'setup variables
	Dim PATH, CCACHE

	'Setup
	PATH = "images"
	CCACHE = 60 * 5 * 1000 'Limit access-control, cache-control, delete old files

	Function FULL_URL()
		Dim a, b
		a = ""
		If Request.ServerVariables("SERVER_PORT")="443" Then
			a = a & "https://"
		Else
			a = a & "http://"
		End If
		a = a & Request.ServerVariables("REMOTE_HOST")
		
		b = Split(Request.ServerVariables("URL"),"/")
		b(Ubound(b))=""

		FULL_URL = a & Join(b,"/")
	End Function

	Function ERROR_HANDLE(s)
		Response.Write callback & "(""error:" & JSENCODE(s) & """);"
		Response.End()
	End Function

	Function DELETE_OLD_FILES(a)
		Dim folder, files, folderIdx, timer

		Set folder = objFSO.GetFolder(a)
		Set files = folder.Files

		timer = DateDiff("s","01/01/1970 00:00:00", Now())

		If files.count <> 0 Then
			For each folderIdx In files
				IF (folderIdx.Name+CCACHE)<timer Then
					If objFSO.Fileexists(a & "\" & folderIdx.Name) Then
						objFSO.DeleteFile a & "\" & folderIdx.Name
					End If
				End If
			Next
		End If
	End Function

	Function FILE_EXISTS(s)
		If s = 0 Then
			FILE_EXISTS = FILE_EXISTS(DateDiff("s","01/01/1970 00:00:00", Now()))
		Else
			If objFSO.Fileexists(s) Then
				FILE_EXISTS = FILE_EXISTS(s+1)
			Else
				FILE_EXISTS = s
			End If
		End If
	End Function

	Function JSENCODE(s)
		'Based in VBS JSON 2.0.3
		'Copyright (c) 2009 Tuðrul Topuz
		'Under the MIT license.

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

	absolutePath = Replace(Server.MapPath("./"),"\.","")

	Response.AddHeader "Access-Control-Max-Age", CCACHE
	Response.AddHeader "Access-Control-Allow-Origin", "*"
	Response.AddHeader "Access-Control-Request-Method", "*"
	Response.AddHeader "Access-Control-Allow-Methods", "OPTIONS, GET"
	Response.AddHeader "Access-Control-Allow-Headers", "*"

	Response.ContentType = "application/javascript"

	url = Request.QueryString("url")
	callback = Request.QueryString("callback")

	If url<>"" AND callback<>"" Then
		On Error Resume Next

			set filesys=CreateObject("Scripting.FileSystemObject")
			If Not filesys.FolderExists(absolutePath & "\" & PATH) Then
				filesys.CreateFolder(absolutePath & "\" & PATH)
			End If

			If Not filesys.FolderExists(absolutePath & "\" & PATH) Then
				ERROR_HANDLE("Failed to create the folder " & absolutePath & "\" & PATH)
			End If

			set xmlHttp = server.Createobject("MSXML2.ServerXMLHttp")

			xmlHttp.Open "GET", url, false

			xmlHttp.setRequestHeader "User-Agent", Request.ServerVariables("HTTP_USER_AGENT")
			xmlHttp.send ""

			If err = 0 Then
				If xmlHttp.Status = 200 Then
					tmpName = absolutePath & "\" & PATH

					Set oas = CreateObject("ADODB.Stream")
					oas.Open
					oas.Type = 1

					oas.Write xmlHttp.ResponseBody
					oas.Position = 0

					Set objFSO = Createobject("Scripting.FileSystemObject")

					DELETE_OLD_FILES(absolutePath & "\" & PATH)

					fileName = FILE_EXISTS(0)
					tmpName = tmpName & "\" & fileName

					'save responseBody to path
					oas.SaveToFile tmpName
					oas.Close
					Set oas = Nothing

					If objFSO.Fileexists(tmpName) Then
						Response.Expires = Round(CCACHE/60)
						Response.Write callback & "(""" & JSENCODE(FULL_URL() & PATH & "/" & fileName) & """);"
						Response.End()
					Else
						serr = "Não pode criar o arquivo " & tmpName
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
	End If

	If serr <> "" Then
		Response.AddHeader "Pragma", "no-cache"
		Response.AddHeader "Cache-control", "no-cache"
		Response.Expires = 0

		ERROR_HANDLE(serr)
	End If

	set xmlHttp = nothing
%>
