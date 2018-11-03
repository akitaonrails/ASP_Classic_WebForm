<%
on error resume next
strFile = Request( "file" )
strFile = Replace( strFile, "..", "", 1, -1 )
strFile = Replace( strFile, "\", "", 1, -1 )
strFile = Replace( strFile, "/", "", 1, -1 )

If strFile <> "" Then
	strFile = "D:\psn\PSN-Portal\frameworks\demo\" & strFile
	Set oFS = Server.CreateObject( "Scripting.FileSystemObject" )
	Set oFile = oFS.OpenTextFile( strFile )

	Response.ContentType = "text/plain"
	Response.Write oFile.ReadAll()

	Set oFile = Nothing
	Set oFS = Nothing
End If
%>