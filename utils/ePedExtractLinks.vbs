Dim Title,URL,ie,objFSO
Title = "Extracting Links from Google"
'Data = InputBox("Type something in Input Box to search Google.com"&vbCrlf&"Example VBScript to pull links from Internet Explorer ",Title,"VBScript to pull links from Internet Explorer ")
'Data="dymmy"
'If Data = "" Then 
' WScript.Quit
'Else
'URL = "https://www.google.com/search?q="&qq(Data)&""
'URL = "http://slleped01.ds.sll.se/eped/lists/14327858708973635638.html"
URL = "http://slleped01.ds.sll.se/eped/lists/14933337685310129998.html"
DestFolder = "F:\test\ePed\pdf\"

Set fso = CreateObject("Scripting.FileSystemObject") 
Set Link = fso.OpenTextFile(DestFolder&"LinkLog.htm",2,True,-1)
Set ie = CreateObject("InternetExplorer.Application")
Set Ws = CreateObject("wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
ie.Navigate (URL)
ie.Visible=false
Link.WriteLine "<html><body><h1>Denna kopia till backup-datorn uppdaterades "&Now&"</h1>"
Link.WriteLine "<h3>L&auml;nkar till ALB Barn l&auml;kemedelsblad</h3>"
Link.WriteLine "<p>Tryck&nbsp;Ctrl+F&nbsp;f&ouml;r&nbsp;att&nbsp;s&ouml;ka&nbsp;i&nbsp;listan<br/>"
Link.WriteLine "F&ouml;r&nbsp;kommentarer&nbsp;kontakta&nbsp;barnlakemedel@karolinska.se</p>"
DO WHILE ie.busy
 wscript.sleep 100
LOOP
Dim colLinks : Set colLinks = IE.Document.getElementsByTagName("a")
Dim objLink,InnetText
For Each objLink In colLinks
 HREF = objLink.href
 InnerText = objLink.InnerText
 'msgbox  HREF & vbcr & " |---------> " & InnerText,64,Title
 'Link.WriteLine HREF &" |---------> "& InnerText
 If InStr(HREF,"pdf")>0 Then
 
   parts2 = split(HREF,"/") 
   saveTo2 = parts2(ubound(parts2))

  myHTTPDownload HREF, DestFolder
  Link.WriteLine "<a href="& saveTo2 & ">" & InnerText&"</a><br>"
 End If
Next
'ws.Run "LinkLog.txt",1,False
Link.WriteLine "</body></html>"
'End If
WScript.Quit

Function qq(strIn)
qq = Chr(34) & strIn & Chr(34)
End Function

Sub myHTTPDownload( myURL, myPath )

  parts = split(myURL,"/") 
  saveTo = parts(ubound(parts))

' Fetch the file
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

objXMLHTTP.open "GET", myURL, false
objXMLHTTP.send()

If objXMLHTTP.Status = 200 Then
Set objADOStream = CreateObject("ADODB.Stream")
objADOStream.Open
objADOStream.Type = 1 'adTypeBinary
Const adSaveCreateOverWrite = 2

objADOStream.Write objXMLHTTP.ResponseBody
objADOStream.Position = 0    'Set the stream position to the start

Set objFSO = Createobject("Scripting.FileSystemObject")
'If objFSO.Fileexists(saveTo) Then
' objFSO.DeleteFile saveTo
 'Set objFSO = Nothing

 objADOStream.SaveToFile myPath & saveTo, adSaveCreateOverWrite
 objADOStream.Close
 Set objADOStream = Nothing
'End if

Set objXMLHTTP = Nothing
End if
End Sub

Sub HTTPDownload( myURL, myPath )
' This Sub downloads the FILE specified in myURL to the path specified in myPath.
'
' myURL must always end with a file name
' myPath may be a directory or a file name; in either case the directory must exist
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
'
' Based on a script found on the Thai Visa forum
' http://www.thaivisa.com/forum/index.php?showtopic=21832

    ' Standard housekeeping
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg
    Const ForReading = 1, ForWriting = 2, ForAppending = 8

    ' Create a File System Object
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ' Check if the specified target file or folder exists,
    ' and build the fully qualified path of the target file
    If objFSO.FolderExists( myPath ) Then
        strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
    ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
        strFile = myPath
    Else
        WScript.Echo "ERROR: Target folder not found."
        Exit Sub
    End If

    ' Create or open the target file
    Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )

    ' Create an HTTP object
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )

    ' Download the specified URL
    objHTTP.Open "GET", myURL, False
    objHTTP.Send

    ' Write the downloaded byte stream to the target file
    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next

    ' Close the target file
    objFile.Close( )
End Sub

