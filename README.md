<div align="center">

## Print this page script


</div>

### Description

The code enables you to single out a section of your page for printing by enclosing it in two tags.
 
### More Info
 
The URL to the page you want printed

A nice, printable page

Non that I know of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter Graves](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter-graves.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-graves-print-this-page-script__4-7994/archive/master.zip)





### Source Code

```
<%
option explicit
'--------------------------------------------------------------------------------------
' Title : Print This Page Script
' Function: To print a section of a page
' In	 : Page URL (?ref=)
' Out  : Formatted page for prining
' By  : Peter Graves (ICQ - 116613728)
'
' Example : /print/thisscript.asp?ref=http://www.myweb.com/mypage.asp
'
' Notes : Could probably be improved, but for now it works
'
' To use : In your documents that you want to print wrap the content in
'   <sp> and </sp>
'--------------------------------------------------------------------------------------
Dim RefPage, objXMLHTTP, HTMLPage
RefPage = Request.QueryString("ref")
if Len(RefPage) = 0 or not Left(RefPage,7) = "http://" then
	response.write "<h1>Invalid reference page - " & RefPage & "</h1>"
	response.end
end if
Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
	objXMLHTTP.Open "GET", RefPage, False
	objXMLHTTP.Send
	HTMLPage = objXMLHTTP.responseText
Set objXMLHTTP = Nothing
%>
<html>
<head>
<title><%=GetPageTitle(HTMLPage)%></title>
<link href="/include/style.css" rel="stylesheet" type="text/css">
</head>
<body onLoad="JavaScript:print();">
<%=GetPageBody(HTMLPage)%>
<%=RefPage%>
</body>
</html>
<%
Function GetPageBody(HTMLstring)
dim tag1, tag2, temp1, temp2
tag1 = "<sp>"
tag2 = "</sp>"
temp1 = Split(HTMLstring,tag1)
If not uBound(temp1) = 1 Then
	Response.Write "An Error Occured - please notify the webmaster about this page - " & RefPage
	response.end
End If
temp2 = Split(temp1(1),tag2)
If not uBound(temp2) = 1 Then
	Response.Write "An Error Occured - please notify the webmaster about this page - " & RefPage
	response.end
End If
GetPageBody = temp2(0)
End Function
Function GetPageTitle(HTMLstring)
dim tag1, tag2, temp1, temp2
tag1 = "<title>"
tag2 = "</title>"
temp1 = split (HTMLstring, tag1)
temp2 = split (temp1(1), tag2)
GetPageTitle = temp2(0)
End Function
%>
```

