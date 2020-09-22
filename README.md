<div align="center">

## Unique Hit Counter


</div>

### Description

This code used cookies to count unique hits.

If the user has the cookie, then the hit isn't

counted. If they don't, then the hit is counted,

and they are given a cookie. All you have to do

is use an #include file="nameoffile.asp" and the

correct action is preformed. I made this script

for my host to count unique hits across his

networked sites. This is also my submission to

the asp world of PSC all my expreience is in the

VB world. I hope you guys like it!

Later

~/Andy
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[atwinda](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atwinda.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atwinda-unique-hit-counter__4-6501/archive/master.zip)

### API Declarations

2001 © Atwinda Software


### Source Code

```
<% 'ASP by Atwinda Software
Response.Expires = 0
Dim strNetSite
'This is the only place you have to change the name
'and the name of the file.
strNetSite = "sitename"
If CheckCookie(strNetSite) = "False" Then
	Call CountHit
	Call AddCookie
End If
Function CheckCookie(strCookieName)
If Request.Cookies(strCookieName) = "" Then
	CheckCookie = "False"
Else
	CheckCookie = "True"
End If
End Function
Function AddCookie()
Response.Buffer = True
Response.Cookies(strNetSite) = strNetSite
Response.Cookies(strNetSite).Expires = Date() + 1
End Function
Function CountHit()
On Error Resume Next
Dim objCntFSO, objCntFile, intHits
Set objCntFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objCntFile = objCntFSO.OpenTextFile(Server.MapPath(strNetSite & ".cnt"), 1)
intHits = objCntFile.ReadLine
objCntFile.Close
Set objCntFile = Nothing
If intHits = "" Then intHits = 0
intHits = intHits + 1
Set objCntFile = objCntFSO.CreateTextFile(Server.MapPath(strNetSite & ".cnt"), True)
objCntFile.Write intHits
objCntFile.Close
Set objCntFSO = Nothing
Set objCntFile = Nothing
End Function
%>
```

