<%Session.Timeout=1000 
UserAgent = Trim(Lcase(Request.Servervariables("HTTP_USER_AGENT")))
If InStr(UserAgent,"teleport") > 0 or InStr(UserAgent,"webzip") > 0 or InStr(UserAgent,"flashget")>0 or InStr(UserAgent,"offline")>0 Then
	Response.Write "请不要采用teleport/Webzip/Flashget/Offline等工具来浏览网页！"
	Response.End
End If
if instr(request.ServerVariables("HTTP_Referer"),"/admin/")<=0 then
%>
<!--#Include File="inc/Neeao_SqlIn.Asp"-->
<%
end if
function format(num)
if num<1 then
response.Write(0)
response.Write(num)
else
response.Write(num)
end if
end function
%>