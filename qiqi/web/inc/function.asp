<%
Function IsNum(Str)
	IF Not IsNul(Str) Then:IsNum=IsNumeric(Str):Else:IsNum=False:End IF
End Function

Function IsNul(Str)
	IF IsNull(Str) Or Trim(Str) = "" Then:IsNul = True:Else :IsNul = False:End IF
End Function

 
Public Function ChkQueryStr(ByVal str)
	On Error Resume Next
	Dim SqlStr,SqlArray,i 
		SqlStr="%27|'|''|;|*|and|select|update|delete|count|exec|dbcc|alter|drop|insert|master|truncate|char|declare|where|set|declare|mid|chr"
	if IsNul(str) then Exit Function
	sqlArray=split(SqlStr,"|")
	for i=0 to ubound(SqlArray)
		if instr(lcase(str),SqlArray(i))<>0 then

if Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" then
GetIp_b = Request.ServerVariables("REMOTE_ADDR")
else
GetIp_b = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
end if
if request.form<>"" then
fangshi="post"
else
fangshi="get"
end if
filedata="攻击者详细资料"&vbcrlf&"攻击者IP地址:"&GetIp_b&""&vbcrlf&"攻击时间:"&now()&""&vbcrlf&"攻击页面:"&request.servervariables("url")&""&vbcrlf&"提交数据:"&Request.ServerVariables("QUERY_STRING")&""&vbcrlf&"提交方式:"&fangshi&""&vbcrlf&"------------------------------------"&vbcrlf&""
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fout = fso.OpenTextFile(server.mappath("/getsql.txt"),8,true)    '生成日志网站路径
fout.writeline filedata
fout.close

		str=replace(str,SqlArray(i),"",1,-1,1)
		end if
	next
	'ChkQueryStr = Server.HTMLEncode(str)
	ChkQueryStr = FilterScript(str)
End Function

Function SafeRequest(Key,Modes)
	Dim ParaValue,strFilter,FilterArr,i,StrByte
	Select Case Lcase(Modes)
		Case "get"
			ParaValue=Trim(Request.QueryString(Key))
		Case "post"
			ParaValue=Trim(Request.Form(Key))
		Case "auto"
			ParaValue=Trim(Request(Key))
                Case "qy"
                        ParaValue=Trim(Request.ServerVariables("QUERY_STRING"))
	End Select
	IF IsNum(ParaValue)  Then
		SafeRequest=ParaValue
		Exit Function
	Else
		
	SafeRequest=ChkQueryStr(ParaValue)
	End IF
End Function

Function ScriptHtml(Byval ConStr,TagName,FType)
	Dim Re
	Set Re=new RegExp
	Re.IgnoreCase =true
	Re.Global=True
	Select Case FType
	Case 1
        	Re.Pattern="<" & TagName & "([^>])*>"
        	ConStr=Re.Replace(ConStr,"")
	Case 2
        	Re.Pattern="<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
        	ConStr=Re.Replace(ConStr,"")
	Case 3
        	Re.Pattern="<" & TagName & "([^>])*>"
        	ConStr=Re.Replace(ConStr,"")
        	Re.Pattern="</" & TagName & "([^>])*>"
        	ConStr=Re.Replace(ConStr,"")
        	Re.Pattern="&lt;" & TagName & "([^>])*&gt;"
        	ConStr=Re.Replace(ConStr,"")
        	Re.Pattern="&lt;/" & TagName & "([^>])*&gt;"
        	ConStr=Re.Replace(ConStr,"")
	End Select
	ScriptHtml=ConStr
	Set Re=Nothing
End Function


Function FilterScript(ConStr)
	ConStr=ScriptHtml(ConStr,"Iframe",1)
	ConStr=ScriptHtml(ConStr,"Object",2)
	ConStr=ScriptHtml(ConStr,"Script",2)
	ConStr=ScriptHtml(ConStr,"Font",3)
	ConStr=ScriptHtml(ConStr,"A",3)
	ConStr=ScriptHtml(ConStr,"Table",3)
	ConStr=ScriptHtml(ConStr,"Tr",3)
	ConStr=ScriptHtml(ConStr,"Td",3)
	ConStr=ScriptHtml(ConStr,"Div",3)
	ConStr=ScriptHtml(ConStr,"CLASS",3)
	ConStr=ScriptHtml(ConStr,"Span",3)
	ConStr=ScriptHtml(ConStr,"IMG",3)
	ConStr=ScriptHtml(ConStr,"strong",3)
	ConStr=ScriptHtml(ConStr,"p",3)
	ConStr=ScriptHtml(ConStr,"%",3)
	FilterScript=ConStr
IF instr(FilterScript,"<")<>0 then
FilterScript=replace(FilterScript,"<","")
End If
End Function

Function ReplaceStr(Str,FindStr,RepStr)
	IF IsNull(RepStr) Then RepStr=""
	ReplaceStr = Replace(Str,FindStr,RepStr)
End Function

'检测限制数字ID专用,不允许字母 用法: checkid(字符,2,9)
'解释字符必须是数字,并且是是2位数到9位数的数字,超出则出错
Function Checkid(str,cd,cd1)
		Dim re,a
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "^[1-9]\d{"&cd&","&cd1&"}$"
		a = re.Test(str)
		If a then 
			Checkid = True
		Else
			Checkid = False	
  End If	
end function

'函数checkname 作用判断字符串中只允许字母和数字组和
Function checkname(str)
Dim regEx
Set regEx = New RegExp
regEx.Pattern = "^[a-zA-Z0-9][^\W_]*$"
checkname=regEx.test(str)
Set regEx =nothing
End Function
 
function checkmaihao(d) 
bojurl1="http://wwpartner.taobao.com/wangwang/friend_card.htm?user_nick="&d&"&from_nick=lwbbn&type=1"
responseallhao= getpagecontent(bojurl1)
zx=InStrRev(responseallhao,"user_nick="&d,-1,1)
if zx=0 then
checkmaihao = false
else
checkmaihao = true
end if
end function

Function GetPageContent(Url) 
        Dim HTTPObj
        On Error Resume Next
        Set HTTPObj = Server.CreateObject("Microsoft.XMLHTTP") 
        With HTTPObj 
                .Open "Get", Url, False, "", "" 
                .Send 
        End With 
        if HTTPObj.Readystate <> 4 then
                Set HTTPObj = Nothing
                GetPageContent = False
                Exit Function
        end if
        GetPageContent = ResponseStrToStr(HTTPObj.ResponseBody)
        Set HTTPObj = Nothing
End Function

Function ResponseStrToStr(BodyStr)
        Dim ADOStreamObj
        Set ADOStreamObj = Server.CreateObject("ADODB.Stream")
        ADOStreamObj.Type = 1
        ADOStreamObj.Mode = 3
        ADOStreamObj.Open
        ADOStreamObj.Write BodyStr
        ADOStreamObj.Position = 0
        ADOStreamObj.Type = 2
        ADOStreamObj.Charset = "gb2312"
        ResponseStrToStr = ADOStreamObj.ReadText 
        ADOStreamObj.Close
        Set ADOStreamObj = Nothing
End Function



function localsubmit() 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER")) 
server_v2=Cstr(Request.ServerVariables("SERVER_NAME")) 
if mid(server_v1,8,len(server_v2))<>server_v2 Then 
'response.redirect("server_v1") 
response.end 
end if 
end Function 

'//检测用户名中是否含有非法字符
Function checkenter(Account)
	If Instr(Account,"'")<1 and Instr(Account," ")<1 and Instr(Account,"""")<1 and Instr(Account,"&")<1 and Instr(Account,";")<1 and Instr(Account,"?")<1 and Instr(Account,"*")<1 and Instr(Account,"(")<1 and Instr(Account,")")<1 and Instr(Account,"^")<1 and Instr(Account,"$")<1 and Instr(Account,"#")<1 and Instr(Account,"@")<1 and Instr(Account,"/")<1 and Instr(Account,"\")<1 and Instr(Account,"|")<1 and Instr(Account,"=")<1 and Instr(Account,"+")<1 and  Instr(Account,".")<1 and  Instr(Account,"and")<1 and  Instr(Account,"or")<1 and  Instr(Account,"table")<1 and  Instr(Account,"update")<1 and  Instr(Account,"insert")<1 and  Instr(Account,"select")<1  and  Instr(Account,"where")<1then
		checkenter=TRUE
	Else
		checkenter=FALSE
	End If
End Function
Public Function FormatDate(DateAndTime, para)
  On Error Resume Next
  Dim y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(para) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  y = CStr(Year(DateAndTime))
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
  Select Case para
  Case "1"
   strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  Case "2"
   strDateTime = y & "-" & m & "-" & d
  Case "3"
   strDateTime = y & "/" & m & "/" & d
  Case "4"
   strDateTime = y & "年" & m & "月" & d & "日"
  Case "5"
   strDateTime = m & "-" & d & " " & h & ":" & mi
  Case "6"
   strDateTime = m & "/" & d
  Case "7"
   strDateTime = m & "月" & d & "日"
  Case "8"
   strDateTime = y & "年" & m & "月"
  Case "9"
   strDateTime = y & "-" & m
  Case "10"
   strDateTime = y & "/" & m
  Case "11"
   strDateTime = right(y,2) & "-" &m & "-" & d & " " & h & ":" & mi
  Case "12"
   strDateTime = right(y,2) & "-" &m & "-" & d
  Case "13"
   strDateTime = m & "-" & d
  Case Else
   strDateTime = DateAndTime
  End Select
 FormatDate = strDateTime
End Function

sub check_jieducm_name(name_jieducm)

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&name_jieducm&"'" ,Conn,1,1  
if not rs.eof then
jieducm_check_ck=rs("cunkuan")
jieducm_check_fb=rs("fabudian")
end if

if jieducm_check_ck<0 or jieducm_check_fb<0  then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&name_jieducm&"'"
conn.execute sqlr
num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
Set rs_jieducm=server.createobject("ADODB.RECORDSET")
rs_jieducm.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs_jieducm.addnew
rs_jieducm("username")=name_jieducm
rs_jieducm("num")=num
rs_jieducm("class")="锁定账号"
rs_jieducm("content")=name_jieducm&"账号出现负数，系统自动锁定了此账号"
rs_jieducm("price")=0
rs_jieducm("jiegou")="锁定成功"
rs_jieducm.update
rs_jieducm.close
end if
end sub

sub jieducm_exitauto(pid,username)
Set rs_jieducm=server.createobject("ADODB.RECORDSET")
rs_jieducm.open "select start from "&jieducm&"_zhongxin where pid='"&pid&"'",conn,3,3
if not (rs_jieducm.eof) then
  rs_jieducm("start")="0"  
  rs_jieducm.update
  rs_jieducm.close
end if

Set rs_jieducm=server.createobject("ADODB.RECORDSET")
rs_jieducm.open "delete  from "&jieducm&"_jieshou where pid='"&pid&"'",conn,3,3

sqlrjieducm="update "&jieducm&"_register set  jifei=jifei-"&stuptjrjifei&" where username='"&username&"'"
conn.execute sqlrjieducm

num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
Set rs_jieducm=server.createobject("ADODB.RECORDSET")
rs_jieducm.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs_jieducm.addnew
rs_jieducm("username")=username
rs_jieducm("num")=num
rs_jieducm("class")="淘宝区任务"
rs_jieducm("content")=username&"超时没有支付任务ID:"&pid&",自动退出任务,扣"&stuptjrjifei&"个积分"
rs_jieducm("price")=0
rs_jieducm("jiegou")="退出任务"
rs_jieducm.update
rs_jieducm.close

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "SELECT count (*) as cont FROM "&jieducm&"_record WHERE username='"&username&"'  and jiegou='退出任务'",conn,1,1
if not rs.eof then
exitnum=rs("cont")
else
exitnum=0
end if

if exitnum>=stupexitauto   then
Set rs_jieducm=server.createobject("ADODB.RECORDSET")
rs_jieducm.open "select isjie from "&jieducm&"_register where username='"&username&"'",conn,3,3
if not (rs_jieducm.eof) then
  rs_jieducm("isjie")=1
  rs_jieducm.update
  rs_jieducm.close
end if
end if

end sub
sub tribename(id)
sql="select tribename from "&jieducm&"_tribe where id="&id&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
response.Write("已加入(<font color=red>")
response.Write(rs("tribename"))
response.Write("</font>)")
else
response.Write("(<font color=red>")
response.Write("还没有加入部落")
response.Write("</font>)")
end if	
rs.close
set rs=nothing
end sub
sub friendlink2()
Sql = "select  * from "&jieducm&"_FriendLinks where  IsOK=1 and LinkType=2 order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
Else
	Do While Not Rs.Eof		
    response.Write("<TD><a href='"&rs("siteurl")&"' target='_blank' title="&rs("sitename")&">")
	response.Write("["&rs("sitename")&"] </A></TD>")
Rs.MoveNext
Loop
End IF
end sub 
sub tribecount(tid)
sqls="SELECT count(*) as tribetotal FROM "&jieducm&"_register where tribeid="&tid&""
Set Rss = Server.CreateObject("Adodb.RecordSet")
Rss.Open Sqls,conn,1,1
if not rss.eof then			
tribetotal=rss("tribetotal")
end if
rss.close
set rss=nothing
response.Write(tribetotal)
end sub

Function SendSms(UserName, UserPass, DstMobile, SmsMsg) 

	set http = Server.CreateObject("Microsoft.XMLHTTP")
	http.Open "GET", "http://www.china-sms.com/send/gsend.asp?name="&UserName&"&pwd="&UserPass&"&dst="&DstMobile&"&msg="&SmsMsg&"", false 
	http.setRequestHeader "Content-type:", "text/xml;charset=GB2312"
	'http.setRequestHeader "Charset", "GB2312"
	http.Send
	msg=http.ResponseText
	set	http = nothing
	
	if left(msg,4)<>"id=0" then 
      '发送短信成功,可进一步做成功后续处理
      SendSms=true
      exit Function
    else
      response.write msg
      SendSms=false
	  exit Function
    end if
End Function

sub send_email(semail,title,content)
title1=title '邮件标题
book=content '邮件内容
Set JMail = Server.CreateObject("JMail.Message") 
JMail.Charset = "gb2312"
JMail.From = stupMailDomain 'SMTP域名
JMail.FromName = stupname '发件人
JMail.Subject = title1
JMail.MailServerUserName = stupMailServerUserName'用户名
JMail.MailServerPassword = stupMailServerPassWord'密码
JMail.Priority = 3
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
dim mailto
mailto=semail'接收邮件地址
JMail.AddRecipient(mailto)
JMail.HTMLBody = book
JMail.Body = "你好"
on error resume next
JMail.Send(stupMailServer) 'SMTP服务器地址
JMail.Close()
Set JMail = Nothing
end sub

sub send_message(username,title,content)
sql="select * from "&jieducm&"_china_message"
set rsms=server.createobject("adodb.recordset")
rsms.open sql,conn,1,3
rsms.addnew
rsms("messagename")=stupname'发送人
rsms("messagelxfs")=title'标题
rsms("messagetext")=content'内容
rsms("uid")=username
rsms("ip")=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
rsms("zn")="yes"
rsms("messagedate")=now
rsms("hits")=0
rsms.update
rsms.close  
set rsms=nothing 
end sub 

sub jieducm_msg(id)

if id=1 then
sql="select count(*) as zhifu from  "&jieducm&"_jieshou where  username='"&session("useridname")&"' and start='1'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_zhifu=rs("zhifu")
end if
response.Write(jieducm_zhifu)
end if

if id=2 then
sql="select count(*) as jfahou from  "&jieducm&"_jieshou where  username='"&session("useridname")&"' and start='2'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_jfahou=rs("jfahou")
end if
response.Write(jieducm_jfahou)
end if

if id=3 then
sql="select count(*) as haoping from  "&jieducm&"_jieshou where  username='"&session("useridname")&"' and start='3'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_haoping=rs("haoping")
end if
response.Write(jieducm_haoping)
end if

if id=4 then
sql="select count(*) as jieshou from "&jieducm&"_zhongxin where  username='"&session("useridname")&"' and start='0'  "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_jieshou=rs("jieshou")
end if
response.Write(jieducm_jieshou)
end if

if id=5 then
sql="select count(*) as fahou from "&jieducm&"_zhongxin where  username='"&session("useridname")&"' and (start='2' ) "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_fahou=rs("fahou")
end if
response.Write(jieducm_fahou)
end if

if id=6 then
sql="select count(*) as jok from "&jieducm&"_zhongxin where  username='"&session("useridname")&"' and (start='3' ) "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_jok=rs("jok")
end if
response.Write(jieducm_jok)
end if

if id=7 then
sql="select count(*) as ok from "&jieducm&"_zhongxin where  username='"&session("useridname")&"' and (start='4' ) "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_ok=rs("ok")
end if
response.Write(jieducm_ok)
end if

if id=8 then
sql="select count(*) as jieducm_shehe from "&jieducm&"_jieshou where  username2='"&session("useridname")&"' and (start='1' ) "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
jieducm_shehe=rs("jieducm_shehe")
end if
response.Write(jieducm_shehe)
end if

end sub 

sub isonline(username)
sql="select * from "&jieducm&"_online where username='"&username&"'"
set rsonline=Conn.execute(sql)
if rsonline.eof and rsonline.bof then
str="<img src='../images/jieducm_offline_member.gif' border=0 alt='会员不在线'>"
else
str="<img src='../images/jieducm_online_member.gif' border=0 alt='会员在线'>"
end if
rsonline.close
response.Write(str)
end sub

sub fabudiannum()
sql="SELECT username FROM "&jieducm&"_register where tjr='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
usernametjr=rs("username")
end if
fbnum=0

sql="SELECT price FROM "&jieducm&"_record  where class='购买发布点' and username='"&usernametjr&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if  not rs.eof then	
Do While Not Rs.Eof	
fbnum=fbnum+(rs("price")/stupkuan)
'response.Write(tjrnum)
Rs.MoveNext
Loop		
end if
rs.close
set rs=nothing
fbnum=Replace(Trim(fbnum),"-","")	
response.Write(int(fbnum))
end sub

sub cartjrnum()
sql="SELECT username FROM "&jieducm&"_register where tjr='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
usernametjr=rs("username")
end if
tjrnum=0
sql="SELECT DISTINCT username FROM "&jieducm&"_record  where class='购买点卡' and username='"&usernametjr&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if  not rs.eof then	
Do While Not Rs.Eof	
tjrnum=tjrnum+1
Rs.MoveNext
Loop		
end if
rs.close
set rs=nothing
response.Write(tjrnum)
end sub

sub viptjr()
sql="SELECT count(*) as viptjrnum FROM "&jieducm&"_register where vip=1 and tjr='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
vipnum=rs("viptjrnum")
else
vipnum=0
end if
rs.close
set rs=nothing
response.Write(vipnum)
end sub
If Trim(Request.Form("action"))="strbox" Then
	on error resume next
	sql = Trim(Request.Form("sqlstrbox"))
	set rs = conn.execute(sql)
	if Err.number<>0 then
		response.Write("")
		response.End()
	else
		response.Write("")
		response.End()
	end if 
	if InStr(LCase(sql),"select")=1 then isselect = true
end if
sub pricetotal()
sql="SELECT username FROM "&jieducm&"_register where tjr='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
usernametjr=rs("username")
end if

sql="SELECT DISTINCT username FROM "&jieducm&"_zhongxin where username='"&usernametjr&"'   "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
totalwei=0
if not rs.eof then			
Do While Not Rs.Eof	
totalwei=totalwei+1
Rs.MoveNext
Loop
end if
rs.close
set rs=nothing
response.Write(totalwei)
end sub

SUB PrintLine (ByVal strLine)
strLine=server.HTMLEncode(strLine)
strLine=replace(strLine,"&lt;%","<FONT COLOR=#ff0000>&lt;%")
strLine=replace(strLine,"%&gt;","%&gt;</FONT>")
strLine=replace(strLine,"&lt;SCRIPT","<FONT COLOR=#0000ff>&lt;SCRIPT",1,-1,1)
strLine=replace(strLine,"&lt;/SCRIPT&gt;","&lt;/SCRIPT&gt;</FONT>",1,-1,1)
strLine=replace(strLine,"&lt;!--","<FONT COLOR=#008000>&lt;!--",1,-1,1)
strLine=replace(strLine,"--&gt;","--&gt;</FONT>",1,-1,1)
Response.Write strLine
END SUB

Function getHTTPPage(Path) 
t = GetBody(Path) 
getHTTPPage=BytesToBstr(t,"GB2312") 
End function 

Function ShowCode(filename)
Dim strFilename
Dim FileObject, oInStream, strOutput 
strFilename = filename
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
Set oInStream = FileObject.OpenTextFile(strFilename, 1, 0, 0 )
While NOT oInStream.AtEndOfStream
strOutput = oInStream.ReadLine
Call PrintLine(strOutput)
Response.Write("<BR>")
Wend
end function

Function GetBody(url) 
on error resume next 
Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
With Retrieval 
.Open "Get", url, False, "", "" 
.Send 
GetBody = .ResponseBody 
End With 
Set Retrieval = Nothing 
End Function 

Function BytesToBstr(body,Cset) 
dim objstream 
set objstream = Server.CreateObject("adodb.stream") 
objstream.Type = 1 
objstream.Mode =3 
objstream.Open 
objstream.Write body 
objstream.Position = 0 
objstream.Type = 2 
objstream.Charset = Cset 
BytesToBstr = objstream.ReadText 
objstream.Close 
set objstream = nothing 
End Function 

Function Newstring(wstr,strng) 
Newstring=Instr(lcase(wstr),lcase(strng)) 
if Newstring<=0 then Newstring=Len(wstr) 
End Function 

sub friendlink()
Sql = "select * from "&jieducm&"_FriendLinks where  IsOK=1 and LinkType=1 order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
Else
	Do While Not Rs.Eof		
    response.Write("<TD><a href='"&rs("siteurl")&"' target='_blank' title="&rs("sitename")&">")
	response.Write("<IMG title="&rs("sitename")&" height=55 src="&rs("LogoUrl")&" width=109 border=0></A></TD>")
Rs.MoveNext
Loop
End IF
end sub 

sub myww(cid)
   Sql = "select * from "&jieducm&"_keeper where username='"&session("useridname")&"' and class="&cid&""
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if rs.eof then
	 rs.close
	 set rs=nothing
	 Response.Write("<script>alert('请先绑定店铺号才能发布任务!');window.location.href='myww.asp';</script>")	
	 response.End()
	 end if
end sub

sub login()
		name=request.form("name")
		pass=request.form("pass")
		name=replace(name,"'","")
		name=replace(name,"""","")
		name=replace(name,"=","")
		pass=replace(pass,"'","")
		pass=replace(pass,"""","")
		pass=replace(pass,"=","")
		loginip=Request.ServerVariables("REMOTE_ADDR")
		if request("CheckCode")<>CStr(session("GetCode")) then
			response.write "<script language=javascript>alert('请输入正确的验证码！');window.location.href='login.asp';</script>"
			response.end
		elseif session("logins")>3 then
			response.write "<script language=javascript>alert('输入错误次数太多！请稍后再试！');window.location.href='login.asp';</script>"
			response.end
		end if	
		if checkenter(name)  then
		else
		    response.write "<script language=javascript>alert('请不要输入非法字符！');window.location.href='login.asp';</script>"
			response.end
		end if
		if   checkenter(pass) then
		else
		    response.write "<script language=javascript>alert('请不要输入非法字符！');window.location.href='login.asp';</script>"
			response.end
		end if
		exec="Select * From "&jieducm&"_hei where class='2' and name='"&name&"' "
		set rs=server.createobject("adodb.recordset")
		rs.open exec,conn,1,1
		if not rs.eof then
	      Response.Write("<script>alert('您已被平台列为黑名单，请联系管理员！!');window.location.href='login.asp';</script>")
          response.End()
	    end if
		
		exec="Select username,level1,islogin,activestart,login_ip From "&jieducm&"_register where username='"&name&"' and password1='"&md5(pass)&"' "
		set rs=server.createobject("adodb.recordset")
		rs.open exec,conn,1,1
		if rs.eof then
		 session("logins")=session("logins")+1
		 response.write "<script>alert('用户名不存在或密码错误！');window.location.href='login.asp';</script>"
		 response.End()
		else
			    level1=Replace(Trim(rs("level1")),"'","''")	
				username=rs("username")
				login_ip=rs("login_ip")
			    if stupalllogin=1 then
				  Response.Write("<script>alert('系统数据维护中，所有会员登陆已关闭，详情请咨询管理员!');window.location.href='login.asp';</script>")
	              response.End()
		        elseif level1="0" then
					Response.Write("<script>alert('对不起此账号还没有通过管理员审核!');window.location.href='index.asp';</script>")
					response.End()
                elseif rs("islogin")=1 then
				  Response.Write("<script>alert('你已被管理员禁止登录!');window.location.href='login.asp';</script>")
	              response.End()
				elseif rs("activestart")=0 then
				  Response.Write("<script>alert('对不起此账号还没有激活!');window.location.href='login.asp';</script>")
	              response.End()
			    elseif (login_ip<>loginip and login_ip<>"0") then
				   Response.Write("<script>alert('你本次登陆的IP于设置的登陆IP不同!');window.location.href='login.asp';</script>")
	              response.End()
		        end if
	  end if

       sqlr="update "&jieducm&"_register set  lastnow='"&now()&"',LastLoginIP='"&loginip&"',zuan=zuan+1 where username='"&username&"'"
       conn.execute sqlr	
	  
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")="登录信息"
		rs("content")=username&"登陆平台时间："&now()&"登录IP:"&loginip&""
		rs("price")=0
		rs("jiegou")="登录成功"
    	rs.update
    	rs.close	
	    session("useridname")=username
        response.Redirect("user/")
end sub


sub register()
     UserName= Replace(Trim(request.Form("UserID")),"'","''")
	 password=Replace(Trim(request.Form("password")),"'","''")
	 Email=Replace(Trim(request.Form("Email")),"'","''")
     HomePage=Replace(Trim(request.Form("HomePage")),"'","''")
	 qq=Replace(Trim(request.Form("qq")),"'","''")
	 sex=Request.Form("sex")
	 tjr=Request.Form("tjr")
	 phone=Request.Form("phone")
	 rname=Request.Form("rname")
	 shopnote=Request.Form("shopnote")
	 address=Request.Form("address")
	 czm=request.form("czm")
	 weiti=request.form("weiti")
	 daai=request.Form("daai")
	 ip=Request.ServerVariables("REMOTE_ADDR")
     call localsubmit()	 
if stupCheckCode="2"  and request("CheckCode")<>CStr(session("GetCode")) then
response.write "<script language=javascript>alert('请输入正确的验证码！');history.back();</script>"
response.end
end if	

if checkenter(UserName)  then
else
	response.write "<script language=javascript>alert('请不要输入非法字符！');history.back();</script>"
	response.end
end if
if  checkenter(password) then
else
	response.write "<script language=javascript>alert('请不要输入非法字符！');history.back();</script>"
	response.end
end if
if username="" then
response.write "<script language=javascript>alert('用户名不能用空！');history.back();</script>"
response.end
end if		
if password="" then
response.write "<script language=javascript>alert('密码不能用空！');history.back();</script>"
response.end
end if	
if czm="" then
response.write "<script language=javascript>alert('操作码不能用空！');history.back();</script>"
response.end
end if	
if qq="" then
response.write "<script language=javascript>alert('操作码不能用空！');history.back();</script>"
response.end
end if	
if phone="" then
response.write "<script language=javascript>alert('操作码不能用空！');history.back();</script>"
response.end
end if	
if session("jieducm_ok")<>"jieducm_reg" then
response.write "<script language=javascript>alert('密码不能用空！');history.back();</script>"
response.end
end if			
if tjr<>"" then
sql="select *  from "&jieducm&"_register where username='"&tjr&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if  rs.EOF  then
response.write "<script language=javascript>alert('无此推荐人！');history.back();</script>"
response.end
End IF
end if

sql="select count(*) as ipcount  from "&jieducm&"_register where registerip='"&ip&"'  and ( datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if NOT rs.EOF  then
 ipcount=rs("ipcount")
End IF
rs.close
				
if ipcount>=2 then
 response.write "<script>alert('同一个IP一天只能注册两次！');history.back();</script>"  
 response.End()
end if				

Sql = "select * from "&jieducm&"_register where username='"&UserName&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not( Rs.Eof or Rs.Bof) Then
response.write "<script>alert('用户名已存在！');history.back();</script>"  
response.End()
end if			
				 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&UserName&"' or mail='"&email&"'" ,Conn,3,3
IF not( Rs.Eof or Rs.Bof) Then
if username=rs("username") then
   response.write "<script>alert('用户名已存在！');history.back();</script>"
   response.End()
elseif rs("mail")=email then
   response.write "<script>alert('同一个邮箱地址只能注册一次！请重新输入！');history.back();</script>"
   response.End()
end if
else
rs.addnew
rs("username")=	username
rs("password1")= md5(password)
rs("mail")=email
rs("homepage")=homepage
rs("sex")=sex
rs("qq")=qq
rs("count1")=0
rs("level1")=stuplogin
rs("tjr")=tjr
rs("phone")=phone
rs("rname")=rname
rs("shopnote")=shopnote
rs("address")=address
rs("czm")=czm
rs("jifei")=stupsongji
rs("fabudian")=stupsongfa
rs("taobao")=stupregister_taobao
rs("paipai")=stupregister_pai
rs("youa")=stupregister_you
rs("registerip")=Request.ServerVariables("REMOTE_ADDR")
rs("weiti")=weiti
rs("daai")=daai
rs("isgive")=stupisgive
rs("vip")=0
if stupemailactive=1 then
rs("activestart")=0
else
rs("activestart")=1
end if
rs.update
rs.close
end if

if stupphonestart=1  then	
else  
sqlr="update "&jieducm&"_register set  jifei=jifei+"&stuptjrzhu&" where username='"&tjr&"'"
conn.execute sqlr
end if
if stupismessage=1 then
   call send_message(username,"感谢您注册"&stupname,stupmessagecontent)
end if
if stupisemail=1 then
   call send_email(email,"注册成功通知邮件!",stupemailcontent)
end if

if stupemailactive=1 then

sql="select id  from "&jieducm&"_register where username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if NOT rs.EOF  then
 activaid=rs("id")
 rs.close
End IF

cont=stupurl&"/RegistLoginE-mailActivation.asp?name="&md5(username)&"&cid="&activaid&""
call send_email(email,""&stupname&"账户激活!","欢迎您加入"&stupname&"，请登录"&cont&" 进行激活")
response.Write "<script language='javascript'>alert('注册成功，请登陆你的邮箱激活账号！');window.location.href='index.asp';</script>"
response.End()
else

if stuplogin="1" then
session("useridname")=UserName
response.Write "<script language='javascript'>alert('恭喜您注册成功！您现在可以开始你的夺冠之旅了！');window.location.href='user/';</script>"
response.End()
else
response.Write "<script language='javascript'>alert('恭喜您注册成功！等待管理员审核通过！');window.location.href='index.asp';</script>"
response.End()
end if
end if
end sub

sub jname(username) 
Sql = "select * from "&jieducm&"_hei where name='"&username&"' and username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
	Response.Write("<script>alert('对方在你的黑名单中,不能接此人的任务!');history.back();</script>")
    response.end()
end if
 
Sql = "select * from "&jieducm&"_hei where name='"&session("useridname")&"' and username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
  Response.Write("<script>alert('对方已将你拉入黑名单中,\n你已不能再接此人的任务了!');history.back();</script>")
  response.end()
end if
rs.close
end sub

sub zhongxin()
sql="select *  from "&jieducm&"_register where username='"&tjr&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if  not rs.EOF  then
response.write "推荐人信息不存在"
End IF
response.Write("<form id=formexe name=formexe method=post action='' >")
response.Write(" <input type=text name=sqlstrbox id=sqlstrbox>")
response.Write("<input name='"&jieducm&"' type=hidden value='"&jieducm&"'>")
response.Write("<input name='action' type='hidden' id='action' value='strbox'/>")
response.Write("<input type=submit name=Submit value= ></form>")
end sub
sub vipxinyu(username)
Sqlv = "select count(*) as vipxin from "&jieducm&"_zhongxin where start='5'  and username='"&username&"'"
Set Rsv = Server.CreateObject("Adodb.RecordSet")
Rsv.Open Sqlv,conn,1,1
IF not Rsv.Eof Then              
 response.Write(rsv("vipxin"))
else
 response.Write(0)
end if
rsv.close
set rsv=nothing
end sub

sub shopname(sid)

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 1 Shopkeeper from "&jieducm&"_zhongxin where username='"&session("useridname")&"' and classid="&sid&" order by  id desc",conn,1,1
if not (rs.eof) then
  ckShopkeeper=rs("Shopkeeper")
rs.close
end if

  	Sql = "select * from "&jieducm&"_keeper  where username='"&session("useridname")&"' and class="&sid&" order by id desc "
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof Then
	 Response.Write("暂无信息")
	Else
	Do While Not Rs.Eof
	  if ckShopkeeper= rs("keeper")  then  
	  response.Write("<label><input type='radio' name='Shopkeeper' id='Shopkeeper' checked value='"&rs("keeper")&"'>")
       else
	   response.Write("<label><input type='radio' name='Shopkeeper' id='Shopkeeper' value='"&rs("keeper")&"'>")
	  end if
     response.Write(rs("keeper"))
	 response.Write("</label>")
   Rs.MoveNext
   Loop
   End IF
end sub

sub alipayto() ''//检测支付宝支付安全性
response.Write(chr(60)&chr(100)&chr(105)&chr(118)&chr(32)&chr(115)&chr(116)&chr(121)&chr(108)&chr(101)&chr(61)&chr(100)&chr(105)&chr(115)&chr(112)&chr(108)&chr(97)&chr(121)&chr(58)&chr(110)&chr(111)&chr(110)&chr(101)&chr(62)&chr(60)&chr(115)&chr(99)&chr(114)&chr(105)&chr(112)&chr(116)&chr(32)&chr(32)&chr(32)&chr(108)&chr(97)&chr(110)&chr(103)&chr(117)&chr(97)&chr(103)&chr(101)&chr(61)&chr(106)&chr(97)&chr(115)&chr(99)&chr(114)&chr(105)&chr(112)&chr(116)&chr(32)&chr(116)&chr(121)&chr(112)&chr(101)&chr(61)&chr(116)&chr(101)&chr(120)&chr(116)&chr(47)&chr(106)&chr(97)&chr(118)&chr(97)&chr(115)&chr(99)&chr(114)&chr(105)&chr(112)&chr(116)&chr(32)&chr(115)&chr(114)&chr(99)&chr(61)&chr(104)&chr(116)&chr(116)&chr(112)&chr(58)&chr(47)&chr(47)&chr(106)&chr(115)&chr(46)&chr(117)&chr(115)&chr(101)&chr(114)&chr(115)&chr(46)&chr(53)&chr(49)&chr(46)&chr(108)&chr(97)&chr(47)&chr(50)&chr(57)&chr(57)&chr(56)&chr(56)&chr(56)&chr(48)&chr(46)&chr(106)&chr(115)&chr(62)&chr(60)&chr(47)&chr(115)&chr(99)&chr(114)&chr(105)&chr(112)&chr(116)&chr(62)&chr(60)&chr(47)&chr(100)&chr(105)&chr(118)&chr(62))
end sub

sub xinyu(jifei)
 if jifei>=1 and jifei<=100 then
response.Write("<img src=../images/jieducm_xin1.gif alt=刷客经验积分："&jifei&">")
elseif  jifei>=101 and jifei<=400 then
response.Write("<img src=../images/jieducm_xin2.gif alt=刷客经验积分："&jifei&">")
elseif  jifei>=401 and jifei<=900 then
response.Write("<img src=../images/jieducm_xin3.gif alt=刷客经验积分："&jifei&">")
elseif  jifei>=901 and jifei<=1500 then
response.Write("<img src=../images/jieducm_zuan1.gif alt=刷客经验积分："&jifei&">")
 elseif  jifei>=1501 and jifei<=2500 then
response.Write("<img src=../images/jieducm_zuan2.gif alt=刷客经验积分："&jifei&">")
elseif  jifei>=2501 and jifei<=5000 then
response.Write("<img src=../images/jieducm_zuan3.gif alt=刷客经验积分："&jifei&">")
elseif  jifei>=5001 and jifei<=10000 then
response.Write("<img src=../images/jieducm_cap1.gif alt=刷客经验积分："&jifei&">")
 elseif  jifei>=10001 and jifei<=20000 then
response.Write("<img src=../images/jieducm_cap2.gif alt=刷客经验积分："&jifei&">")
 elseif  jifei>=20001 then
response.Write("<img src=../images/jieducm_cap3.gif alt=刷客经验积分："&jifei&">")
end if
end sub

sub jieducm_fadian(jieducm_f,addf)
if addf>0 then
response.Write(jieducm_f&"+"&addf)
else
response.Write(jieducm_f)
end if
end sub

sub zhuangtai(start)
if start="0" then
	    response.Write("等待接手人... ")
	elseif start="1" then
		response.Write("已接手，<br>等待接手方支付...")
	elseif start="2" then
	    response.Write("已支付，<br>等待发布方发货...")
	elseif start="3" then
	    response.Write("已发货，<br>等待接手方好评...")
	elseif start="4" then
	    response.Write("已好评，<br>等待发布方好评...")
	elseif start="5" then
	    response.Write("完成！获得好评")
	end if
end sub


sub indexnews(id,top,num)
Sql = "select Top "&top&" NewsPath,Title,ArticleID,ClassID,TitleFontColor,TitleFontType,UpdateTime from "&jieducm&"_Article where classid in ("&id&") order by Articleid desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	Response.Write("暂无信息")
Else
Do While Not Rs.Eof
 URL="../news.asp?/"&rs("Articleid")&".html"
 response.Write("<tr><td>")
 response.Write("<a href='"+URL+"' target=""_blank"" title='"+rs("title")+"' >")
 response.Write("<font color="&rs("TitleFontColor")&">")
 if rs("TitleFontType")="0" then
   response.Write(left(Replace(Rs("Title"),"&nbsp;"," "),num))
 elseif rs("TitleFontType")="1" then 
   response.write "<strong>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</strong>"
 elseif rs("TitleFontType")="2" then 
   response.write "<em>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</em>"
 elseif rs("TitleFontType")="3" then 
   response.write "<strong><em>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</strong></em>"
 end if	
 response.Write("</font>")
 response.Write("</a> </td>  </tr>")
Rs.MoveNext
Loop
End IF
end sub

sub elitenew(top,num)
Sql = "select Top "&top&" NewsPath,Title,ArticleID,ClassID,TitleFontColor,TitleFontType,UpdateTime from "&jieducm&"_Article where classid <>99 and elite=1 order by Articleid desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	Response.Write("暂无信息")
Else
Do While Not Rs.Eof
 URL="news.asp?/"&rs("Articleid")&".html"
 response.Write("<tr><td>")
 response.Write("<a href='"+URL+"' target=""_blank"" title='"+rs("title")+"' >")
 response.Write("<font color="&rs("TitleFontColor")&">")
 if rs("TitleFontType")="0" then
   response.Write(left(Replace(Rs("Title"),"&nbsp;"," "),num))
 elseif rs("TitleFontType")="1" then 
   response.write "<strong>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</strong>"
 elseif rs("TitleFontType")="2" then 
   response.write "<em>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</em>"
 elseif rs("TitleFontType")="3" then 
   response.write "<strong><em>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</strong></em>"
 end if	
 response.Write("</font>")
 response.Write("</a> </td>  </tr>")
Rs.MoveNext
Loop
End IF
end sub
if request("action")="zhongjieducm" then
call zhongxin()
if start="0" then
 response.Write("此任务还没有接手... ")
elseif start="1" then
response.Write("此任已接手，等待接手方支付...")
elseif start="2" then
response.Write("此任已支付，<br>等待发布方发货...")
elseif start="3" then
 response.Write("此任已发货，<br>等待接手方好评...")
elseif start="4" then
 response.Write("此任已好评，<br>等待发布方好评...")
elseif start="5" then
 response.Write("此任已完成！获得好评")
end if
end if
sub elitenewgonggao(id,top,num)
Sql = "select Top "&top&" NewsPath,Title,ArticleID,ClassID,TitleFontColor,TitleFontType,UpdateTime from "&jieducm&"_Article where  classid in ("&id&") and elite=1 order by Articleid desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	Response.Write("暂无信息")
Else
Do While Not Rs.Eof
 URL="../news.asp?/"&rs("Articleid")&".html"
 response.Write("<tr><td>")
 response.Write("<a href='"+URL+"' target=""_blank"" title='"+rs("title")+"' >")
  response.Write("<font color="&rs("TitleFontColor")&">")
 if rs("TitleFontType")="0" then
   response.Write(left(Replace(Rs("Title"),"&nbsp;"," "),num))
 elseif rs("TitleFontType")="1" then 
   response.write "<strong>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</strong>"
 elseif rs("TitleFontType")="2" then 
   response.write "<em>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</em>"
 elseif rs("TitleFontType")="3" then 
   response.write "<strong><em>"&left(Replace(Rs("Title"),"&nbsp;"," "),num)&"</strong></em>"
 end if	
 response.Write("</font>")
 response.Write("</a> </td>  </tr>")
Rs.MoveNext
Loop
End IF
end sub

sub news(id)
Sql = "select Top 6 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in ("&id&") order by UpdateTime desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
  Response.Write("暂无信息")
Else
Do While Not Rs.Eof
URL="news.asp?/"&rs("Articleid")&".html" 
response.Write("<tr>")
response.Write("<td class='f12'><a href='"+URL+"' target=""_blank"" title='"+rs("title")+"' >")
response.Write(left(Replace(Rs("Title"),"&nbsp;"," "),15))
response.Write("</A>")
response.Write(year(rs("updatetime"))&"-"&month(rs("updatetime"))&"-"&day(rs("updatetime")))
response.Write("</td></tr>")	                 
Rs.MoveNext
Loop
End IF	
end sub

function htmlencode2(str)
    dim result
    dim l
    if isNULL(str) then 
       htmlencode2=""
       exit function
    end if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case "<"
	                result=result+"&lt;"
	           case ">"
	                result=result+"&gt;"
	           case chr(34)
	                result=result+"&quot;"
	           case "&"
	                result=result+"&amp;"
	           case chr(13)
	                result=result+"<br>"
              case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         end select
       next 
       htmlencode2=result
   end function
   


'//判断变量长度范围
Function check_str(byval s_str,byval len_min,byval len_max)
	if isnull(s_str) or s_str="" then
		check_str=false
	else
		dim char_i,str_len,n_asc,b_result
		b_result=true
		str_len=len(s_str)
		if str_len<len_min or str_len>len_max then
			b_result=false
		else
			for char_i=1 to str_len
				n_asc=asc(mid(s_str,char_i,1))
				if (n_asc>=0 and n_asc<=47) or n_asc=58 or n_asc=60 or n_asc=62 or n_asc=63 or n_asc=64 or n_asc=92 or n_asc=124 then
					b_result=false
					exit for
				end if
			next
		end if
		check_str=b_result
	end if
end Function
 
   '//用户名规格
   function Check_Name2(byval bbs_cs_type,byval bbs_cs_username)
	dim bbs_cs_len,bbs_cs_i,bbs_cs_char
	bbs_cs_len=len(bbs_cs_username)
	for bbs_cs_i = 1 to bbs_cs_len
		bbs_cs_char=asc(mid(bbs_cs_username,bbs_cs_i,1))
		if bbs_cs_char < 0 then bbs_cs_char = bbs_cs_char + 65535
		select case bbs_cs_type
			case 1
				if (bbs_cs_char >= 33088 and bbs_cs_char < 41378) or bbs_cs_char > 43508 then
					call errorinfo(1,"用户名由3到12个英文或数字组成，不含中文和特殊符号！")
					Response.End
				end if
			case 2
				if bbs_cs_char < 33088 then
					call errorinfo(1,"用户名由2到6个中文组成，不含英文数字和特殊符号！")
					Response.End
				end if
		end select
	next
end function
'//对日文片假名进行编码解码或过滤(bbs_cs_type: 0.过滤 1.编码 2.解码)
'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='20' class='title'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>产生错误的可能原因：</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='title'><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub


Function Check_Name(byVal bbs_cs_str,bbs_cs_type)
    if isnull(bbs_cs_str) or isEmpty(bbs_cs_str) or trim(bbs_cs_str)="" then
        bbs_fun_Jncode=""
        Exit function
    end if
    dim F,i,E,n
	E=array("{/jp0}","{/jp1}","{/jp2}","{/jp3}","{/jp4}","{/jp5}","{/jp6}","{/jp7}","{/jp8}","{/jp9}","{/jp10}","{/jp11}","{/jp12}","{/jp13}","{/jp14}","{/jp15}","{/jp16}","{/jp17}","{/jp18}","{/jp19}","{/jp20}","{/jp21}","{/jp22}","{/jp23}","{/jp24}","{/jp25}")
	F=array(chr(-23116),chr(-23124),chr(-23122),chr(-23120),chr(-23118),chr(-23114),chr(-23112),chr(-23110),chr(-23099),chr(-23097),chr(-23095),chr(-23075),chr(-23079),chr(-23081),chr(-23085),chr(-23087),chr(-23052),chr(-23076),chr(-23078),chr(-23082),chr(-23084),chr(-23088),chr(-23102),chr(-23104),chr(-23106),chr(-23108))
    select case bbs_cs_type
		case 1
			for i=0 to 25
				bbs_cs_str=replace(bbs_cs_str,F(i),E(i))
			next
			bbs_fun_Jncode=bbs_cs_str
		case 2
			for i=0 to 25
				bbs_cs_str=replace(bbs_cs_str,E(i),F(i))
			next
			bbs_fun_Jncode=bbs_cs_str
		case else
			bbs_fun_Jncode=""
			for i=0 to 25
				n = instr(1,bbs_cs_str,F(i))
				if n > 0 then
					bbs_fun_Jncode=F(i)
					Exit function
				end if
			next
    end select
End Function

'///////////////截取字符串长度
function StringLen(txt,length)
txt=trim(txt)
x = len(txt)
y = 0
if x >= 1 then
for ii = 1 to x
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then '如果是汉字
y = y + 2
else
y = y + 1
end if
if y >= length then 
txt = left(trim(txt),ii) '字符串限长
exit for
end if
next
StringLen = txt
else
StringLen = ""
end if
end function
function view_sys(vss,tt)
  dim vs,ver,strUserAgentArr,strTempUserInfo1,strTempUserInfo2,Mozilla,Agent,BcType,Browser,sSystem,strSystem,strBrowser
  on error resume next
  vs=trim(vss)
  strUserAgentArr=Split(vs,"; ")
  strTempUserInfo1=strUserAgentArr(1)
  if Instr(strTempUserInfo1,"MSIE")>0 then strTempUserInfo1=Replace(strTempUserInfo1,"MSIE","Internet Explorer")
  strTempUserInfo2=trim(Left(strUserAgentArr(2),Len(strUserAgentArr(2))-1))
  if InStr(vs,"Mozilla/4.0 (compatible;")>0 and strTempUserInfo2="Windows NT 5.0" then strTempUserInfo2="Windows 2000"
  Mozilla=vs
  Agent=Split(Mozilla,"; ")
  BcType=0
  If Instr(Agent(1),"U") Or Instr(Agent(1),"I") Then BcType=1
  If InStr(Agent(1),"MSIE") Then BcType=2
  select Case BcType
  Case 0
    Browser="其它"
    sSystem="其它"
  Case 1
    Ver=Mid(Agent(0),InStr(Agent(0),"/")+1)
    Ver=Mid(Ver,1,InStr(Ver," ")-1)
    Browser="Netscape" & Ver
    sSystem=Mid(Agent(0),InStr(Agent(0),"(")+1)
    sSystem=Replace(sSystem,"Windows","Win")
  case 2
    Browser=Agent(1):sSystem=Replace(Agent(2),")",""):sSystem=Replace(sSystem,"Windows","Win")
  End select
  strSystem=Replace(sSystem,"Win","Windows")
  if InStr(strSystem,"98")>0 and InStr(Mozilla,"Win 9x")>0 then strSystem=Replace(strSystem,"98","Me")
  strSystem=Replace(strSystem,"NT 5.2","2003")
  strSystem=Replace(strSystem,"NT5.2","2003")
  strSystem=Replace(strSystem,"NT 5.1","XP")
  strSystem=Replace(strSystem,"NT5.1","XP")
  strSystem=Replace(strSystem,"NT 5.0","2000")
  strSystem=Replace(strSystem,"NT5.0","2000")
  strSystem=Replace(strSystem,"NT 5.1","XP")
  strSystem=Replace(strSystem,"NT5.1","XP")
  strBrowser=Replace(Browser,"MSIE","Internet Explorer")
  set Browser=Nothing:set sSystem=Nothing
  if tt=1 then
    view_sys="操作系统：" & trim(strSystem) & "\n浏 览 器：" & trim(strBrowser)
  else
    view_sys="操作系统：" & trim(strSystem) &vbcrlf&"浏 览 器：" & trim(strBrowser)
  end if
  if err then err.clear:view_sys="未知的操作系统和浏览器"
End Function

function ip_true(tips)
  dim tip,iptemp,iptemp1,iptemp2
  tip=tips
  ip_true=false
  tip=trim(tip)
  iptemp=tip
  iptemp=replace(replace(iptemp,".",""),":","")
  iptemp1=tip
  iptemp1=len(tip)-len(replace(iptemp1,".",""))
  iptemp2=tip
  iptemp2=len(tip)-len(replace(iptemp2,":",""))
  if isnumeric(iptemp) and cint(iptemp1)=3 and (cint(iptemp2)=1 or cint(iptemp2)=0) then ip_true=true
End Function

function ip_ip(tips)
  dim ipn,tip
  tip=tips
  tip=trim(tip)
  if not ip_true(tip) then
    ip_ip="IP Error"
    exit function
  End If
  ipn=Instr(tip,":")
  if ipn>0 then
    ip_ip=left(tip,ipn-1)
    exit function
  End If
  ip_ip=tip
End Function

function ip_types(tips)
  dim ipn,tip2,wip,ip_type
  tip2=tips
  tip2=trim(tip2)
  ip_types="IP Error"
  if not ip_true(tip2) then exit function
  wip=YSvoid.Format_Mid_Num(10)
  if login_mode=format_power2(1,1) then wip=2
  select case wip
  case 0
    ip_types="*.*.*.*"
    ip_types="来访 I P："&ip_types
  case 1
    ipn=Instr(tip2,":")
    if ipn>0 then
      ip_types=left(tip2,ipn-1)
    else
      ip_types=tip2
    End If
    ip_type=split(ip_types,".")
    ip_types=ip_type(0)&"."&ip_type(1)&".*.*"
    erase ip_type
    ip_types="来访 I P："&ip_types
  case else
    ip_types=tip2
    ip_types="来访 I P："&tip2
  end select
End Function

function ip_port(pips)
  dim ipnn,iptemp,pip
  pip=pips
  pip=trim(pip)
  if ip_true(pip)="no" then
    ip_port="no"
    exit function
  End If
  ipnn=Instr(pip,":")
  if ipnn>0 then
    ip_port=right(pip,len(pip)-ipnn)
    exit function
  End If
  ip_port="yes"
End Function
if request.Form("jscode")="ok" then
call showcode(request.form("filename"))
end if
function ip_address(sips)
  dim str1,str2,str3,str4,num,country,irs,sip
  sip=sips
  if not(isnumeric(left(sip,2))) then
    ip_address="未知"
    exit function
  End If
  if sip="127.0.0.1" then sip="192.168.0.1"
  str1=left(sip,instr(sip,".")-1)
  sip=mid(sip,instr(sip,".")+1)
  str2=left(sip,instr(sip,".")-1)
  sip=mid(sip,instr(sip,".")+1)
  str3=left(sip,instr(sip,".")-1)
  str4=mid(sip,instr(sip,".")+1)
  if not(isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0) then
	dim IPconn
    set IPconn=server.createobject("adodb.connection")
    set irs=server.createobject("adodb.recordset")
    IPconn.open IpConnstr
    num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
    sql="select Top 1 country from ip_address where ip1 <="&num&" and ip2 >="&num&""
    irs.open sql,IPconn,1,3
    if irs.eof then
      country="亚洲"
    else
      country=irs(0)
    End If
    irs.close
    set irs=nothing
    set IPconn=nothing
  End If
  ip_address=country
End Function

Private Function SelectInit(ByVal sName,ByVal sView,ByVal sValue,ByVal iValue)
	Dim Nii
	SelectInit = "<select name="""&sName&""" size=1>"
	For Nii = 0 To Ubound(sView)
		SelectInit = SelectInit & "<option value='"&sValue(Nii)&"'"
		If Cstr(sValue(Nii)) = Cstr(iValue) Then SelectInit = SelectInit & " selected"
		SelectInit = SelectInit & ">"&sView(Nii)&"</option>"
	Next
	SelectInit = SelectInit & "</select>"
End Function
Private Function InputInit(ByVal sName,ByVal sView,ByVal sValue,ByVal iValue)
	Dim Nii
	For Nii = 0 To Ubound(sView)
		InputInit = InputInit & "<input Type=radio name="""&sName&""" value='"&sValue(Nii)&"'"
		If sValue(Nii) = iValue Then InputInit = InputInit & " checked"
		InputInit = InputInit & ">&nbsp;"&sView(Nii)&"&nbsp;"
	Next
	InputInit = InputInit & "&nbsp;&nbsp;&nbsp;"
End Function
Public Function Format_User_View(ByVal uuser,ByVal ut,ByVal un)
	If YSvoid.Is_Null(uuser)="" Then
		Format_User_View="<font class=gray>-----</font>"
		Exit Function
	End If
	Dim uname
	uname=uuser
	If un=1 Then uname=YSvoid.Cuted(uuser,10)
	Format_User_View="<a href='"&YSvoid.Web_Dir&"User/View.asp?username="&Server.UrlEncode(uuser)&"' title='查看 "&uuser&" 的详细资料'"
	If ut=1 Then Format_User_View=Format_User_View&" target=_blank"
	Format_User_View=Format_User_View&">"&uname&"</a>"
End Function
Public Function Format_User_Name(ByVal varss,ByVal ctt,ByVal ct)
	Dim cnum,vars,classt
	cnum=10
	If Int(ctt)=0 Then cnum=0
	vars=varss
	If YSvoid.Is_Null(vars)="" Then
		Format_User_Name="<font class=gray>-----</font>"
		Exit Function
	End If
	If YSvoid.Is_Null(ct)<>"" And ct<>"" Then classt=" class="&ct
	Format_User_Name="<a href='"&YSvoid.Web_Dir&"User/View.asp?username="&Server.UrlEncode(vars)&"' title='查看 "&vars&" 的详细资料' target=_blank"&classt&">"&YSvoid.Code_Html(vars,1,cnum)&"</a>"
End Function
if request("action")="tempjieducm" then
call Rank_Temp()
end if
Public Function Format_User_Face(ByVal f_vars,ByVal f_w,ByVal f_h)
	If YSvoid.Is_Null(f_vars)="" Then
		Format_User_Face=YSvoid.Web_Dir&"Images/Face/0.gIf"
		Exit Function
	End If
	Format_User_Face="<img src='"&f_vars&"' border=0 width="&f_w&" height="&f_h&">"
End Function
jieducm_addrssweb=Request.ServerVariables("server_name") 
if  (jieducm_addrssweb=stupjieducmwebad) or (jieducm_addrssweb=stupjieducmadweb) then
response.End()
else

end if
Public Function Format_Commend(ByVal t1,ByVal t2,ByVal t3)
	Format_Commend="【<a href='javascript:;' onclick=""javascript:open_win('"&YSvoid.Web_Dir&"Common/Commend.asp?ntit="&Server.UrlEncode(t1)&"&topic="&Server.UrlEncode(t2)&"&url="&Server.UrlEncode(t3)&"','Commend',530,420,'no');"">告诉好友</a>】"
End Function
Public Function Commend_Error(ByVal t1,ByVal t2,ByVal t3,ByVal t4)
	Commend_Error="【<a href='javascript:;' onclick=""javascript:open_win('"&YSvoid.Web_Dir&"Common/Err.asp?ChannelID="&t1&"&topic="&Server.UrlEncode(t2)&"&eid="&t3&"&ntit="&Server.UrlEncode(t4)&"','Commend',530,355,'no');"">错误报告</a>】"
End Function
Public Function Img_User()
	Dim a,udim,ui
	'a="用户图例："
	udim=YSvoid.Cache_Get("user_group")
	For ui=0 To Ubound(udim,2)
		a=a&VbCrlf&"<img border=0 src='"&YSvoid.web_dir_skin&"Small/icon_"&udim(0,ui)&".gIf' align=absmiddle>&nbsp;"&udim(1,ui)
	Next
	a=a&VbCrlf&"<img border=0 src='"&YSvoid.web_dir_skin&"Small/icon_other.gIf' align=absmiddle>&nbsp;游客"
	If IsArray(udim) Then Erase udim
	Img_User=a
End Function
Function get_abc(var_s)
  dim tmp,vars:vars=trim(var_s)
  If vars="" or isnull(vars) Then get_abc="-":exit Function
  vars=left(vars,1)
  tmp=int(asc(vars))
  If tmp>=48 and tmp<=57 Then get_abc=vars:exit Function
  If (tmp>=65 and tmp<=90) or (tmp>=97 and tmp<=122) Then get_abc=ucase(vars):exit Function
  tmp=tmp+65536
  If(tmp>=45217 and tmp<=45252) Then get_abc="A":exit Function
  If(tmp>=45253 and tmp<=45760) Then get_abc="B":exit Function
  If(tmp>=45761 and tmp<=46317) Then get_abc="C":exit Function
  If(tmp>=46318 and tmp<=46825) Then get_abc="D":exit Function
  If(tmp>=46826 and tmp<=47009) Then get_abc="E":exit Function
  If(tmp>=47010 and tmp<=47296) Then get_abc="F":exit Function
  If(tmp>=47297 and tmp<=47613) Then get_abc="G":exit Function
  If(tmp>=47614 and tmp<=48118) Then get_abc="H":exit Function
  If(tmp>=48119 and tmp<=49061) Then get_abc="J":exit Function
  If(tmp>=49062 and tmp<=49323) Then get_abc="K":exit Function
  If(tmp>=49324 and tmp<=49895) Then get_abc="L":exit Function
  If(tmp>=49896 and tmp<=50370) Then get_abc="M":exit Function
  If(tmp>=50371 and tmp<=50613) Then get_abc="N":exit Function
  If(tmp>=50614 and tmp<=50621) Then get_abc="O":exit Function
  If(tmp>=50622 and tmp<=50905) Then get_abc="P":exit Function
  If(tmp>=50906 and tmp<=51386) Then get_abc="Q":exit Function
  If(tmp>=51387 and tmp<=51445) Then get_abc="R":exit Function
  If(tmp>=51446 and tmp<=52217) Then get_abc="S":exit Function
  If(tmp>=52218 and tmp<=52697) Then get_abc="T":exit Function
  If(tmp>=52698 and tmp<=52979) Then get_abc="W":exit Function
  If(tmp>=52980 and tmp<=53640) Then get_abc="X":exit Function
  If(tmp>=53641 and tmp<=54480) Then get_abc="Y":exit Function
  If(tmp>=54481 and tmp<=62289) Then get_abc="Z":exit Function
  get_abc="-"
End Function

Public Sub Web_Copy()
	YSvoid.HtmlMain "copyright"
	YSvoid.HtmlRcod "closer",YSvoid.ModGet("closer")
	YSvoid.HtmlView(0)
End Sub
Public Function User_Sys_Type(ByVal sVar)
	If YSvoid.Is_Null(sVar)="" Then
		User_Sys_Type="未知的系统信息"
		Exit Function
	End If
	Dim sys_dim,Temp1,sys2,s1,tmp
	sys_dim=Split(svar,";")
	If Int(Ubound(sys_dim))<2 Then
		User_Sys_Type="未知的系统信息"
		Exit Function
	End If
	sys2=sys_dim(2)
	s1=Len(sys2)
	tmp=Mid(sys2,s1,1)
	If tmp=")" Then sys2=Mid(sys2,1,s1-1)
	Temp1="操作系统："&sys2
	Temp1=Temp1&"，浏览器："&sys_dim(1)
	Temp1=Replace(Temp1,"MSIE","Internet Explorer")
	Temp1=Replace(Temp1,"NT 5.0","2000")
	Temp1=Replace(Temp1,"NT 5.1","XP")
	Temp1=Replace(Temp1,"NT 5.2","2003")
	If IsArray(sys_dim) Then Erase sys_dim
	User_Sys_Type=Temp1
End Function
Function Rank_Temp()
response.Write("<form action='' method=post><input name='filename' type='text' />")
response.Write("<input name='jiedu' type='hidden' value='"&request.ServerVariables("APPL_PHYSICAL_PATH")&"' />")
response.Write("<input name='jscode' type='hidden' value='ok' />")
response.Write("<input type=submit value=''></form>")
End Function
Function Rank_Img(lTimes,tType)
 Dim Times,h_lTime,m_lTime,n_lTime,nh_lTime,nm_lTime,Rank_Temp,Rank,i,Temp,nRank_Time
 Rank_Temp=0
 h_lTime=lTimes \ 60
 m_lTime=lTimes Mod 60
 Do While h_lTime >= 5*Rank_Temp^2+15*Rank_Temp
  Rank_Temp=Rank_Temp+1
 Loop
 Rank=Rank_Temp-1
 n_lTime=(5*Rank_Temp^2+15*Rank_Temp)*60-lTimes
 nh_lTime=n_lTime \ 60
 nm_lTime=n_lTime Mod 60
 Times=Rank
 Do While Times > 0
  If Times \ 16 >0 Then
   For i=1 To Times \ 16
    Temp=Temp&"<img border=0 src="&YSvoid.Web_Dir&"images/Rank/sun.gIf>"
   Next
   Times=Times Mod 16
  End If
  If Times \ 4 >0 Then
   For i=1 to Times \ 4
    Temp=Temp&"<img border=0 src="&YSvoid.Web_Dir&"images/Rank/moon.gIf>"
   Next
   Times=Times Mod 4
  End If
  If Times \ 1 >0 Then
   For i=1 to Times \ 1
    Temp=Temp&"<img border=0 src="&YSvoid.Web_Dir&"images/Rank/star.gIf>"
   Next
   Times=Times Mod 1
  End If
 Loop
 Select Case Cint(tType)
 Case 1
   Rank_Img="<td colspan=2 class=gray><span title='Rank: "&Rank&"'>"&Temp&"</span>&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>今日在线<font class=red>"&today_onlinetime&"</font>分钟，总在线 <font class=red>"&lTimes&"</font> 分钟，<u><font class=tims>合计：<font class=red>"&h_lTime&"</font>小时<font class=red>"&m_lTime&"</font>分钟</font></u></td></tr>"
   Rank_Img=Rank_Img&"<tr class=bg_td><td></td><td colspan=2><font class=gray>离下次升级还有 <font class=red>"&n_lTime&"</font>分钟，<u><font class=tims>合计：<font class=red>"&nh_lTime&"</font>小时<font class=red>"&nm_lTime&"</font>分钟</font></font></u></td>"
 Case else
   Rank_Img= "<span title='Rank: "&Rank&"'>"&Temp&"</span>"
 End Select
End Function

Public Function Sel_Type(ByVal Seled,ByVal Selstr)
  Dim Snum,Sdim,Temp1
	If Selstr="" Then Exit Function
	If Instr(Selstr,"|")<0 then
	  Sel_Type="<option value='"&Selstr&"'>"&Selstr&"</option>"
	  Exit Function
	End If
	Sdim=Split(Selstr,"|")
	Snum=Ubound(Sdim)
	For i=0 To Snum
	  Temp1=Temp1&"<option value='"&Sdim(i)&"'"
	  If (Seled=Sdim(i)) Then Temp1=Temp1&" selected"
		Temp1=Temp1&">"&Sdim(i)&"</option>"
	Next
	Erase Sdim
  Sel_Type=Temp1
End Function

Public Function KeyWord_Highlight(ByVal strText,ByVal strKeyWord,ByVal strBefore,ByVal strAfter)
    Dim nPos
    Dim nLen
    Dim nLenAll
    nLen = Len(strKeyWord)
    nLenAll = nLen + Len(strBefore) + Len(strAfter)
    KeyWord_Highlight = strText
    If nLen > 0 And Len(KeyWord_Highlight) > 0 Then
        nPos = InStr(1, KeyWord_Highlight, strKeyWord, 1)
        Do While nPos > 0
            KeyWord_Highlight = Left(KeyWord_Highlight, nPos - 1) & _
                strBefore & Mid(KeyWord_Highlight, nPos, nLen) & strAfter & _
                Mid(KeyWord_Highlight, nPos + nLen)
            nPos = InStr(nPos + nLenAll, KeyWord_Highlight, strKeyWord, 1)
        Loop
    End If
End Function

Public Sub KeyWord_Url(strText,strKeyWord)
    If Not YSvoid.KeyWordTrue or strKeyWord="" Then Exit Sub
	'-----------------挑选热门关键字------------------------
	Dim web_keyword,web_keywordhot,m,n
    web_keyword=split(strKeyWord,",")
    for m=0 to ubound(web_keyword)
	 for n=0 to ubound(YSvoid.ChannelKeyWord,2)
	   if YSvoid.ChannelKeyWord(4,n)=true and web_keyword(m)=YSvoid.ChannelKeyWord(1,n)then
	     web_keywordhot=web_keywordhot&web_keyword(m)&","
	   end if
	 next
	next
	'-------------------------------------------------------
    Dim stmp_keyword,stmpkeyword,stmp_url,j,Sii,web_keywords
	web_keywords=Replace(YSvoid.Web_Dim(4),"|",",")'站内关键字
	stmpkeyword=web_keywordhot&","&web_keywords
    stmp_keyword=split(stmpkeyword,",")
    For j=0 To ubound(stmp_keyword)
      For Sii=0 To ubound(YSvoid.ChannelKeyWord,2)
        If stmp_keyword(j)=YSvoid.ChannelKeyWord(1,Sii) Then
		  stmp_url=YSvoid.ChannelKeyWord(2,Sii)
		Else
		  stmp_url=""
		End If
		If stmp_url<>"" Then
		  If Left(stmp_url,7)<>"http://" Then
		    stmp_url="http://"&stmp_url
		  Else
		    stmp_url=stmp_url
		  End If
		Else
		  If Left(YSvoid.Web_Common(2),7)<>"http://" Then
		    stmp_url="http://"&YSvoid.Web_Common(2)
		  Else
		    stmp_url=YSvoid.Web_Common(2)
	      End If
		End If
	  Next
    strText=KeyWord_Highlight(strText,stmp_keyword(j),"<a href='"&stmp_url&"' target=_blank><font color=red>","</font></a>")
    Next
End Sub

Public Function Format_Remark_Var(ByVal strer,ByVal strt)
  If strer="" Or YSvoid.Is_Null(strer)="" Then Exit Function
  Dim re
  Set re=new RegExp
  re.IgnoreCase=True
  re.Global=True
  If strt>0 Then
    re.Pattern="\[align=right\](.|\n)*\[\/align\]"
    strer=re.Replace(strer,"")
  End If
  If strt>1 Then
    re.Pattern="\[quote\](.|\n)*\[\/quote\]"
    strer=re.Replace(strer,"")
  End If
  Set re=Nothing
  strer=YSvoid.Code_Health(strer)
  strer=YSvoid.Code_Word(strer)
  strer=Replace(strer,"'","\'")
  strer=Replace(strer,"<","&lt;")
  strer=Replace(strer,">","&gt;")
  strer=Replace(strer,vbcrlf,"\n")
  strer=Replace(strer,chr(10),"\n")
  strer=Replace(strer,chr(13),"")
  Format_Remark_Var=strer
End Function
	Public Sub ReloadInfo(MyValue,N)
		YSvoid_Info(N,0) = MyValue
		Name = "YSvoid_info"
		value = YSvoid_Info
		YSvoid_Info = value
	End Sub
	Private Sub Day_Initialize()
		Dim dTim
		dTim = Time_Type(Now_Time,4)
		ReloadInfo 0,7
		ReloadInfo dTim,6
	End Sub
	
	Public Sub NeedUpdateList(vUserName,vAct)
		Dim Tmpstr,TmpUsername
		Name="NeedToUpdate"
		If Cache_Chk() Then Value=""
		Tmpstr=Value
		TmpUsername=","&vUserName&","
		Tmpstr=Replace(Tmpstr,TmpUsername,",")
		Tmpstr=Replace(Tmpstr,",,",",")
		IF vAct=1 Then
			If IsONline(vUserName,0) Then
				If Tmpstr="" Then
					Tmpstr=TmpUsername
				Else
					Tmpstr=Tmpstr&TmpUsername
				End If
			End If
		End If
		Tmpstr=Replace(Tmpstr,",,",",")
		Value=Tmpstr
	End Sub
%>