<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
dim From_url,Serv_url
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "����ʱ,��رձ�ҳ�沢���µ�½����!"
response.end
end If
response.contenttype = "text/html"
response.expires = 0 
response.expiresabsolute = now() - 1 
response.addHeader "pragma","no-cache" 
response.addHeader "cache-control","private" 
response.cachecontrol = "no-cache"
Response.ContentType = "text/html; charset=gb2312"
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
On Error Resume Next 
Server.ScriptTimeOut=9999999

if request.Form("fa")="ok" then				
Dim Price(6)
Dim ProUrl(6)
Price(1)=request.Form("Price1")
ProUrl(1)=Request.Form("ProUrl1")
Price(2)=request.Form("Price2")
ProUrl(2)=Request.Form("ProUrl2")
Price(3)=request.Form("Price3")
ProUrl(3)=Request.Form("ProUrl3")
Price(4)=request.Form("Price4")
ProUrl(4)=Request.Form("ProUrl4")
Price(5)=request.Form("Price5")
ProUrl(5)=Request.Form("ProUrl5")
Shopkeeper=Replace(Trim(request.Form("Shopkeeper")),"'","''")
bei=Request.Form("bei")
baohu2=request.Form("baohu2")
isprice=request.form("isprice")
num=request.Form("num")
depot=request.form("depot")
tips=Replace(Trim(request.Form("tips")),"'","''")
play=request.Form("play")
Limit=request.Form("Limit")
title=Replace(Trim(Request.Form("title")),"'","''")
addfabu=request.Form("addfabu")
ping=Replace(Trim(Request.Form("ping")),"'","''")
pingtxt=Replace(Trim(Request.Form("pingtxt")),"'","''")
outtime=Replace(Trim(Request.Form("Timing")),"'","''")
jieducm_IfComeSet=request.Form("jieducm_IfComeSet")
IfComeSet=request.Form("IfComeSet")
ComeKey=Replace(Trim(Request.Form("ComeKey")),"'","''")
ComeNote=Replace(Trim(Request.Form("ComeNote")),"'","''")
if outtime="" then
outtime=now
end if
fabu2=0
k=0
if num="" then
	Response.Write("<script>alert('����ѡ��������������!');history.back();</script>")
    response.End()
elseif  instr (ProUrl(1),"shop") then
 Response.Write("<script>alert('����д��ȷ������ַ��');history.back();</script>")
 response.End()
elseif Price(1)="" then
	Response.Write("<script>alert('����1��Ʒ����۲���Ϊ��!');history.back();</script>")
    response.End()
elseif (Price(1))<0.3  or (Price(0.3))>1000 then
    Response.Write("<script>alert('Ŀǰ��ƽֻ̨֧�ַ���0.3-1000Ԫ������ ������������!');history.back();</script>")
    response.End()		
elseif Shopkeeper="" then
	Response.Write("<script>alert('�㻹û��ѡ���ƹ���!');history.back();</script>")
    response.End()	
elseif ProUrl(1)="" then
	Response.Write("<script>alert('����1��Ʒ��ַ����Ϊ��!');history.back();</script>")
    response.End()
elseif bei="" then
	Response.Write("<script>alert('�㻹û��ѡ��ע!');history.back();</script>")
    response.End()
elseif not isNumeric(addfabu) then
  Response.Write("<script>alert('�����������������룡');history.back();</script>")
  response.End()
elseif int(addfabu)<0 then
  Response.Write("<script>alert('�����������������룡');history.back();</script>")
  response.End()
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='1' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx+num)>int(stupfaweiok+5)  then
	Response.Write("<script>alert('��VIP��Աֻ�ܷ���"&stupfaweiok+5&"��δ��ɵ�����\n��������ѷ�������!���ܼ���������');history.back();</script>")
    response.End()
end if 

for i=1 to num
 if Price(i)="" then
	Response.Write("<script>alert('����"&i&"����Ʒ�۲���Ϊ��!');history.back();</script>")
    response.End()
 elseif ProUrl(i)="" then
	Response.Write("<script>alert('����"&i&"����Ʒ��ַ����Ϊ��!');history.back();</script>")
    response.End()
 elseif (Price(i))<0.3  or (Price(i))>1000 then
    Response.Write("<script>alert('Ŀǰ��ƽֻ̨֧�ַ���0.3-1000Ԫ������ ������������!');history.back();</script>")
    response.End()
elseif  instr (ProUrl(i),"meal") and addfabu=0 then
 Response.Write("<script>alert('�����ײ�����ı����������ӷ����㣡');history.back();</script>")
 response.End()
end if
 
url= ProUrl(i)
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" href=""[^""]+""[^>]*>(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk=Matches(0).SubMatches(0)
 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" title=""(.*?)"""
set Matches=objRegExp.Execute(wstr)
jieducm_sk3=Matches(0).SubMatches(0)


if Price(i)>=0.3 and Price(i)<=40 then
   fabu=2
elseif Price(i)>41 and Price(i)<=100 then
   fabu=3
elseif Price(i)>101 and Price(i)<=200 then
   fabu=4
elseif Price(i)>201 and Price(i)<=500 then
  fabu=5
elseif Price(i)>501 and Price(i)<=1000 then
  fabu=6
end if 

if bei="һ����ջ�����" then
fabu=fabu*1+1
elseif bei="������ջ�����"  then
fabu=fabu*1+2
elseif bei="������ջ�����"  then
fabu=fabu*1+3
elseif bei="������ջ�����"  then
fabu=fabu*1+4
elseif bei="������ջ�����"  then
fabu=fabu*1+5
end if
if jieducm_IfComeSet=1 then
fabu=fabu+1
end if
fabu2=int(fabu2)+int(fabu)+int(addfabu)
k=k+Price(i) 
Next 

jieducm_ck=cunkuan*100
jieducm_price=k*100

fabudian2=fabudian*100
fabu3=fabu2*100

if isfa=1 then
    Response.Write("<script>alert('���ѱ�����Ա��ֹ�˷�������!');history.back();</script>")
    response.End()
elseif int(jieducm_ck)< int(jieducm_price) then
 Response.Write("<script>alert('�㷢���������Ǯ���ܴ�����Ĵ��!');history.back();</script>")
 response.End()
elseif(int(fabudian2)<int(fabu3)) then
	Response.Write("<script>alert('��ķ�������������,�뾡���ֵ��ȡ������!');history.back();</script>")
	response.End()	
end if

for i=1 to num				 
randomize
if Price(i)>=0.3 and Price(i)<=40 then
   fabu=2
elseif Price(i)>40 and Price(i)<=100 then
   fabu=3
elseif Price(i)>100 and Price(i)<=200 then
   fabu=4
elseif Price(i)>200 and Price(i)<=500 then
  fabu=5
elseif Price(i)>500 and Price(i)<=1000 then
  fabu=6
end if 

if bei="һ����ջ�����" then
fabu=fabu*1+1
elseif bei="������ջ�����"  then
fabu=fabu*1+2
elseif bei="������ջ�����"  then
fabu=fabu*1+3
elseif bei="������ջ�����"  then
fabu=fabu*1+4
elseif bei="������ջ�����"  then
fabu=fabu*1+5
end if

if jieducm_IfComeSet=1 then
fabu=fabu+1
end if

ranNum=int(90000*rnd)+10000	
pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin where pid='"&pid&"'" ,Conn,3,3  
if  rs.eof then
rs.addnew
rs("classid")=1
rs("Price")=Price(i)
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl(i)
rs("bei")=bei	
if baohu2="" then
rs("baohu2")=0
else
rs("baohu2")=baohu2
end if
rs("username")=session("useridname")
rs("pid")=pid
rs("now")=outtime
rs("isprice")=isprice
rs("play")=play
rs("tips")=tips
rs("Limit")=Limit
rs("addfabu")=addfabu
rs("ping")=ping
rs("pingtxt")=pingtxt
rs("jieducm_fb")=fabu
if jieducm_IfComeSet=1 then
rs("IfComeSet")=IfComeSet
rs("ComeKey")=ComeKey
rs("ComeNote")=ComeNote
end if
rs.update
rs.close
else
 Response.Write("<script>alert('�����ˣ��뷵�����·���!');history.back();</script>")
  response.End()
end if
				
if depot=1 then
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_depot " ,Conn,3,3  
rs.addnew
rs("classid")=1
rs("Price")=Price(i)
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl(i)
rs("bei")=bei
if baohu2="" then
rs("baohu2")=0
else
rs("baohu2")=baohu2
end if
rs("username")=session("useridname")
rs("now")=now
rs("isprice")=isprice
rs("pid")=pid
rs("num")=1
rs("play")=play
rs("tips")=tips
rs("company")=company
rs("numbers")=numbers
rs("Limit")=Limit
rs("title")=title
rs("addfabu")=addfabu
rs("ping")=ping
rs("pingtxt")=pingtxt
rs("jieducm_fb")=fabu
if jieducm_IfComeSet=1 then
rs("IfComeSet")=IfComeSet
rs("ComeKey")=ComeKey
rs("ComeNote")=ComeNote
end if
rs.update
rs.close
end if

Next  
 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
rs("fabudian")=rs("fabudian")-fabu2
rs("cunkuan")=rs("cunkuan")-k
rs.update
rs.close
call check_jieducm_name(session("useridname"))
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record" ,Conn,3,3  
rs.addnew
rs("username")=session("useridname")
rs("num")=pid
rs("class")="�Ա�VIP��������"
rs("content")=session("useridname")&"����vip��������ɹ�,������"&k&"Ԫ�������������"&fabu2&"��"
rs("price")="-"&k
rs("jiegou")="�ɹ�"
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing		
Response.Write("<script>alert('�����ɹ�!�ȴ������ˣ�');window.location.href='MyMission.asp';</script>")	
response.End()
end if
%>