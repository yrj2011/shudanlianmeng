<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
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
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
On Error Resume Next 
Server.ScriptTimeOut=9999999 
if request.Form("fa")="ok" then
Price=Request.Form("Price1")
ProUrl=Replace(Trim(Request.Form("ProUrl1")),"'","''")
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
bei=Replace(Trim(Request.Form("bei")),"'","''")
baohu2=request.Form("baohu2")
isprice=request.form("isprice")
depot=request.form("depot")
play=request.form("play")
tips=Replace(Trim(Request.Form("tips")),"'","''")
Limit=request.Form("Limit")
title=Replace(Trim(Request.Form("title")),"'","''")
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
if  instr (ProUrl,"shop") then
 Response.Write("<script>alert('����д��ȷ������ַ��');history.back();</script>")
 response.End()
elseif  instr (ProUrl,"meal") then
 Response.Write("<script>alert('ֻ��VIP��Ա�ſ��Է����ײ�����ı�����');history.back();</script>")
 response.End()
elseif cunkuan<0.1 then
 Response.Write("<script>alert('������0.1Ԫʱ�����ٷ�����');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif fabudian<1 then
 Response.Write("<script>alert('���������1��ʱ�����ٷ�����');window.location.href='../user/md5_pay.asp';</script>")
 response.End() 
elseif price="" then
	Response.Write("<script>alert('��Ʒ����۲���Ϊ��!');history.back();</script>")
    response.End()
elseif int(price)<0.1  or int(price)>1000 then
	Response.Write("<script>alert('Ŀǰ��ƽֻ̨֧�ַ���0.1-1000Ԫ������ ������������!');history.back();</script>")
    response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('�㻹û��ѡ���ƹ���!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('��Ʒ��ַ����Ϊ��!');history.back();</script>")
    response.End()
elseif ping="" then
	Response.Write("<script>alert('�㻹û��ѡ�������ʽ!');history.back();</script>")
    response.End()
elseif ping="�Զ�������"  and pingtxt="" then
	Response.Write("<script>alert('�������Զ�������!');history.back();</script>")
    response.End()
elseif bei="" then
	Response.Write("<script>alert('�㻹û��ѡ��ע!');history.back();</script>")
    response.End()
elseif not isNumeric(price) then
  Response.Write("<script>alert('�����������������룡');history.back();</script>")
  response.End()
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='1' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(stupfaweiok) and vip<>"1" then
	Response.Write("<script>alert('������ͨ��Աֻ�ܷ���"&stupfaweiok&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupfaweiok+5) and vip="1" then
	Response.Write("<script>alert('��VIP��Աֻ�ܷ���"&stupfaweiok+5&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
end if 

url= ProUrl
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

								
if price>=0.1 and price<=40 then
   fabu=1
elseif price>40 and price<=100 then
   fabu=2
elseif price>100 and price<=200 then
   fabu=3
elseif price>200 and price<=500 then
  fabu=4
elseif price>500 and price<=1000 then
  fabu=5
end if 

if bei="һ����ջ�����" then
fabu=fabu*2
elseif bei="������ջ�����"  then
fabu=fabu*2+1
elseif bei="������ջ�����"  then
fabu=fabu*2+2
end if
if jieducm_IfComeSet=1 then
fabu=fabu+1
end if
jieducm_ck=cunkuan*100
jieducm_price=price*100
fabudian2=fabudian*100
fabu2=fabu*100
if int(jieducm_ck)< int(jieducm_price) then
	 Response.Write("<script>alert('�㷢���������Ǯ���ܴ�����Ĵ��!');history.back();</script>")
	 response.End()
elseif(int(fabudian2)<int(fabu2)) then
	Response.Write("<script>alert('��ķ�������������,�뾡���ֵ��ȡ������!');history.back();</script>")
	response.End()  
elseif isfa="1" then
	Response.Write("<script>alert('���ѱ�����Ա��ֹ�˷�������!');history.back();</script>")
	response.End()
end if
          
randomize	
ranNum=int(90000*rnd)+10000	
pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin  where pid='"&pid&"'" ,Conn,3,3
if  rs.eof then  
rs.addnew
rs("classid")=1
rs("Price")=Price
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
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
rs("Price")=Price
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
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

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
rs("cunkuan")=rs("cunkuan")-price
rs("fabudian")=rs("fabudian")-fabu
rs.update
rs.close
call check_jieducm_name(session("useridname"))	
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs.addnew
rs("username")=session("useridname")
rs("num")=pid
rs("class")="�Ա�������"
rs("content")=session("useridname")&"��������"&pid&"�ɹ�,������"&price&"Ԫ�������������"&fabu&"��"
rs("price")="-"&price
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