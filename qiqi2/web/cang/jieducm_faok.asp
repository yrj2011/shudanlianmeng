<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
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
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
info=Replace(Trim(Request.Form("info[remarks]")),"'","''")
ProUrl=Replace(Trim(Request.Form("ProUrl")),"'","''")
product=Replace(Trim(Request.Form("product[number]")),"'","''")
depot=request.form("depot")
title=Replace(Trim(Request.Form("title")),"'","''")
tips=Replace(Trim(Request.Form("tips")),"'","''")

if  instr(ProUrl,"shop") and info="�����ղ�����" then
 Response.Write("<script>alert('����д��ȷ������ַ��');history.back();</script>")
 response.End()
elseif  instr(ProUrl,"shop")<>8 and info="�����ղ�����" then
 Response.Write("<script>alert('����д��ȷ���̵�ַ��');history.back();</script>")
 response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('�㻹û��ѡ���ƹ���!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('��������̻򱦱���ַ!');history.back();</script>")
    response.End()
elseif product="" then
	Response.Write("<script>alert('�ղ���������Ϊ��!');history.back();</script>")
    response.End()
elseif product<1 then
	Response.Write("<script>alert('�ղ���������С��1!');history.back();</script>")
    response.End()
elseif not isNumeric(product) then
  Response.Write("<script>alert('�����������������룡');history.back();</script>")
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


Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a href="".*?"" class=""hCard fn"">(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk2=Matches(0).SubMatches(0)

fabu=product*stupcang
fabudian2=fabudian*100
fabu2=fabu*100
if(int(fabudian2)<int(fabu2)) then
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
rs("classid")=6
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
rs("bei")=info
rs("username")=session("useridname")
rs("pid")=pid
rs("num")=product
rs("now")=now
rs("tips")=tips
rs("jieshou_num")=0
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
rs("classid")=6
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
rs("username")=session("useridname")
rs("now")=now
rs("bei")=info
rs("pid")=pid
rs("num")=1
rs("num2")=product
rs("title")=title
rs("tips")=tips
rs.update
rs.close
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
rs("fabudian")=rs("fabudian")-fabu
rs.update
rs.close
call check_jieducm_name(session("useridname"))	
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs.addnew
rs("username")=session("useridname")
rs("num")=pid
rs("class")="�Ա��ղ���"
rs("content")=session("useridname")&"��������"&pid&"�ɹ�,�����������"&fabu&"��"
rs("price")=0
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