<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
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
id=request("id")
level1=request("level1")
isjie=request("isjie")
isdun=request("isdun")
islogin=request("islogin")
isfa=request("isfa")
isgive=request("isgive")
isgives=request("isgives")
ismessage=request("ismessage")
wangwang=request("wangwang")
username=request("username1")
taobao=request("taobao")
paipai=request("paipai")
youa=request("youa")
each1=request("each")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_register where id="&id&"",conn,3,3
if not (rs.eof) then
if level1<>"" then
  rs("level1")=level1
else
  rs("level1")=0
end if
if isjie<>"" then
  rs("isjie")=isjie
else
  rs("isjie")=0
end if
if isdun<>"" then
  rs("isdun")=isdun
else
  rs("isdun")=0
end if
if islogin<>"" then
  rs("islogin")=islogin
else
 rs("islogin")=0
end if
if isfa<>"" then
  rs("isfa")=isfa
else
rs("isfa")=0
end if
if isgive<>"" then
  rs("isgive")=isgive
 else
   rs("isgive")=0
end if
if isgives<>"" then
  rs("isgives")=isgives
else
  rs("isgives")=0
end if
if ismessage<>"" then
  rs("ismessage")=ismessage
else
 rs("ismessage")=0
end if

if taobao<>"" then
  rs("taobao")=taobao
 else
   rs("taobao")=0
end if
if paipai<>"" then
  rs("paipai")=paipai
else
 rs("paipai")=0
end if
if youa<>"" then
  rs("youa")=youa
else
 rs("youa")=0
end if
if each1<>"" then
  rs("each")=each1
else
 rs("each")=0
end if
rs.update
rs.close
  
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
rs.addnew
rs("username")=session("AdminName")
rs("class")="�޸��������"
rs("content")="�޸���:"&request("username1")&"�����������"
rs("jiegou")="�ɹ�"
rs.update
rs.close

set rs=nothing
conn.close
set conn=nothing	
Response.Write("<script>alert('�����ɹ�');window.location.href='usermanage.asp?action=sear&text="&username&"';</script>")
response.End()
  end if
%>