<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
id=request.QueryString("id")
num=request("xiaohaoname")
ip=Request.ServerVariables("REMOTE_ADDR")
if isjie=1 then
 Response.Write("<script>alert('���ѱ�����Ա��ֹ�˽�������!');history.back();</script>")
 response.End()
elseif num="" then 
    Response.Write("<script>alert('�����Ա��Ų���Ϊ��');history.back();</script>")
    response.End()		 
end if		 

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select Shopkeeper,prourl,username,num,jieshou_num from "&jieducm&"_zhongxin where pid='"&id&"'",conn,1,1
if rs.eof then
    Response.Write("<script>alert('�Բ���,�������Ѳ�����!');history.back();</script>")
	response.End()
else
 username2=rs("username")
 prourlj=rs("prourl")
 jieducm_num=rs("num")
 jieshou_num=rs("jieshou_num")
end if

if jieducm_num=jieshou_num then
    Response.Write("<script>alert('�Բ���,�������ѽ���!');history.back();</script>")
	response.End()
elseif username2=session("useridname") then
    Response.Write("<script>alert('���ܽ��Լ�������!');history.back();</script>")
	response.End()
end if 

call jname(username2)

sql="SELECT  count(*) as totalnum FROM  "&jieducm&"_jieshou where username='"&session("useridname")&"' and num='"&num&"' and proUrl='"&ProUrlj&"' and classid='6'"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jh=rs("totalnum")
end if
if jh>=1 then
 Response.Write("<script>alert('��ͬ��Ʒ��ַֻ�ܽ���һ��!');history.back();</script>")				 
 response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' ",conn,3,3
if not (rs.eof) then
  Shopkeeper=rs("Shopkeeper")
  ProUrl=rs("ProUrl")
  bei=rs("bei")
  rs("jieshou_num")=rs("jieshou_num")+1
  rs.update   
  rs.close
end if


     Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_jieshou " ,Conn,3,3  
	    rs.addnew
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("bei")=bei		
		rs("username")=session("useridname")
		rs("username2")=username2
		rs("pid")=id
		rs("start")="1"
		rs("num")=num
		rs("classid")=6
		rs("ip")=ip
    	rs.update
    	rs.close

	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�Ա��ղ���"
		rs("content")=session("useridname")&"����"&user&"������"&id&"�ɹ�"
		rs("price")=0
		rs("jiegou")="�ȴ��ղ�"
    	rs.update
    	rs.close	

set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('������ȡ���������,���ղ���Ʒ����̵�ַ,\n������ɺ�,�㽫��÷�����!');window.location.href='remission.asp';</script>")
response.End()
%>