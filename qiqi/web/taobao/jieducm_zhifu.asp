<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
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
id=request.QueryString("id")
Sql= "select now,start,username from "&jieducm&"_jieshou where pid='"&id&"' and start ='1'  and username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,3,3
if not rs.eof then			
 jnow=rs("now")
 sss= DateDiff( "n",jnow, Now())
  if 20-sss<1 then
    Response.Write("<script>alert('�Բ���,�������ѳ�ʱ!');window.location.href='remission.asp';</script>")
	response.End()
  end if
 rs("start")="2"
 rs.update
 rs.close
else
 Response.Write("<script>alert('�Բ���,��������!');window.location.href='remission.asp';</script>")
 response.End()
end if


Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
if not (rs.eof) then
  user=rs("username")
  rs("start")="2"
  rs.update
  rs.close
end if


	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�Ա�������"
		rs("content")=session("useridname")&"����"&user&"������"&id&"ȷ����֧��"
		rs("price")=0
		rs("jiegou")="��֧��"
    	rs.update
    	rs.close	
set rs=nothing
conn.close
set conn=nothing
response.Redirect("remission.asp")

%>