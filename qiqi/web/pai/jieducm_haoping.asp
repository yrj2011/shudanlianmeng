<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
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
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"' and start='3' and username='"&session("useridname")&"'",conn,3,3
if (rs.eof) or  (rs.bof) then
    Response.Write("<script>alert('�Բ���,��������!');window.location.href='remission.asp';</script>")
	response.End()
else

nowj=rs("now")
bei=rs("bei")
if bei="�����ջ�����"  then
j=0
elseif bei="һ����ջ�����" then
j=24
elseif bei="������ջ�����"  then
j=48
elseif bei="������ջ�����" then
j=72
end if

sss= DateDiff( "h", nowj, Now())
if cint(j)-cint(sss)>=1 then
  Response.Write("<script>alert('�Բ��𣬴������ջ�����ʱ�仹δ��!');window.location.href='remission.asp';</script>")
  response.End()
end if
  rs("start")="4"
  rs("now")=now
  rs.update
  rs.close
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select start,username from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
if not (rs.eof) then
  user=rs("username")
  rs("start")="4"
  rs.update
  rs.close
end if
	num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="����������"
		rs("content")=session("useridname")&"����"&user&"������ID:"&id&"ȷ���Ѻ���"
		rs("price")=0
		rs("jiegou")="ȷ���Ѻ���"
    	rs.update
    	rs.close	
		response.Redirect("remission.asp")     
set rs=nothing
conn.close
set conn=nothing
%>