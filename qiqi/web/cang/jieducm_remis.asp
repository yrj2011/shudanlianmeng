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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
id=request.QueryString("id")

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_jieshou where id='"&id&"'  and username='"&session("useridname")&"' and classid='6' and start='1'",conn,1,1
if rs.eof then
    Response.Write("<script>alert('�Բ���,���������޴���Ϣ!');window.location.href='remission.asp';</script>")
	response.End()
else
pid=rs("pid")
end if 

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_jieshou where id='"&id&"' and username='"&session("useridname")&"' and classid='6'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&pid&"' and classid='6'",conn,3,3
if not (rs.eof) then
  rs("jieshou_num")=rs("jieshou_num")-1
  rs.update
  rs.close
else
    Response.Write("<script>alert('�Բ���,���������޴���Ϣ!');window.location.href='remission.asp';</script>")
	response.End()
end if

	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�Ա��ղ���"
		rs("content")=session("useridname")&"����ID:"&pid&"�˳�����"
		rs("price")=0
		rs("jiegou")="�˳�����"
    	rs.update
    	rs.close	
set rs=nothing
conn.close
set conn=nothing  
response.Redirect("remission.asp")

%>