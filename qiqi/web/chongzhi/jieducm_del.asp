<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
rs.open "select price,bei from "&jieducm&"_zhongxin where pid='"&id&"' and classid='5' and  username='"&session("useridname")&"' and start='0' ",conn,1,1
if ( rs.eof ) or (rs.bof) then
  Response.Write("<script>alert('����״̬�ѷ����仯!');window.location.href='MyMission.asp';</script>")
  response.end()
else
price=rs("price")
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3


 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select cunkuan From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
		rs("cunkuan")=rs("cunkuan")+price
    	rs.update
    	rs.close	

     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="��ֵ����������"
		rs("content")=session("useridname")&"ȡ������"&id&"�ɹ�,����"&price&"Ԫ���"
		rs("price")="+"&price
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close	
			call check_jieducm_name(session("useridname"))  
set rs=nothing
conn.close
set conn=nothing
response.Redirect("MyMission.asp")

%>