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
rs.open "select num from "&jieducm&"_zhongxin where pid='"&id&"' and classid='6' and  username='"&session("useridname")&"'  and jieshou_num=0",conn,1,1
if ( rs.eof ) or (rs.bof) then
  Response.Write("<script>alert('����״̬�ѷ����仯!');window.location.href='MyMission.asp';</script>")
  response.end()
else
num=rs("num")
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3

fabu=num*stupcang

 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
		rs("fabudian")=rs("fabudian")+fabu
    	rs.update
    	rs.close	

     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�Ա��ղ���"
		rs("content")=session("useridname")&"ȡ������"&id&"�ɹ�,���������㣺"&fabu&"��"
		rs("price")=0
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close	
		
set rs=nothing
conn.close
set conn=nothing
response.Redirect("MyMission.asp")

%>