<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->

<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ**************************************
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

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_tixian where username='"&session("useridname")&"' and id="&id&" and start='0'",conn,3,3
if  (rs.eof) then
  Response.Write("<script>alert('�������ѽ���!�벻Ҫ�ظ�����!ƽ̨���¼�����Ϊ!');window.history.back();</script>")
  response.end()
else
  price=rs("ReNum")
  rs("start")=2
  rs.update
  rs.close
end if

 			  	Sql = "select cunkuan from "&jieducm&"_register where username='"&session("useridname")&"'"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,3,3
				if not(rs.eof) then
				  rs("cunkuan")=rs("cunkuan")+price
				  rs.update
				  rs.close
				end if
								
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="��������"
		rs("content")=session("useridname")&"���г������� "&price&"Ԫ"
		rs("price")="+"&price
		rs("jiegou")="�����ɹ�"
    	rs.update
    	rs.close
				
 Response.Write("<script>alert('���ѳɹ�������������,\���Ľ���ѷ�������Ĵ����!');window.location.href='ReMoneyList.asp';</script>")	
set rs=nothing
call CloseConn()
%>