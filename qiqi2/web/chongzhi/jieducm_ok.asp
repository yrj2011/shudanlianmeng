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
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and classid='5' and start='4' and username='"&session("useridname")&"'",conn,3,3
if  (rs.eof) or  (rs.bof) then
  Response.Write("<script>alert('�Բ���,��������!����Ϣ״̬�ѷ����ı䣡');window.history.back();</script>")
  response.end()
else
  
  fusername=rs("username")
  price=rs("price")
  rs("start")="5"
  rs.update
  rs.close
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select bei,start,username From "&jieducm&"_jieshou where pid='"&id&"'" ,Conn,3,3  
if not (rs.eof) then
jusername=rs("username")
rs("start")="5"
rs.update
rs.close
end if
   
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan From "&jieducm&"_register where username='"&jusername&"'" ,Conn,3,3
rs("jifei")=rs("jifei")+price 
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
		rs("content")=session("useridname")&"ȷ�Ͻ�����"&jusername&"�ĺ���-����ID:"&id&"ȷ���Ѻ����Է��յ������"&price&"Ԫ������"
		rs("price")=0
		rs("jiegou")="ȷ���Ѻ����������!"
    	rs.update
    	rs.close
			 
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=jusername
    	rs("num")=num
		rs("class")="��ֵ����������"
		rs("content")=session("useridname")&"ȷ����ĺ���-����ID:"&id&"ȷ���Ѻ��������յ��˶Է���"&price&"Ԫ������,����"&price&"��"
		rs("price")="+"&price
		rs("jiegou")="ȷ���Ѻ����������!"
    	rs.update
    	rs.close
set rs=nothing
conn.close
set conn=nothing	
response.Redirect("mymission.asp")

%>