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
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and classid='1' and  username='"&session("useridname")&"' and start='0' ",conn,1,1
if ( rs.eof ) or (rs.bof) then
  Response.Write("<script>alert('����״̬�ѷ����仯!');window.location.href='MyMission.asp';</script>")
  response.end()
else
bei=rs("bei")
price=rs("price")
addfabu=rs("addfabu")
jieducm_fb=rs("jieducm_fb")
end if

sqlr="delete  from "&jieducm&"_zhongxin where pid='"&id&"'"
conn.execute sqlr

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
rs("cunkuan")=rs("cunkuan")+price
rs("fabudian")=rs("fabudian")+(jieducm_fb+addfabu)
rs.update
rs.close	

     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�Ա�������"
		rs("content")=session("useridname")&"ȡ������"&id&"�ɹ�,���������㣺"&jieducm_fb+addfabu&"��,������"&price&"Ԫ"
		rs("price")="+"&price
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close	
		
set rs=nothing
conn.close
set conn=nothing
response.Redirect("MyMission.asp")

%>