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
rs.open "select now from "&jieducm&"_jieshou where pid='"&id&"' and start='1' ",conn,3,3
if  not (rs.eof) then
  rs("now")=now
  rs.update
  rs.close
  response.Redirect("MyMission.asp")
  response.End()
else
  Response.Write("<script>alert('�������󣬴���Ϣ״̬�ѷ��ֱ仯��!');window.location.href='MyMission.asp';</script>")
  response.End()
end if
set rs=nothing
call closeconn()
%>