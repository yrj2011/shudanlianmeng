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
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
dim i
i=0
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
rs.open "select prourl,username from "&jieducm&"_zhongxin where pid='"&id&"' and start='0' ",conn,1,1
if rs.eof then
    Response.Write("<script>alert('�Բ���,�������ѱ������˽���!');history.back();</script>")
	response.End()
else
 username2=rs("username")
 prourlj=rs("prourl")
 Shopkeeper=rs("Shopkeeper")
end if

sql="SELECT  level FROM "&jieducm&"_register where username='"&username2&"' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_level1=rs("level1")
end if 

if jieducm_level1="0" then
  Response.Write("<script>alert('�˻�Ա�˺��ѱ�����!�����ٽӴ˻�Ա����');history.back();</script>")
  response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and start='0'",conn,3,3
if not (rs.eof) then
  classid=rs("classid")
  Price=rs("Price")
  Shopkeeper=rs("Shopkeeper")
  ProUrl=rs("ProUrl")
  bei=rs("bei")
  baohu2=rs("baohu2")
  user=rs("username")
  isprice=rs("isprice")
  rs("start")=1
  rs.update   
  rs.close
else
  Response.Write("<script>alert('�Բ���,�������ѱ������˽���!');history.back();</script>")				 
  response.End()
end if


     Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_jieshou where pid='"&id&"'" ,Conn,3,3  
	  if rs.eof then 
	    rs.addnew
		rs("Price")=Price
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("bei")=bei
		rs("baohu2")=baohu2
		rs("username")=session("useridname")
		rs("username2")=username2
		rs("pid")=id
		rs("start")="1"
		rs("num")=num
		rs("classid")=classid
		rs("ip")=ip
		rs("isprice")=isprice
    	rs.update
    	rs.close
	  else
	    Response.Write("<script>alert('�Բ���,�������ѱ������˽���!');history.back();</script>")				 
        response.End()
	  end if
	  
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="��ֵ����������"
		rs("content")=session("useridname")&"����"&user&"������"&id&"�ɹ�"
		rs("price")=0
		rs("jiegou")="δ֧��"
    	rs.update
    	rs.close	
    Response.Write("<script>alert('������ȡ���������,��������Ʒ,\n��ϵ������,�������֧��,\n�����֧��,��������������ʱ�ѿ����˱�֤��,\n������ɺ�,�㽫��Ĵ��������!');window.location.href='remission.asp';</script>")
response.End()

set rs=nothing
conn.close
set conn=nothing
%>