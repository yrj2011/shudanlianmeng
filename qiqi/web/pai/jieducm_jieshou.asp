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
    Response.Write("<script>alert('�������ĺŲ���Ϊ��');history.back();</script>")
    response.End()		 
end if		 

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select Shopkeeper,prourl,username,Limit from "&jieducm&"_zhongxin where pid='"&id&"' and start='0' ",conn,3,3
if rs.eof then
    Response.Write("<script>alert('�Բ���,�������ѱ������˽���!');history.back();</script>")
	response.End()
else
 username2=rs("username")
 prourlj=rs("prourl")
 Shopkeeper=Replace(Trim(rs("Shopkeeper")),"'","''")
 Limit=rs("Limit")
end if
if username2=session("useridname") then
Response.Write("<script>alert('���ܽ����Լ�������!');history.back();</script>")
response.End()
end if
call jname(username2)

if Limit=2 and jifei<100 then
  Response.Write("<script>alert('����������100�������ϵĻ�Ա�ſ��Խ���!');history.back();</script>")
  response.End()
end if
if Limit=3 and jifei<300 then
  Response.Write("<script>alert('����������300�������ϵĻ�Ա�ſ��Խ���!');history.back();</script>")
  response.End()
end if
if Limit=4 and vip="0" then
  Response.Write("<script>alert('����������ֻ����VIP��Ա�ſ��Խ���!');history.back();</script>")
  response.End()
end if

sql="SELECT  count(*) as totalj FROM "&jieducm&"_jieshou where classid='2' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalj")
end if 

if int(totalzx)>=int(stupjieweiok) and vip<>"1" then
	Response.Write("<script>alert('������ͨ��Աֻ�ܽ���"&stupjieweiok&"��δ��ɵ�����\n��������ѽӵ�����!�����ټ���������');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupjieweiok+5) and vip="1" then
	Response.Write("<script>alert('VIP��Աֻ�ܽ���"&stupjieweiok+5&"��δ��ɵ�����\n��������ѽӵ�����!�����ټ���������');history.back();</script>")
    response.End()
end if 

sql="SELECT  level1,vip FROM "&jieducm&"_register where username='"&username2&"' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
vipf=rs("vip")
jieducm_level1=rs("level1")
end if 

if jieducm_level1="0" then
  Response.Write("<script>alert('�˻�Ա�˺��ѱ�����!�����ٽӴ˻�Ա����');history.back();</script>")
  response.End()
end if


sql="SELECT  count(*) as totalnum FROM  "&jieducm&"_jieshou where username='"&session("useridname")&"' and  num='"&num&"' and proUrl='"&ProUrlj&"' and  (datediff(month,[now],getdate())=0) and classid='2'"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jh=rs("totalnum")
end if
if jh>=1 then
 Response.Write("<script>alert('��ͬ��Ʒ��ַһ����ֻ�ܽ���һ��!');history.back();</script>")				 
 response.End()
end if


sql="SELECT  count(*) as monthnum FROM  "&jieducm&"_jieshou where Shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"' and  num='"&num&"' and classid='2' and  (datediff(month,[now],getdate())=0) "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_monthnum=rs("monthnum")
end if
if jieducm_monthnum>=stupgeshu then
 Response.Write("<script>alert('���С���Ѿ��������ű����ѽ��׹�"&stupgeshu&"����,\n�ٽ�����ȥ�Է��ᷢŭ��!\n��������С��,���߽ӱ��˵������!');history.back();</script>")				 
 response.End()
end if

sql="SELECT  count(*) as monthip FROM  "&jieducm&"_jieshou where Shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"' and  ip='"&ip&"' and classid='2' and  (datediff(month,[now],getdate())=0) "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_monthip=rs("monthip")
end if

if jieducm_monthip>=3 then
 Response.Write("<script>alert('һ������ÿ��IPֻ�ܽ���3��ͬһ�����̵���Ʒ!');history.back();</script>")				 
 response.End()
end if

if vipf="1" then

sql="SELECT  count(*) as datanum FROM  "&jieducm&"_jieshou where shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"'  and num='"&num&"' and classid='2' and  ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_datamnum=rs("datanum")
end if
if jieducm_datamnum>=5 then
 Response.Write("<script>alert('ÿ��ÿ���Ա�������ֻ�ܽ���ͬһ������5����Ʒ����!');history.back();</script>")				 
 response.End()
end if


sql="SELECT  count(*) as jieducm_aobaonum FROM  "&jieducm&"_jieshou where  num='"&num&"' and classid='2' and  ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_tbdatamnum=rs("jieducm_aobaonum")
end if
if jieducm_tbdatamnum>=20 then
 Response.Write("<script>alert('ÿ��ÿ���Ա�������һ��ֻ�ܽ�20����Ʒ����!');history.back();</script>")				 
 response.End()
end if

sql="SELECT  count(*) as dataip FROM  "&jieducm&"_jieshou where Shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"' and  ip='"&ip&"' and classid='2' and  (datediff(day,[now],getdate())=0) "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_dataip=rs("dataip")
end if

if jieducm_dataip>=1 then
 Response.Write("<script>alert('ÿ��ÿ��IPֻ�ܽ���һ��ͬһ�����̵���Ʒ!');history.back();</script>")				 
 response.End()
end if

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
		rs("classid")=2
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
		rs("class")="����������"
		rs("content")=session("useridname")&"����"&user&"������"&id&"�ɹ�"
		rs("price")=0
		rs("jiegou")="δ֧��"
    	rs.update
    	rs.close	

set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('������ȡ���������,��������Ʒ,\n��ϵ������,�������֧��,\n�����֧��,��������������ʱ�ѿ����˱�֤��,\n������ɺ�,�㽫��ô��������!');window.location.href='remission.asp';</script>")
response.End()
%>