<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select num from "&jieducm&"_jieshou where pid='"&id&"' and start='1' and classid='6' and username='"&session("useridname")&"'",conn,1,1
if  (rs.eof) or  (rs.bof) then
  Response.Write("<script>alert('�Բ���,��������!����Ϣ״̬�ѷ����ı䣡');window.history.back();</script>")
  response.end()
else
  num=rs("num")
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin where pid='"&id&"' and classid='6'" ,Conn,1,1  
if not (rs.eof) then
jusername=rs("username")
bei=rs("bei")
Shopkeeper=rs("Shopkeeper")
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select prourl From "&jieducm&"_xinyu where shopname='"&num&"' and class=4" ,Conn,1,1 
if not (rs.eof) then
jieducm_str=rs("prourl")
end if


jieducm_cangid= mid(jieducm_str,54,15)

if bei="�����ղ�����" then
jieducm_idnum=1
else
jieducm_idnum=2
end if
jieducm_ProUrl= "http://favorite.taobao.com/collect_list.htm?itemtype="&jieducm_idnum&"&nuid="&jieducm_cangid&"&pagesize=20"

url=jieducm_ProUrl
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=true
objRegExp.Pattern="<span class=""J_WangWang"" data-nick=""(.*?)"">"
set Matches=objRegExp.Execute(wstr)
jieducm_cang=Matches(0).SubMatches(0)
if jieducm_cang<>Shopkeeper then
   Response.Write("<script>alert('ϵͳ��ⲻ�������ղ�!');history.back();</script>")
   response.End()
end if
   
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select jifei,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+ stupfajifei
rs("fabudian")=rs("fabudian")+(stupcang*0.8)
rs.update
rs.close

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select start from "&jieducm&"_jieshou where pid='"&id&"' and start='1' and classid='6' and username='"&session("useridname")&"'",conn,3,3
if  not rs.eof then
  rs("start")=2
  rs.update
  rs.close
end if

	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�Ա��ղ���"
		rs("content")=session("useridname")&"ȷ���ղ�����ID:"&id&"��õ�"&stupcang*0.8&"��������,��Ҳ�õ�"&stupfajifei&"������"
		rs("price")=0
		rs("jiegou")="ȷ�����ղ�"
    	rs.update
    	rs.close
			 
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=jusername
    	rs("num")=num
		rs("class")="�Ա��ղ���"
		rs("content")=session("useridname")&"ȷ���ղ����������ID:"&id&""
		rs("price")=0
		rs("jiegou")="ȷ�����ղ�"
    	rs.update
    	rs.close
set rs=nothing
conn.close
set conn=nothing
response.Redirect("ReMission.asp")
%>