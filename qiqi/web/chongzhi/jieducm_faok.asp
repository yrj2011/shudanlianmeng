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
On Error Resume Next 
Server.ScriptTimeOut=9999999 
if request.Form("fa")="ok" then
Price=Replace(Trim(Request.Form("Price1")),"'","''")
ProUrl=Replace(Trim(Request.Form("ProUrl1")),"'","''")
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
bei=Replace(Trim(Request.Form("bei")),"'","''")
baohu2=request.Form("baohu2")
tips=Replace(Trim(Request.Form("tips")),"'","''")

if cunkuan<100 then
 Response.Write("<script>alert('������100Ԫʱ���ܷ�����������');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif  instr (ProUrl,"shop") then
 Response.Write("<script>alert('����д��ȷ������ַ��');history.back();</script>")
 response.End()
elseif price="" then
	Response.Write("<script>alert('���ֽ���Ϊ��!');history.back();</script>")
    response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('�㻹û��ѡ���ƹ���!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('��Ʒ��ַ����Ϊ��!');history.back();</script>")
    response.End()
elseif bei="" then
	Response.Write("<script>alert('�㻹û��ѡ��ע!');history.back();</script>")
    response.End()
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='5' and username='"&session("useridname")&"' and ( datediff(day,[now],getdate())=0) "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(1) and vip<>"1" then
	Response.Write("<script>alert('��ͨ��Աÿ��ֻ������һ�����֣�');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(3) and vip="1" then
	Response.Write("<script>alert('VIP��Աÿ��ֻ���������������֣���');history.back();</script>")
    response.End()
end if 

url= ProUrl
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" href=""[^""]+""[^>]*>(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk=Matches(0).SubMatches(0)

wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" title=""(.*?)"""
set Matches=objRegExp.Execute(wstr)
jieducm_sk3=Matches(0).SubMatches(0)

if (jieducm_sk=shopkeeper)   or (jieducm_sk3=shopkeeper) then
else
	Response.Write("<script>alert('���ƹ�������д����Ʒ��ַ����!����������!');history.back();</script>")
    response.End()
end if


jieducm_ck=cunkuan*100
jieducm_price=price*100

if int(jieducm_ck)< int(jieducm_price) then
	 Response.Write("<script>alert('�����ֽ��ܴ�����Ĵ��!');history.back();</script>")
	 response.End()   
elseif isfa="1" then
	Response.Write("<script>alert('���ѱ�����Ա��ֹ�˷�������!');history.back();</script>")
	response.End()
end if
          
        randomize
		ranNum=int(90000*rnd)+10000	
        pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_zhongxin  where pid='"&pid&"'" ,Conn,3,3
		if  rs.eof then  
	    rs.addnew
		rs("classid")=5
		rs("Price")=Price
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("bei")=bei	
		if baohu2="" then
		rs("baohu2")=0
		else
		rs("baohu2")=baohu2
		end if
		rs("username")=session("useridname")
		rs("pid")=pid
		rs("now")=now
		rs("tips")=tips
    	rs.update
    	rs.close
		else
		   Response.Write("<script>alert('�����ˣ��뷵�����·���!');history.back();</script>")
	       response.End()
		end if


      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select cunkuan From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
    	rs("cunkuan")=rs("cunkuan")-price
    	rs.update
    	rs.close
		
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=pid
		rs("class")="��ֵ����������"
		rs("content")=session("useridname")&"������������"&pid&"�ɹ�,������"&price&"Ԫ"
		rs("price")="-"&price
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close
	call check_jieducm_name(session("useridname"))  			
set rs=nothing 
conn.close
set conn=nothing
Response.Write("<script>alert('������������ɹ�!�ȴ������ˣ�');window.location.href='MyMission.asp';</script>")	
response.End()
end if
%>