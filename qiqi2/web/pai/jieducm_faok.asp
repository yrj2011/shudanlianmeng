<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
On Error Resume Next 
Server.ScriptTimeOut=9999999 
action=request.form("action")
if action="ok" then
Price=Replace(Trim(Request.Form("Price1")),"'","''")
ProUrl=Replace(Trim(Request.Form("ProUrl1")),"'","''")
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
bei=Replace(Trim(Request.Form("bei")),"'","''")
baohu2=request.Form("baohu2")
isprice=request.form("isprice")
depot=request.form("depot")
play=request.form("play")
tips=Replace(Trim(Request.Form("tips")),"'","''")
Limit=request.Form("Limit")
title=Replace(Trim(Request.Form("title")),"'","''")

if cunkuan<1 then
 Response.Write("<script>alert('������1Ԫʱ�����ٷ�����');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif fabudian<1 then
 Response.Write("<script>alert('���������1��ʱ�����ٷ�����');window.location.href='../user/md5_pay.asp';</script>")
 response.End() 
elseif price="" then
	Response.Write("<script>alert('��Ʒ����۲���Ϊ��!');history.back();</script>")
    response.End()
elseif int(price)<1  or int(price)>1000 then
	Response.Write("<script>alert('Ŀǰ��ƽֻ̨֧�ַ���1-1000Ԫ������ ������������!');history.back();</script>")
    response.End()
elseif isfa="1" then
	Response.Write("<script>alert('���ѱ�����Ա��ֹ�˷�������!');history.back();</script>")
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
elseif not isNumeric(price) then
          Response.Write("<script>alert('�����������������룡');history.back();</script>")
          response.End()
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='2' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(stupfaweiok) and vip<>"1" then
	Response.Write("<script>alert('������ͨ��Աֻ�ܷ���"&stupfaweiok&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupfaweiok+5) and vip="1" then
	Response.Write("<script>alert('��VIP��Աֻ�ܷ���"&stupfaweiok+5&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
end if 



if price>=1 and price<=40 then
   fabu=1
elseif price>40 and price<=100 then
   fabu=2
elseif price>100 and price<=200 then
   fabu=3
elseif price>200 and price<=500 then
  fabu=4
elseif price>500 and price<=1000 then
  fabu=5
end if 
if bei="һ����ջ�����" then
fabu=fabu*2
elseif bei="������ջ�����"  then
fabu=fabu*2+1
elseif bei="������ջ�����"  then
fabu=fabu*2+2
end if

price2=price*100
cunkuan2=cunkuan*100
fabudian2=fabudian*100
fabu2=fabu*100

if(int(price2)>int(cunkuan2)) then
    Response.Write("<script>alert('�㷢���������Ǯ���ܴ�����Ĵ��!');history.back();</script>")
    response.End()
elseif(int(fabudian2)<int(fabu2)) then
   Response.Write("<script>alert('��ķ�������������,�뾡���ֵ��ȡ������!');history.back();</script>")
   response.End()			
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='2' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(stupfaweiok) and vip<>"1" then
	Response.Write("<script>alert('������ͨ��Աֻ�ܷ���"&stupfaweiok&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupfaweiok+5) and vip="1" then
	Response.Write("<script>alert('��VIP��Աֻ�ܷ���"&stupfaweiok+5&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
end if 


randomize
if bei="�����ջ�����"  then
company=""
numbers=""
else
Cid=int(7*rnd)+1
if cid=1 then
  company="��ͨE����"
  if cid>4 then
     top="268"
  else
     top="888"
  end if
  rnum=int(900000000*rnd)+100000000
  numbers=top&rnum
elseif cid=2 then
  company="Բͨ�ٵ�"
  if cid=1 then
    top="12"
  elseif cid=2 then
    top="21"
  elseif cid=3 then
    top="22"
  else
    top="82"
  end if
  rnum=int(90000000*rnd)+10000000
  numbers=top&rnum
elseif cid=3 then
  company="��ͨ���"
 if cid=1 then
  top="6180"
 elseif cid=2 then
  top="6800"
 elseif cid=3 then
  top="2008"
 elseif cid=4 then
  top="5711"
 else
  top="0100"
 end if
  rnum=int(90000000*rnd)+10000000
  numbers=top&rnum
elseif cid=4 then
  company="լ����"
  if cid=1 then
   top="1"
  elseif cid=2 then
   top="2"
  elseif cid=3 then
    top="6"
  elseif cid=4 then
    top="7"
  else
    top="8"
  end if
  rnum=int(900000000*rnd)+100000000
  numbers=top&rnum
elseif cid=5 then
  company="������"
  top="120"
  rnum=int(900000000*rnd)+100000000
  numbers=top&rnum
elseif cid=6 then
  company="�°�����"
  top="1"
  rnum=int(9000000*rnd)+1000000
  numbers=top&rnum
elseif cid=7 then
  company="����ƽ��"
  if cid>4 then
  top="KA"
  else
  top="PA"
  end if
  rnum=int(90000000000*rnd)+10000000000
  numbers=top&rnum
elseif cid=8 then
  company="EMS���"
  c=Chr(Int(Rnd * 26 + 65))
  rnum=int(900000000*rnd)+100000000
  numbers="E"&c&rnum&"CN"
end if
end if
		
ranNum=int(90000*rnd)+10000	
pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin  where pid='"&pid&"'" ,Conn,3,3
if  rs.eof then  
rs.addnew
rs("classid")=2
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
rs("isprice")=isprice
rs("play")=play
rs("tips")=tips
rs("company")=company
rs("numbers")=numbers
rs("Limit")=Limit
rs("jieducm_fb")=fabu
rs.update
rs.close
else
Response.Write("<script>alert('�����ˣ��뷵�����·���!');history.back();</script>")
response.End()
end if

if depot=1 then
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_depot " ,Conn,3,3  
rs.addnew
rs("classid")=2
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
rs("now")=now
rs("isprice")=isprice
rs("pid")=pid
rs("num")=1
rs("play")=play
rs("tips")=tips
rs("company")=company
rs("numbers")=numbers
rs("Limit")=Limit
rs("title")=title
rs("jieducm_fb")=fabu
rs.update
rs.close
end if
      
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
rs("cunkuan")=rs("cunkuan")-price
rs("fabudian")=rs("fabudian")-fabu
rs.update
rs.close
		 
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=pid
		rs("class")="����������"
		rs("content")=session("useridname")&"��������"&pid&"�ɹ�,������"&price&"Ԫ,�����������"&fabu&"��"
		rs("price")="-"&price
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close	
call check_jieducm_name(session("useridname")) 		
set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('�����ɹ�!�ȴ������ˣ�');window.location.href='MyMission.asp';</script>")	
response.End()
end if
%>