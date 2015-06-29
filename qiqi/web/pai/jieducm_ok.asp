<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and  classid='2' and start='4' and username='"&session("useridname")&"'",conn,3,3
if  (rs.eof) or  (rs.bof) then
  Response.Write("<script>alert('对不起,操作有误!此信息状态已发生改变！');window.history.back();</script>")
  response.end()
else 
 fusername=rs("username")
 price=rs("price")
 jieducm_fb=rs("jieducm_fb")
  rs("start")="5"
  rs.update
  rs.close
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_jieshou where pid='"&id&"'" ,Conn,3,3  
if not (rs.eof) then
jusername=rs("username")
rs("start")="5"
rs.update
rs.close
end if

Sql = "select username,jifei,tjr,vip from "&jieducm&"_register  where username='"&jusername&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
jpromotion=rs("tjr")
jifei=rs("jifei")
vipj=rs("vip")
end if
				
if vipj="1" then
fabu1=stupvipjifei*int(jieducm_fb)
else		
if jifei>stupjifeibig then
fabu1=stupkou*int(jieducm_fb)
else
fabu1=jieducm_fb
end if
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&jusername&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+ stupfajifei
rs("fabudian")=rs("fabudian")+fabu1
rs("cunkuan")=rs("cunkuan")+price
rs.update
rs.close

sqlr="update "&jieducm&"_register set  jifei=jifei+"&stupfajifei&" where username='"&fusername&"'"
conn.execute sqlr

if jpromotion<>"" then
sqlr="update "&jieducm&"_register set  jifei=jifei+"&stupfajifei&" where username='"&jpromotion&"'"
conn.execute sqlr
end if

if tjr<>"" then
sqlr="update "&jieducm&"_register set  jifei=jifei+"&stupfajifei&" where username='"&tjr&"'"
conn.execute sqlr
end if	
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="拍拍区任务"
		rs("content")=session("useridname")&"确认接手人"&jusername&"的好评-任务ID:"&id&"确认已好评对方收到了你的"&price&"元担保金和"&fabu1&"个发布点"
		rs("price")=0
		rs("jiegou")="确认已好评任务完成!"
    	rs.update
    	rs.close
			 
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=jusername
    	rs("num")=num
		rs("class")="拍拍区任务"
		rs("content")=session("useridname")&"确认你的好评-任务ID:"&id&"确认已好评你已收到了对方的"&price&"元担保金和"&fabu1&"个发布点"
		rs("price")="+"&price
		rs("jiegou")="确认已好评任务完成!"
    	rs.update
    	rs.close
		response.Redirect("mymission.asp")

set rs=nothing
conn.close
set conn=nothing
%>