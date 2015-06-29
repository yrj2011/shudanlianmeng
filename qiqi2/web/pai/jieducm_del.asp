<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and  classid='2' and  username='"&session("useridname")&"' and start='0' ",conn,1,1
if ( rs.eof ) or (rs.bof) then
  Response.Write("<script>alert('任务状态已发生变化!');window.location.href='MyMission.asp';</script>")
  response.end()
else
bei=rs("bei")
price=rs("price")
jieducm_fb=rs("jieducm_fb")
end if

sqlr="delete  from "&jieducm&"_zhongxin where pid='"&id&"'"
conn.execute sqlr

 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
		rs("fabudian")=rs("fabudian")+jieducm_fb
		rs("cunkuan")=rs("cunkuan")+price
    	rs.update
    	rs.close	

     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="拍拍区任务"
		rs("content")=session("useridname")&"取消任务"&id&"成功,返还发布点："&jieducm_fb&"个,返还存款："&price&"元"
		rs("price")="+"&price
		rs("jiegou")="成功"
    	rs.update
    	rs.close
call check_jieducm_name(session("useridname")) 
set rs=nothing
conn.close
set conn=nothing	
response.Redirect("MyMission.asp")	

%>