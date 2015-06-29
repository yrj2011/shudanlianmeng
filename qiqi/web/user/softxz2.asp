<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统**************************************
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
action=HtmlEncode(trim(request.form("action")))
id=HtmlEncode(trim(request.form("id")))
pushnum=HtmlEncode(trim(request.form("pushnum")))

if action="chongzhi" then
if vip=1 then
pushnum1=int(pushnum)*stupkuanvip
else
pushnum1=int(pushnum)*stupkuan
end if
cunkuan2=cunkuan*100
pushnum2=pushnum1*100

	 if int(cunkuan2)<int(pushnum2) then
			 Response.Write("<script>alert('您的存款金额不足!');window.location.href='mai.asp';</script>")
			 response.End()
	elseif int(pushnum)<1 then
			 Response.Write("<script>alert('输入有误!');history.back();</script>")
	        response.End()
     end if	  
	 
elseif action="carchong" then
     
	 Sql= "select * from "&jieducm&"_car where id="&id&""
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if not rs.eof then
	    car_name=rs("car_name")
		car_price=rs("car_price")
	    car_fabudian=rs("car_fabudian")
		rs.close
	else
		 Response.Write("<script>alert('无此信息!');history.back();</script>")
        response.End()
	 end if 
	 if int(cunkuan)<int(5) then
		Response.Write("<script>alert('您的存款金额不足!');history.back();</script>")
        response.End()
	 end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select jifei,cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+ 0
rs("cunkuan")=rs("cunkuan")-5
rs("fabudian")=rs("fabudian")+0
rs.update
rs.close
	
	
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="购买软件"
		rs("username")=session("useridname")
    	rs("fabudian")=car_fabudian
		rs("cunkuan")=car_price
		rs("jifei")=car_fabudian
    	rs.update
    	rs.close 
		
	   num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("price")="-"&car_price
		rs("jiegou")="扣款成功！进入下载页面？"
    	rs.update
    	rs.close 
call check_jieducm_name(session("useridname"))	 
Response.Write("<script>alert('扣款成功！进入下载页面？');window.location.href='http://yunpan.cn/Q9vs3ZdsUybWx';</script>") 
end if	
set rs=nothing
call closeconn()	
%>
			  
