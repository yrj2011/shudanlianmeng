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
	 elseif czm<>request("code") then
		    Response.Write("<script>alert('操作码不正确!');history.back();</script>")
	        response.End()
	elseif int(pushnum)<1 then
			 Response.Write("<script>alert('输入有误!');history.back();</script>")
	        response.End()
	elseif not isNumeric(pushnum) then
          Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
          response.End()
     end if	  
	 
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select fabudian,cunkuan From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
		rs("fabudian")=rs("fabudian")+pushnum
		rs("cunkuan")=rs("cunkuan")-pushnum1
    	rs.update
    	rs.close  
		 
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="购买发布点"
		rs("username")=session("useridname")
    	rs("fabudian")=pushnum
    	rs.update
    	rs.close  
		  
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="购买发布点"
		rs("content")=session("useridname")&"进行在线购买发布点"&pushnum&"个"
		rs("price")="-"&pushnum1
		rs("jiegou")="购买成功"
    	rs.update
    	rs.close    
		call check_jieducm_name(session("useridname"))         
	    Response.Write("<script>alert('充值成功!');window.location.href='mai.asp';</script>")	
	    response.End()
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
	 if int(cunkuan)<int(car_price) then
		Response.Write("<script>alert('您的存款金额不足!');history.back();</script>")
        response.End()
	 end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select jifei,cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+ car_price
rs("cunkuan")=rs("cunkuan")-car_price
rs("fabudian")=rs("fabudian")+car_fabudian
rs.update
rs.close
	
if tjr<>"" then
 tk=car_price*0.1
 sqlr="update "&jieducm&"_register set cunkuan=cunkuan+"&tk&" where username='"&tjr&"'"
 conn.execute sqlr
num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs.addnew
rs("username")=tjr
rs("num")=num
rs("class")="推广奖励"
rs("content")="你推荐的会员"&session("useridname")&"进行在线购买"&car_name&",系统奖励你"&tk&"元存款"
rs("price")=tk
rs("jiegou")="奖励成功"
rs.update
rs.close
end if
	
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="购买点卡"
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
		rs("class")="购买点卡"
		rs("content")=session("useridname")&"进行在线购买"&car_name&",存款减少了"&car_price&"元,发布点增加了"&car_fabudian&"个"
		rs("price")="-"&car_price
		rs("jiegou")="购买成功"
    	rs.update
    	rs.close 
call check_jieducm_name(session("useridname"))
call send_message(tjr,"推广奖励","你推荐的会员"&session("useridname")&"购买了平台点卡，平台奖励你"&tk&"元存款")		 
Response.Write("<script>alert('购买成功!同时你也增加了积分!');window.location.href='mai.asp';</script>") 
end if	
set rs=nothing
call closeconn()	
%>
			  
