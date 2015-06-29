<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
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
zfbnum=request("zfbnum")
action=request("action")
if action="ok" then
Set rs1=server.createobject("ADODB.RECORDSET")
rs1.open "select * from "&jieducm&"_tixian where id="&id&"",conn,3,3
if not (rs1.eof) then
price=rs1("ReNum")
username=rs1("username")

  rs1("start")="1"  
  rs1("zfbnum")=zfbnum
  rs1.update
  rs1.close
  
  	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=username
    	rsr("num")=num
		rsr("class")="申请提现"
		rsr("content")="管理员确认提现 "&price&"元"
		rsr("price")=price
		rsr("jiegou")="提现成功"
    	rsr.update
    	rsr.close
  Response.Write("<script>alert('确定此用户提现金额\n已打到该用户支付宝账号上!');window.location.href='admin_bankroll.asp';</script>")
end if

elseif action="rem" then

       Set Rs = Server.CreateObject("Adodb.RecordSet")
	   rs.open "Select * From "&jieducm&"_tixian where  id="&id&" and start<>'2'"  ,Conn,3,3  
	   if not rs.eof then
        rs("start")=2
		username=rs("username")
		price=rs("ReNum")
    	rs.update
    	rs.close
		
 			  	Sql3 = "select * from "&jieducm&"_register where username='"&username&"'"
				Set Rs3 = Server.CreateObject("Adodb.RecordSet")
				Rs3.Open Sql3,conn,3,3
				if not(rs3.eof) then
				  rs3("cunkuan")=rs3("cunkuan")+price
				  rs3.update
				  rs3.close
				end if
				
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=username
    	rsr("num")=num
		rsr("class")="撤销提现"
		rsr("content")="管理员进行撤销提现 "&price&"元"
		rsr("price")="+"&price
		rsr("jiegou")="撤销成功"
    	rsr.update
    	rsr.close
  Response.Write("<script>alert('您已成功撤销此用户提现,\相应金额已返还到此用户的存款中!');window.location.href='admin_bankroll.asp';</script>")
  end if
elseif action="shou" then
       Set Rs = Server.CreateObject("Adodb.RecordSet")
	   rs.open "Select * From "&jieducm&"_tixian where  id="&id&" and start<>'2'"  ,Conn,3,3  
	   if not rs.eof then
        rs("start")=3
    	rs.update
    	rs.close
		 Response.Write("<script>alert('您已成功锁定此用户撤消提现了!');window.location.href='admin_bankroll.asp';</script>")
       end if
elseif action="del" then

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_tixian where id in ("&id&")",conn,3,3
set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('已删除此次提现!');window.location.href='admin_bankroll.asp';</script>")

end if
%>