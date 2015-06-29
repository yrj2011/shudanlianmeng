<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
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
id=request("id")
action=request("action")
set Rs=server.createobject("adodb.recordset")
sql="select fajifei from "&jieducm&"_stup "
Rs.open sql,conn,1,1
if not(Rs.eof)  then
stupfajifei=rs("fajifei")
end if
if  action="qing" then
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and start='1'",conn,3,3
       if not (rs.eof) then
           rs("start")="0" 
           rs.update
		   rs.close
		else
	      Response.Write("<script>alert('操作有误!');history.back();</script>")
          response.End()
		end if
		
      sqlr="delete  from "&jieducm&"_jieshou where pid in ('"&id&"')"
      conn.execute sqlr
		   
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="清理买家"
		rs("content")=session("AdminName")&"进行清理买家"&id&"任务"
		rs("price")=renum
		rs("jiegou")="清理买家成功"
    	rs.update
    	rs.close
		
       response.Redirect("admin_mymission4.asp")
      
elseif action="del" then

      Sqli = "select * from "&jieducm&"_zhongxin where pid='"&id&"'and start='0' order by id"
	  Set rs = Server.CreateObject("Adodb.RecordSet")
	  rs.Open Sqli,conn,1,1
	  if not rs.eof or not rs.bof then				
       price=rs("price")
       username=rs("username")
      end if
	  
sqlr="update "&jieducm&"_register set cunkuan=cunkuan+"&price&" where username='"&username&"'"
conn.execute sqlr

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "delete  from "&jieducm&"_zhongxin where pid = ('"&id&"')",conn,3,3
	   	   
	   
	  num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="清理删除"
		rs("content")=session("AdminName")&"进行清理删除"&id&"任务"
		rs("price")=renum
		rs("jiegou")="清理删除成功"
    	rs.update
    	rs.close	   
       response.Redirect("admin_mymission4.asp")
	   
elseif action="back" then

      Set rs=server.createobject("ADODB.RECORDSET")
      rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
       if not (rs.eof) then
           rs("start")=int(rs("start"))-1
           rs.update
		   rs.close
		end if
		
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		   rs("start")=int(rs("start"))-1
           rs.update
		   rs.close
		   end if
		   
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="返回上级"
		rs("content")=session("AdminName")&"进行返回上级"&id&"任务"
		rs("price")=renum
		rs("jiegou")="返回上级成功"
    	rs.update
    	rs.close 		
       response.Redirect("admin_mymission4.asp")

elseif action="ok" then
      Set rs=server.createobject("ADODB.RECORDSET")
      rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
       if not (rs.eof) then
           rs("start")=3
           rs.update
		   rs.close
		end if
		   
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		   rs("start")=3
           rs.update
		   rs.close
		   end if
		   
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="确认发货"
		rs("content")=session("AdminName")&"进行确认发货"&id&"任务"
		rs("price")=renum
		rs("jiegou")="确认发货成功"
    	rs.update
    	rs.close 
        response.Redirect("admin_mymission4.asp")

elseif action="hao" then

       Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
       if not (rs.eof) then
           rs("start")=4
           rs.update
		   rs.close
		end if
		   
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		   rs("start")=4
           rs.update
		   rs.close
		   end if
		   
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="买方好评"
		rs("content")=session("AdminName")&"进行买方好评"&id&"任务"
		rs("price")=renum
		rs("jiegou")="买方好评成功"
    	rs.update
    	rs.close 		   
     response.Redirect("admin_mymission4.asp")

elseif action="haook" then

       Set rs=server.createobject("ADODB.RECORDSET")
      rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
       if not (rs.eof) then
	       fusername=rs("username")
           rs("start")=5
           rs.update
		   rs.close
		end if
				   
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		      jusername=rs("username")
              price=rs("price")
		      rs("start")=5
              rs.update
			  rs.close
			end if
		   
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&jusername&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+price
rs("cunkuan")=rs("cunkuan")+price
rs.update
rs.close
		   	
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="卖方完成"
		rs("content")=session("AdminName")&"进行卖方完成"&id&"任务"
		rs("price")=renum
		rs("jiegou")="卖方完成成功"
    	rs.update
    	rs.close 	   
        response.Redirect("admin_mymission4.asp")
end if

%>