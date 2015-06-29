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
'Copyright (C) 2008－2009 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
id=request("id")
action=request("action")
set Rs=server.createobject("adodb.recordset")
sql="select fajifei,kou,vipjifei,cang from "&jieducm&"_stup "
Rs.open sql,conn,1,1
if not(Rs.eof)  then
stupfajifei=rs("fajifei")
stupkou=rs("kou")
stupvipjifei=rs("vipjifei")
stupcang=rs("cang")
end if
if  action="qing" then
  
  	     
elseif action="del" then

      Sqli = "select * from "&jieducm&"_zhongxin where id="&id&" and jieshou_num=0"
	  Set rs = Server.CreateObject("Adodb.RecordSet")
	  rs.Open Sqli,conn,1,1
	  if not rs.eof or not rs.bof then				
       username=rs("username")
       num=rs("num")
	  else
	  	Response.Write("<script>alert('此任务已在进行中，不能删除');history.back();</script>")
        response.End()	
      end if

fabu=num*stupcang
sqlr="update "&jieducm&"_register set fabudian=fabudian+"&fabu&"  where username='"&username&"'"
conn.execute sqlr

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "delete  from "&jieducm&"_zhongxin where id in ("&id&")",conn,3,3
	   	   
	   
	  num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")="清理删除"
		rs("content")=session("AdminName")&"进行清理删除"&id&"任务,返还发布点："&fabu&"个"
		rs("price")=0
		rs("jiegou")="清理删除成功"
    	rs.update
    	rs.close
		
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")="清理删除"
		rs("content")=session("AdminName")&"进行清理删除"&id&"任务"
		rs("price")=0
		rs("jiegou")="清理删除成功"
    	rs.update
    	rs.close
			   
set rs=nothing
conn.close
set conn=nothing		
response.Redirect("admin_mymission5.asp")  	   
elseif action="qingli" then

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_jieshou where id='"&id&"' and start='1'",conn,1,1
       if not (rs.eof) then
	       usrnamef=rs("username")
		   pid=rs("pid")
		else
	      Response.Write("<script>alert('操作有误!此任务已不存在或已完成');history.back();</script>")
          response.End()
		end if
		
      sqlr="delete  from "&jieducm&"_jieshou where id = ('"&id&"')"
      conn.execute sqlr
	
	   Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_zhongxin where pid='"&pid&"'",conn,3,3
       if not (rs.eof) then
           rs("jieshou_num")=rs("jieshou_num")-1
           rs.update
		   rs.close
		else
	      Response.Write("<script>alert('操作有误!');history.back();</script>")
          response.End()
		end if
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=usrnamef
    	rs("num")=num
		rs("class")="清理买家"
		rs("content")=session("AdminName")&"进行清理买家"&pid&"任务"
		rs("price")=0
		rs("jiegou")="清理买家成功"
    	rs.update
    	rs.close
		
	   Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")="清理买家"
		rs("content")=session("AdminName")&"进行清理买家"&pid&"任务"
		rs("price")=0
		rs("jiegou")="清理买家成功"
    	rs.update
    	rs.close
		
set rs=nothing
conn.close
set conn=nothing		
response.Redirect("admin_MyMissionjie5.asp")
elseif action="ok" then

		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where id='"&id&"' and start='1'",conn,3,3
           if not (rs.eof) then
		      jusername=rs("username")
			  pid=rs("pid")
		      rs("start")=2
              rs.update
			  rs.close
			else
			  Response.Write("<script>alert('操作有误!此任务已不存在或已完成');history.back();</script>")
               response.End()
			end if
			
       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_zhongxin where pid='"&pid&"'",conn,3,3
       if not (rs.eof) then
	       fusername=rs("username")
		end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&jusername&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+stupfajifei
rs("fabudian")=rs("fabudian")+stupcang
rs.update
rs.close
	   	
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=jusername
    	rs("num")=num
		rs("class")="后台完成收藏"
		rs("content")=session("AdminName")&"进行后台处理卖方完成"&id&"任务，你已得到"&stupcang&"个发布点"&stupfajifei&"积分"
		rs("price")=0
		rs("jiegou")="后台完成收藏"
    	rs.update
    	rs.close 
			   
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=fusername
    	rs("num")=num
		rs("class")="后台完成收藏"
		rs("content")=session("AdminName")&"进行后台处理卖方完成"&id&"任务，接手方已得到"&stupcang&"个发布点"&stupfajifei&"积分"
		rs("price")=0
		rs("jiegou")="后台完成收藏"
    	rs.update
    	rs.close 
		
set rs=nothing
conn.close
set conn=nothing		
response.Redirect("admin_MyMissionjie5.asp") 
end if

%>