<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------

id=request("id")
action=request("action")
if id<>"" and action="del" then
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_depot where id="&id&"  and  username='"&session("useridname")&"'",conn,3,3
response.Redirect("MySave.asp")	
elseif id<>"" and  action="ok" then
if cunkuan<1 then
 Response.Write("<script>alert('存款低于1元时不能再发任务！');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif fabudian<1 then
 Response.Write("<script>alert('发布点低于1个时不能再发任务！');window.location.href='../user/md5_pay.asp';</script>")
 response.End() 
end if	

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='1' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(stupfaweiok) and vip<>"1" then
	Response.Write("<script>alert('你是普通会员只能发布"&stupfaweiok&"个未完成的任务！\n请先完成已发的任务!才能再继续发任务！');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupfaweiok+5) and vip="1" then
	Response.Write("<script>alert('你VIP会员只能发布"&stupfaweiok+5&"个未完成的任务！\n请先完成已发的任务!才能再继续发任务！');history.back();</script>")
    response.End()
end if 
			
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_depot where id="&id&" and  username='"&session("useridname")&"'",Conn,1,1
   if not rs.eof or not rs.bof then
    Price=rs("Price")
    ProUrl=rs("ProUrl")
    Shopkeeper=rs("Shopkeeper")
    bei=rs("bei")
    baohu2=rs("baohu2")
    isprice=rs("isprice")
    depot=rs("depot")
	isprice=rs("isprice")
	play=rs("play")
	tips=rs("tips")
	addfabu=rs("addfabu")
	Limit=rs("Limit")
	ping=rs("ping")
    pingtxt=rs("pingtxt")
	jieducm_fb=rs("jieducm_fb")
	IfComeSet=rs("IfComeSet")
    ComeKey=rs("ComeKey")
    ComeNote=rs("ComeNote")
	else
	    Response.Write("<script>alert('操作有误!');history.back();</script>")
	    response.End()
	end if
	

fabu=jieducm_fb+addfabu
jieducm_ck=cunkuan*100
jieducm_price=price*100
fabudian2=fabudian*100
fabu2=fabu*100

if isfa="1" then
Response.Write("<script>alert('你已被管理员禁止了发任务功能!');history.back();</script>")
response.End()
elseif(int(jieducm_price)>int(jieducm_ck)) then
 Response.Write("<script>alert('你发部的任务价钱不能大于你的存款!');history.back();</script>")
 response.End()
elseif(int(fabudian2)<int(fabu2)) then
 Response.Write("<script>alert('你的发布点数量不足,请尽快充值或换取!');history.back();</script>")
 response.End()
end if
	
		randomize
		ranNum=int(90000*rnd)+10000	
        pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_zhongxin" ,Conn,3,3  
	    rs.addnew
		rs("classid")=1
		rs("Price")=Price
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("bei")=bei
		rs("baohu2")=baohu2
		rs("username")=session("useridname")
		rs("pid")=pid
		rs("now")=now
		rs("isprice")=isprice
		rs("play")=play
		rs("tips")=tips
	    rs("addfabu")=addfabu
		rs("Limit")=Limit
		rs("ping")=ping
        rs("pingtxt")=pingtxt
		rs("jieducm_fb")=jieducm_fb
		rs("IfComeSet")=IfComeSet
        rs("ComeKey")=ComeKey
        rs("ComeNote")=ComeNote
    	rs.update
    	rs.close

      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
		rs("fabudian")=rs("fabudian")-fabu
    	rs("cunkuan")=rs("cunkuan")-price
    	rs.update
    	rs.close
		
call check_jieducm_name(session("useridname"))			
sql="update "&jieducm&"_depot set num=num+1 ,now='"&now()&"'  where id="&id&""
conn.execute sql

     Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=pid
		rs("class")="淘宝区任务"
		rs("content")=session("useridname")&"从任务仓库中发布任务"&pid&"成功,发布点减少了"&fabu&"个"
		rs("price")="-"&price
		rs("jiegou")="成功"
    	rs.update
    	rs.close
	Response.Write("<script>alert('发布成功!');window.location.href='../taobao/MyMission.asp';</script>")
	response.End()
end if
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">    
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->

<div align="center">
<DIV 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
<UL id=task_nav>
  <LI><A  href="index.asp">淘宝互动区</A> </LI>
  <LI><A  href="<%if vip="1" then%>pushmissionvip.asp<%else%>pushmission.asp<%end if%>">发布任务</A> </LI>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
  <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="MySave.asp">任务仓库</A> </LI> <li><a href="../user/Express.asp">生成快递单</a> </li>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 0px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 125px; COLOR: #006600; TEXT-ALIGN: center">任务ID</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 95px; COLOR: #006600; TEXT-ALIGN: center">价格</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 95px; COLOR: #006600; TEXT-ALIGN: center">商品信息</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 155px; COLOR: #006600; TEXT-ALIGN: center">备注 </DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 140px; COLOR: #006600; TEXT-ALIGN: center">上次发布时间 
</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 130px; COLOR: #006600; TEXT-ALIGN: center">已经发布</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 150px; COLOR: #006600; TEXT-ALIGN: center">操&nbsp;&nbsp;作</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
 <%
	Sql = "select * from "&jieducm&"_depot where classid='1'  and username='"&session("useridname")&"' order by id desc"
   Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	url="mysave.asp"
	 rs.pagesize=10
	 PageNo=REQUEST("PageNo")
	 if PageNo="" or PageNo=0 then PageNo=1
	 RS.AbsolutePage=PageNo
	 TSum=rs.pagecount
	 maxperpage=rs.pagesize
	 RowCount=rs.PageSize
	   PageNo=PageNo+1
	   PageNo=PageNo-1
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  response.Write("对不起没有您想要的页数")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("没有这一页!")
		Response.End
	 End if
	
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	DO WHILE NOT rs.EOF AND RowCount>0
	IfComeSet=rs("IfComeSet")
	  %>
			
<DIV class=missionbg 
style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 90px">
<DIV style="FLOAT: left; WIDTH: 125px;LINE-HEIGHT: 150%;  TEXT-ALIGN: center">
<%if IfComeSet>0 then response.Write("<img src=../images/jieducm_lailu.gif border='0' title='来路搜索任务'><br>") end if%>
<%=rs("title")%><br>
<%=rs("pid")%>  <br>
备注：发布点<%
price=rs("price")
bei=rs("bei")
addfabu=rs("addfabu")
jieducm_fb=rs("jieducm_fb")
call jieducm_fadian(jieducm_fb,addfabu)
 %>个<br>
   </DIV>
<DIV style="FLOAT: left; WIDTH: 95px; LINE-HEIGHT: 150%; TEXT-ALIGN: center"><font color=#ff6600>平台担保<br>
      <img src="../images/j_z.gif" width="13" height="16"><%=rs("Price")%>元<br><%=rs("isprice")%></font></DIV>
<DIV style="FLOAT: left; WIDTH: 95px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<span style="FLOAT: left;LINE-HEIGHT: 150%; WIDTH: 95px;  TEXT-ALIGN: center">
<input name="<%=rs("pid")%>" type="text" id="<%=pid%>" size="9" value="<%=rs("prourl")%>">
<INPUT style="WIDTH: 90px" onClick="jsCopy('<%=rs("pid")%>')" type=button value=复制商品地址>
 <BR>
掌柜：<%=rs("Shopkeeper")%></span></DIV>
<DIV style="FLOAT: left; WIDTH: 155px; TEXT-ALIGN: center">
  <DIV style="CLEAR: both; WIDTH: 100%; HEIGHT: 20px"><%=rs("bei")%></DIV>
</DIV>
<DIV style="FLOAT: left; WIDTH: 140px; LINE-HEIGHT: 140%; TEXT-ALIGN: center"><%=rs("now")%></DIV>
<DIV style="FLOAT: left; WIDTH: 130px; LINE-HEIGHT: 120%; TEXT-ALIGN: center"><B 
style="COLOR: red"><%=rs("num")%> 个</B></DIV>
<DIV style="FLOAT: left; WIDTH: 150px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">
<a href="#1" onClick="javascript:showDialog('你确认要继续发布此任务了吗？',1,7,'?id=<%=rs("id")%>&action=ok')" title="你确认要继续发布此任务了吗！">
<SPAN class=anniu>继续发布</SPAN></A> 
<a href="#1" onClick="javascript:showDialog('你确认要删除此任务了吗？',1,7,'?id=<%=rs("id")%>&action=del')" title="你确认要删除此任务了吗！">
<SPAN class=anniu>删除</SPAN></A><BR><BR>
<a href="#1" onClick="javascript:showDialog('你确认要重新编辑此任务吗？',1,7,'MySavePush.asp?id=<%=rs("id")%>')" title="你确认要重新编辑此任务吗！">
<SPAN class=anniu>重新编辑</SPAN></A>
</DIV>
</DIV>
<%RowCount = RowCount - 1
i=i-1
rs.MoveNext 
Loop %>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></DIV></DIV>
</div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</BODY></HTML>
