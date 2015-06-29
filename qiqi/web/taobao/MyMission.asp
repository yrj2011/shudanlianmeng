<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
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

%>
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.5921" name=GENERATOR></HEAD>
<BODY>
<SCRIPT language=JavaScript type=text/javascript>
 var time =  new Date();
   function reloadTask()
   {
   now = new Date();
   if((now-time)<3000)
   {
   alert('间隔3秒,资源宝贵,请不要重复频繁刷新');
   }
   else
   {
   document.all["shuaimg"].innerHTML="数据读取中......";
   time = new Date();
   waitingimagePosition();
   $('#waitingimage').show();
   param = location.search.substring(1);
   if(param!=''){
param ='&'+	param;	   
   }
   $.post('/taobao/jieducm-ajax.asp?open=myajex',function (data) {
$('#content').html(data);
   $('#waitingimage').hide();
    document.all["shuaimg"].innerHTML="刷新页面   任务不断跳出";
   });
   //设置超时后自动关闭载入框 防止出现长时间不返回 用户以为卡住的情况 这样关闭后如果没有新任务，用户会重新点击一次
   setTimeout(function(){$('#waitingimage').hide();}, 1000);
   }
   }
 
function waitingimagePosition(){
  		var obj = $("#content");
 	var offset = obj.offset();
 	var middle = offset.left+440;
 	var down = offset.top+200; 
$("#waitingimage").css({
"position": "absolute",
"top": down,
"left": middle
});
  	}
</SCRIPT>
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->

<div align="center">
  <div style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
    <div style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
      <div style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
        <ul id=task_nav>
          <li><a  href="index.asp">淘宝互动区</a> </li>
          <li><a  href="pushmission.asp">发布任务</a> </li>
          <li><a href="ReMission.asp">已接任务</a> </li>
          <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="javascript:void(0)" onClick="reloadTask()">已发任务</a> </li>
		  <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
          <li><a href="Buyhao.asp">绑定买号</a> </li>
          <li><a href="MySave.asp">任务仓库</a> </li>
		   <li><a href="../user/Express.asp">生成快递单</a> </li>
        </ul>
      </div>
    </div>
    <div style="CLEAR: both; WIDTH: 910px"><img 
src="../images/task_nav_line.gif"></div>
  </div>
  <DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="BACKGROUND: url(images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; HEIGHT: 38px">
<DIV style="MARGIN-TOP: 10px; FLOAT: left"><IMG 
src="../images/task_02.gif"></DIV>
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px"><A href="?sort=0"><SPAN class=anniu>全部等待接手任务</SPAN></A>
<A href="?sort=1"><SPAN class=anniu>已接手</SPAN></A>
 <A href="?sort=2"><SPAN class=anniu>已支付</SPAN></A>
  <A href="?sort=3"><SPAN class=anniu>已发货</SPAN></A>
  <A href="?sort=4"><SPAN class=anniu>已好评</SPAN></A>
<A href="?sort=5"><SPAN class=anniu>已完成</SPAN></A></DIV>
<DIV style="CLEAR: right; MARGIN-TOP: 12px; FLOAT: right"><A title=点击刷新 href="javascript:void(0)" onClick="reloadTask()"><SPAN class=anniu id=shuaimg>刷新页面 
任务不断跳出</SPAN></A></DIV></DIV></DIV>

<div style="width:512px;height:32px;z_index:5px; display:none; position:fixed; _position:absolute;" id="waitingimage"><img src="../images/jieducm_shua.gif"/></div>
<div id="content">
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 0px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 145px; COLOR: #006600; TEXT-ALIGN: center">任务ID</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">价格</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 95px; COLOR: #006600; TEXT-ALIGN: center">商品信息 </DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 120px; COLOR: #006600; TEXT-ALIGN: center">接手方</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 135px; COLOR: #006600; TEXT-ALIGN: center">接手时间 </DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 130px; COLOR: #006600; TEXT-ALIGN: center">状&nbsp;&nbsp;态</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 150px; COLOR: #006600; TEXT-ALIGN: center">操&nbsp;&nbsp;作</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
<%
action=request.QueryString("sort")
if action="" then
Sql = "select * from "&jieducm&"_zhongxin where classid='1'  and username='"&session("useridname")&"'  order by start  asc ,id desc"
elseif action="6" then
Sql = "select * from "&jieducm&"_zhongxin where classid='1'  and username='"&session("useridname")&"'  order by start  asc ,now desc"
else
Sql = "select * from "&jieducm&"_zhongxin where classid='1'  and username='"&session("useridname")&"' and start='"&action&"' order by start  asc ,id desc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	if action="" then
	url="mymission.asp"
	else
	 url="mymission.asp?sort="&action&""
	end if
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
 pid=rs("pid")
price=rs("price")
bei=rs("bei")
start=rs("start")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
play=rs("play")
isprice=rs("isprice")
nowf=rs("now")
addfabu=rs("addfabu")
ping=rs("ping")
pingtxt=rs("pingtxt")
tips=rs("tips")
jieducm_fb=rs("jieducm_fb")
IfComeSet=rs("IfComeSet")
 %> 	  
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; HEIGHT: 90px">
<DIV style="FLOAT: left; WIDTH: 145px;LINE-HEIGHT: 150%;  TEXT-ALIGN: center">
<%if IfComeSet>0 then response.Write("<img src=../images/jieducm_lailu.gif border='0' title='来路搜索任务'><br>") end if%>
<%=pid%>  <br>
备注：<%
if bei="马上收货好评"  then
response.Write("<IMG alt=立即收货评价 src=../images/jieducm_xuni.gif>")
j=0
elseif bei="一天后收货好评"  then
j=24
response.Write("<IMG alt=延迟收货 src=../images/jieducm_shiwu.gif>")
elseif bei="二天后收货好评"  then
j=48
response.Write("<IMG alt=延迟收货 src=../images/jieducm_shiwu.gif>")
elseif bei="三天后收货好评" then
j=72
response.Write("<IMG alt=延迟收货 src=../images/jieducm_shiwu.gif>")
end if
response.Write("发布点")
call jieducm_fadian(jieducm_fb,addfabu)%>个
<br><%=nowf%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; TEXT-ALIGN: center">
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<font color=#ff6600>平台担保<br>
<img src="../images/j_z.gif" width="13" height="16"><%=Price%>元<br><%=isprice%></font>
</DIV>

</DIV>
<DIV style="FLOAT: left; WIDTH: 95px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<input name="<%=pid%>" type="text" id="<%=pid%>" size="9" value="<%=prourl%>">
<INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=复制商品地址>
 掌柜：<%=Shopkeeper%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; TEXT-ALIGN: center">
<%Sql2 = "select * from "&jieducm&"_jieshou where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if rs2.eof then	
jie=0	
response.Write("暂无接手人")

else
baohu2=rs2("baohu2")
jie=1
nowj=rs2("now")
numj=rs2("num")
ipj=rs2("ip")

Sql1 = "select * from "&jieducm&"_register where username='"&rs2("username")&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql1,conn,1,1
if not rs1.eof then
jifei=rs1("jifei") 
username=rs1("username")
%>
<%call isonline(username)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=username%>" title="发送站内信息" target="_blank"><%=username%></a>
 <br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=rs1("qq")%>&site=qq&menu=yes">
<img src="http://wpa.qq.com/pa?p=1:<%=rs1("qq")%>:41"  border="0"  />
</a>
<%
call xinyu(jifei)
if vip="1" then
response.Write("<br>接手方IP:")
response.Write(ipj)
end if
%>
<br><a href="../user/tousu.asp?pid=<%=pid%>&usernamet=<%=username%>" target="_blank"><font color="#FF0000">
[申诉]</font></a>
<a href="../user/name.asp?heiname=<%=username%>" target="_blank"><font color="#FF0000">
[拉黑]</font></a>
<br />[<a href="../user/sms.asp?fausername=<%=username%>" target="_blank"><font color="#FF0000">发手机短信</font></a>]
<%
rs1.close
End IF
rs2.close
end if %>
</DIV>
<DIV style="FLOAT: left; WIDTH: 135px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">

<%if jie=0 then
response.Write("暂无接手人")	
else
response.Write(nowj)
response.Write("<br><font color=#FF0000>接手淘宝号：</font>")
response.Write(numj)
end if
%>

</DIV>
<DIV style="FLOAT: left; WIDTH: 130px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=0 then%>
<B style="COLOR: red">等待接手人... </B>
<%elseif start=1 then%>
<B style="COLOR:#006600">已接手，<BR>等待接手方支付...  </B>
<%elseif start=2 then%>
<B style="COLOR: #80b309">已支付，<BR>等待发布方发货...</B>
<%elseif start=3 then%>
<B style="COLOR:blue">已发货，<BR>等待接手方好评...   </B>
<%elseif start=4 then%>
<B style="COLOR: #000000">已好评，<BR>等待发布方好评... </B>
<%elseif start=5 then%>
<B style="COLOR:#006600">完成！获得好评 </B>
<%end if%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 150px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">
<%if start=0 then%>
<a href="#1" onClick="javascript:showDialog('你确认要取消此任务了吗？',1,7,'jieducm_del.asp?id=<%=pid%>')" title="长时间没人接手，取消重填，&#13;平台将退还你的担保金和发布点，&#13;或者刷新把任务排前！">
<SPAN class=anniu>取消重填</SPAN></A><BR>
取消任务平台将退还你的担保金和发布点，或者刷新把任务排前！<br>
<a href="#1" onClick="javascript:showDialog('你确认要刷新把任务排前吗？',1,7,'jieducm_shuaxin.asp?id=<%=pid%>')" title="长时间没人接手刷新把任务排前！">
<SPAN class=anniu>刷新排前</SPAN></A>
<%elseif start=1 then 
sss= DateDiff( "n",nowj, Now())
if 20-sss<1 then
call jieducm_exitauto(pid,username)
else%>

<%if baohu2="1" then%>
<a href="#1" onClick="javascript:showDialog('你确认要审核通过该接手人吗？\n审核后接手方可以看到商品地址！',1,7,'jieducm_shehe.asp?id=<%=pid%>')" title="审核后接手方可以看到商品地址！">
<SPAN class=anniu>审核接手人</SPAN></A><br>
<%if vip=1 then%>

<a href="#1" onClick="javascript:showDialog('你确认要清理买家吗？\n任务将返回未接手状态！',1,7,'jieducm_qingli.asp?id=<%=pid%>')" title="清理买家！任务返回未接手状态">
<SPAN class=anniu>清理买家</SPAN></A><br>
<%end if%>
<%else%>
<%if vip=1 then%>

<a href="#1" onClick="javascript:showDialog('你确认要清理买家吗？\n任务将返回未接手状态！',1,7,'jieducm_qingli.asp?id=<%=pid%>')" title="清理买家！任务返回未接手状态">
<SPAN class=anniu>清理买家</SPAN></A><br>
<%end if%>
已审核<br>
<%end if%>
剩余<span style="color: #ff0000"><span id="min1"><%=20-sss%></span> </span>分名钟可支付
<%end if%>
<br>
<a href="#1" onClick="javascript:showDialog('你确认要为对方加时吗？',1,7,'jieducm_jiashi.asp?id=<%=pid%>')" title="为对方加时！">
<SPAN class=anniu>为对方加时</SPAN></A>
<%elseif start=2 then%>
<a href="#1" onClick="javascript:showDialog('你确认要进行已发货处理吗？\n认真核对支付方是否支付成功,并且你已发货处理！',1,7,'jieducm_fahou.asp?id=<%=pid%>')" title="认真核对支付方是否支付成功,并且你已发货处理！">
<SPAN class=anniu>我已发货</SPAN></A><BR>核实后请发货并平台处理！
<%elseif start=3 then
if j=0 then
response.Write("联系接手方好评！")
else
sss= DateDiff( "h", nowj, Now())
if cint(j)-cint(sss)<1 then
response.Write("联系接手方好评！<br>收货时间已到")
else
response.Write("接手方不能立即收货<br>还剩")
response.Write(cint(j)-cint(sss))
response.Write("小时")
end if
end if
elseif start=4 then%>
<a href="#1" onClick="javascript:showDialog('你确认要确认好评吗？',1,7,'jieducm_ok.asp?id=<%=pid%>')" title="认真核对支付方是否支付成功,并且你已发货处理!">
<SPAN class=anniu>确认好评</SPAN></A><BR>请立刻确认好评，并处理完成！
<%end if%>
</DIV>
</DIV>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 10px; PADDING-TOP: 10px;  TEXT-ALIGN: center">  
<div align="left">店铺动态打分：<%=play%>!<%=ping%>! <%=bei%>!<font color="#FF0000"> 禁用旺旺联系 </font></div> 
</DIV>

 <%if tips<>"" then%>
 <DIV style="WIDTH: 98%; LINE-HEIGHT: 10px; PADDING-TOP: 10px;  TEXT-ALIGN: center">  
 <div align="left">  
 卖家提示：<font color="#FF0000"><%=tips%></font>
 </div></DIV>
 <%end if%>
  
<%if ping="自定义评语" then%>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 20px; PADDING-TOP: 2px;  TEXT-ALIGN: center">  
<div align="left">  自定义评语：<font color="#FF0000"><%=pingtxt%></font>
</div>
</DIV>
<%end if%> 

<DIV style="WIDTH: 98%; LINE-HEIGHT: 3px; BORDER-BOTTOM: #06314a 1px dashed;">
</DIV>   
<%
RowCount = RowCount - 1
i=i-1
rs.MoveNext 
Loop
%>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center">
 <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></DIV></DIV>
</div>

</div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
