<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
<SCRIPT language=JavaScript src="../js/jquery.min.js"></SCRIPT>
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
   $.post('/cang/jieducm-ajax.asp?open=Reajex',function (data) {
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
<!--#include file=../taobao/jieducm_toptao.asp-->
<div align="center">
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
<UL id=task_nav>
  <LI><A  href="index.asp">淘宝收藏区</A> </LI>
  <LI><A  href="pushmission.asp">发收藏任务</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="ReMission.asp">已接任务</A> </LI>
  <LI><A  href="MyMission.asp">已发任务</A> </LI>
    <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A href="MySave.asp">任务仓库</A> </LI>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>

<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="BACKGROUND: url(../images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; HEIGHT: 38px">
<DIV style="MARGIN-TOP: 10px; FLOAT: left"><IMG 
src="../images/task_02.gif"></DIV>
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px">
<A href="?sort=6"><SPAN class=anniu>我接手进行中任务</SPAN></A>
<A href="?sort=5"><SPAN class=anniu>已完成</SPAN></A>
</DIV>
<DIV style="CLEAR: right; MARGIN-TOP: 12px; FLOAT: right"><A title=点击刷新 href="javascript:void(0)" onClick="reloadTask()"><SPAN class=anniu id=shuaimg>刷新页面  任务不断跳出</SPAN></A></DIV></DIV></DIV>
<div style="width:512px;height:32px;z_index:5px; display:none; position:fixed; _position:absolute;" id="waitingimage"><img src="../images/jieducm_shua.gif"/></div>
<div id="content">
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 5px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 125px; COLOR: #006600; TEXT-ALIGN: center">任务ID</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">发布人</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">任务价格</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">商品信息</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 140px; COLOR: #006600; TEXT-ALIGN: center">接手时间 
</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 120px; COLOR: #006600; TEXT-ALIGN: center">状&nbsp;&nbsp;态</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 160px; COLOR: #006600; TEXT-ALIGN: center">操&nbsp;&nbsp;作</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
<%
action=request.QueryString("sort")
if action="" then
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"'  order by start asc ,id desc"
elseif action="6" then
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"'  order by start asc"
else
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"' and start='"&action&"' order by start  asc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then
 response.write("暂无接手信息")
 end if
if not rs.eof then				
 	if action="" then
	 url="remission.asp"
	else
	 url="remission.asp?sort="&action&""
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
now1=rs("now")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
num=rs("num")
start=rs("start")
id=rs("id")
Sql2 = "select * from "&jieducm&"_zhongxin where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then
play=rs2("play")
nowf=rs2("now")
usernamef=rs2("username")
tips=rs2("tips")
bei=rs2("bei")
end if
 %>	  
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; HEIGHT: 70px">
<DIV style="FLOAT: left; WIDTH: 125px; LINE-HEIGHT: 150%; TEXT-ALIGN: left"><%=pid%>
<BR><%=bei%>
<BR>
<%=nowf%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; TEXT-ALIGN: center">
<DIV style="OVERFLOW: hidden; WIDTH: 100%;  TEXT-ALIGN: center">
<%	Sql = "select jifei,qq,vip from "&jieducm&"_register where username='"&usernamef&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql,conn,1,1
if not rs1.eof then
jifei=rs1("jifei")
qq=rs1("qq")
vipf=Replace(Trim(rs1("vip")),"'","''")
rs1.close
end if
 %>
<%call isonline(usernamef)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=usernamef%>" title="发送站内信息" target="_blank"><%=usernamef%></a> 
<%if vipf="1" then %><img src="../images/jieducm_vip.gif"  alt="尊贵VIP,信誉值：<%call vipxinyu(usernamef)%>"  border="0"/><% end if%>
<br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes">
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0"  />
</a>
<%call xinyu(jifei)%><br>
</DIV>

</DIV>
<DIV style="FLOAT: left; WIDTH: 115px;  TEXT-ALIGN: center"><font color=#ff6600>平台担保<br>
  <img src="../images/jieducm_zf.gif" width="13" height="16"><%=stupcang%>个发布点</font><br>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 100%; TEXT-ALIGN: center">
    
  <label>
  <input name="<%=pid%>" type="hidden" id="<%=pid%>" size="9" value="<%=prourl%>">
  <INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=复制商品地址>
  <br>
  </label>
  <a onClick="javascript:clipboardData.setData('text','<%=tips%>');alert('标签已成功复制');return false;" href="#">→复制标签←</a> </DIV>
<DIV style="FLOAT: left; WIDTH: 140px; LINE-HEIGHT: 150%; TEXT-ALIGN: center"><%=now1%><BR>
<font color="#FF0000">接手淘宝号：</font><br><%=num%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
response.Write("<B style='COLOR: red'>等待收藏...  </B>")
elseif start=2 then
response.Write("<B style='COLOR:#006600'>已完成收藏 </B>")
end if
%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 160px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
sss= DateDiff( "n",now1, Now())
if 20-sss>=1 then
%>
<a href="#1" onClick="javascript:showDialog('你确认你已经收藏了吗！',1,7,'jieducm_ok.asp?id=<%=pid%>')" >
<SPAN class=anniu>已经收藏</SPAN>
</A>
<BR> 
剩余<span style="color: #ff0000"><span id="min1"><%=20-sss%></span> </span>分钟可收藏
<%else

sqlr="update "&jieducm&"_zhongxin set  jieshou_num=jieshou_num-1 where  pid='"&pid&"'  and classid='6'"
conn.execute sqlr

sqlrs="delete  from "&jieducm&"_jieshou where id='"&id&"'  and classid='6'"
conn.execute sqlrs

end if%>

<BR>联系发布方可加时！<BR>
<a href="#1" onClick="javascript:showDialog('你确认要退出该任务了吗？',1,7,'jieducm_remis.asp?id=<%=id%>')" title="你确认要退出该任务了吗？">
<SPAN class=anniu>退出任务</SPAN></A>
<%end if%>

</DIV>
</DIV>
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
</div></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</BODY></HTML>
