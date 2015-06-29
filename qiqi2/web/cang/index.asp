<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
Sqls_jieducm = "select * from "&jieducm&"_jieshou where start='1' and classid='6'"
Set jieducm_Rsss = Server.CreateObject("Adodb.RecordSet")
jieducm_Rsss.Open Sqls_jieducm,conn,1,1
if not jieducm_Rsss.eof then
Do While Not jieducm_Rsss.Eof
now_jieducm=jieducm_Rsss("now")
jieducm_time= DateDiff( "n",now_jieducm, Now())
if 20-jieducm_time<1 then

sqlr="update "&jieducm&"_zhongxin set  jieshou_num=jieshou_num-1 where  pid='"&jieducm_Rsss("pid")&"'"
conn.execute sqlr

sqlrs="delete  from "&jieducm&"_jieshou where id='"&jieducm_Rsss("id")&"'"
conn.execute sqlrs
end if
jieducm_Rsss.MoveNext
Loop
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 1 num from "&jieducm&"_jieshou where username='"&session("useridname")&"' and classid='6'  order by  id desc",conn,1,1
if not (rs.eof) then
  checknum=rs("num")
rs.close
end if
%>
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="../css/global.css" type=text/css rel=stylesheet>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT language=JavaScript src="../js/jquery.min.js"></SCRIPT>
<META content="MSHTML 6.00.2900.5921" name=GENERATOR></HEAD>
<BODY>
<SCRIPT language=JavaScript type=text/javascript>

function showxiao(id,stat){
var idxiao=id;
sAlert("<div style=margin:10px;line-height:40px;><FORM id=xiaohao method=post  name=formf><div style=line-height:40px;>  【已绑定的淘宝收藏小号】<br><%	
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=4  order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
				    wei=0
					Response.Write("<a href=buyhao.asp>对不起，你还没有绑定买号！绑定买号后，才可以接手任务!</a>")
				Else
				    wei=1
					Do While Not Rs.Eof
			  %> <input type=radio name=xiaohaoname <%if checknum=rs("shopname") then%> checked <%end if%> value=<%=rs("shopname")%> /><%=rs("shopname")%>&nbsp;&nbsp; <%
			  	Rs.MoveNext
				Loop
				End IF
			  %><br><%if wei=1 then %><INPUT type=submit value=确定 onClick=this.disabled=true;document.formf.submit();>&nbsp;<input type=button value=关闭 id=\"do_OK\" onclick=\"doOk()\" /><%end if%><%if wei=0 then%><input type=button value=关闭 id=\"do_OK\" onclick=\"doOk1()\" /><% end if%></div><div style=line-height:40px;></div></FORM></div>");
document.getElementById('xiaohao').action='jieducm_jieshou.asp?id='+idxiao;
}

function sAlert(txt){
    var eSrc=(document.all)?window.event.srcElement:arguments[1];
    var shield = document.createElement("DIV");
    shield.id = "shield";
    shield.style.position = "absolute";
    shield.style.left = "0px";
    shield.style.top = "0px";
    shield.style.width = "100%";
    shield.style.height = document.body.scrollHeight+"px";
    shield.style.background = "#333";
    shield.style.textAlign = "center";
    shield.style.zIndex = "10000";
    shield.style.filter = "alpha(opacity=0)";
    shield.style.opacity = 0;
    var alertFram = document.createElement("DIV");
    alertFram.id="alertFram";
    alertFram.style.position = "absolute";
    alertFram.style.left = "50%";
    alertFram.style.top = "50%";
    alertFram.style.marginLeft = "-225px" ;
    alertFram.style.marginTop = -75+document.documentElement.scrollTop+"px";
    alertFram.style.width = "500px";
    alertFram.style.height = "100px";
    alertFram.style.background = "#ccc";
    alertFram.style.textAlign = "center";
    alertFram.style.lineHeight = "100px";
    alertFram.style.zIndex = "10001";
    strHtml = "<ul style=\"list-style:none;margin:0px;padding:0px;width:100%\">\n";
    strHtml += "    <li style=\"background:#DD828D;text-align:left;padding-left:20px;font-size:14px;font-weight: bold;color:#fff;height:25px;line-height:25px;border:1px solid #F9CADE;\">[请选择你收藏此商品的淘宝用户名：]</li>\n";
    strHtml += "    <li style=\"background:#fff;text-align:center;font-size:12px;height:120px;line-height:80px;border-left:1px solid #F9CADE;border-right:1px solid #F9CADE;\">"+txt+"</li>\n";
    strHtml += "    <li style=\"background:#FDEEF4;text-align:left;font-weight:bold;height:25px;line-height:25px; border:1px solid #F9CADE;\"><span style='color: #0000FF;'>官方提示：请严格按卖家提示及买号要求操作，否则勿接！违反规定者封号！</span><br\>\n";
    strHtml +="";
    strHtml +="</li>\n";
    
    strHtml += "</ul>\n";
    alertFram.innerHTML = strHtml;
    document.body.appendChild(alertFram);
    document.body.appendChild(shield);
    this.setOpacity = function(obj,opacity){
        if(opacity>=1)opacity=opacity/100;
        try{ obj.style.opacity=opacity; }catch(e){}
        try{ 
            if(obj.filters.length>0&&obj.filters("alpha")){
                obj.filters("alpha").opacity=opacity*100;
            }else{
                obj.style.filter="alpha(opacity=\""+(opacity*100)+"\")";
            }
        }catch(e){}
    }
    var c = 0;
    this.doAlpha = function(){
        if (++c > 20){clearInterval(ad);return 0;}
        setOpacity(shield,c);
    }
    var ad = setInterval("doAlpha()",1);
    this.doOk = function(){
        //alertFram.style.display = "none";
        //shield.style.display = "none";
        document.body.removeChild(alertFram);
        document.body.removeChild(shield);
        eSrc.focus();
        document.body.onselectstart = function(){return true;}
        document.body.oncontextmenu = function(){return true;}
    }
	    this.doOk1 = function(){
        //alertFram.style.display = "none";
        //shield.style.display = "none";
        document.body.removeChild(alertFram);
        document.body.removeChild(shield);
        eSrc.focus();
        document.body.onselectstart = function(){return true;}
        document.body.oncontextmenu = function(){return true;}
		window.location.href='buyhao.asp'
    }
    document.getElementById("do_OK").focus();
    eSrc.blur();
    document.body.onselectstart = function(){return false;}
    document.body.oncontextmenu = function(){return false;}
}

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
   $.post('/cang/jieducm-ajax.asp?open=indexajex',function (data) {
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
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="index.asp">淘宝收藏区</A> </LI>
  <LI><A  href="pushmission.asp">发收藏任务</A> </LI>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
  <LI><A href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A href="MySave.asp">任务仓库</A> </LI>
  </UL>
  </DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV 
style="BACKGROUND: url(../images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; HEIGHT: 38px">
<DIV style="MARGIN-TOP: 10px; FLOAT: left"><IMG 
src="../images/task_02.gif"></DIV>
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px">
<A href="?sort=0"><SPAN class=anniu>所有未完成任务</SPAN></A>
<A href="?sort=1"><SPAN class=anniu>宝贝收藏任务</SPAN></A>
<A href="?sort=2"><SPAN class=anniu>店铺收藏任务</SPAN></A></DIV>
<DIV style="CLEAR: right; MARGIN-TOP: 12px; FLOAT: right">
<A title=点击刷新 href="javascript:void(0)" onClick="reloadTask()" class="yell_font"> <SPAN class=anniu id=shuaimg>刷新页面   任务不断跳出</SPAN></A>
</DIV></DIV></DIV>


<div style="width:512px;height:32px;z_index:5px; display:none; position:fixed; _position:absolute;" id="waitingimage"><img src="../images/jieducm_shua.gif"/></div>
<div id="content">
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="CLEAR: both; WIDTH: 930px; BACKGROUND-COLOR: #ffffff">
<DIV style="CLEAR: both; WIDTH: 100%">
<DIV style="CLEAR: both; WIDTH: 100%; HEIGHT: 45px">
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">任务编号</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">发布人</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">任务价格</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">发布时间</DIV>
<DIV class=c_01 style="WIDTH: 20%"><IMG src="../images/task_01.gif">收藏进度</DIV>
<DIV class=c_01 style="CLEAR: right; WIDTH: 20%"><IMG src="../images/task_01.gif">操&nbsp;&nbsp;作</DIV></DIV></DIV></DIV></DIV>
<DIV id=missionset style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 950px"></DIV>
 <%	 
action=request("sort")
if action="" then
  Sql = "select top 100  username,price,pid,bei,start,now,num,jieshou_num from "&jieducm&"_zhongxin where classid='6'  order by (num-jieshou_num) desc"
elseif action=0 then
  Sql = "select top 100 username,price,pid,bei,start,now,num,jieshou_num from "&jieducm&"_zhongxin where  classid='6' and jieshou_num<num  order by now desc"
elseif action=1 then 
  Sql = "select top 100 username,price,pid,bei,start,now,num,jieshou_num from "&jieducm&"_zhongxin where  classid='6' and  bei='宝贝收藏任务' order by id desc"
elseif action=2 then
  Sql = "select top 100 username,price,pid,bei,start,now,num,jieshou_num from "&jieducm&"_zhongxin where   classid='6' and bei='店铺收藏任务' order by id desc"
ElseIf action=6 Then
  Sql = "select top 100  username,price,pid,bei,start,now,num,jieshou_num from "&jieducm&"_zhongxin where classid='6' order by (num-jieshou_num) desc"

end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then				
 	if action="" then 
	  url="index.asp"
	else
	  url="index.asp?sort="&action&""
	end if
	 rs.pagesize=15
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
 bei=rs("bei")
 pid=rs("pid")
 now1=rs("now")
 username=rs("username")
 num=rs("num")
 jieshou_num=rs("jieshou_num")
 %>  
<DIV class=tt5>
<TABLE style="MARGIN: 3px" cellSpacing=2 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD width="15%"><%=pid%><BR><%=bei%></TD>
    <TD>
    <TD align=middle width="15%"><SPAN  style="Z-INDEX: 20;">
<%
Sql2 = "select jifei,qq,vip from "&jieducm&"_register where username='"&username&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then	
jifei=rs2("jifei")
vip=Replace(Trim(rs2("vip")),"'","''")	
qq=rs2("qq")
rs2.close
end if				
 %>
<%call isonline(username)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=username%>" title="发送站内信息" target="_blank"><%=username%></a></SPAN>
<%if vip="1" then %><img src="../images/jieducm_vip.gif"  alt="尊贵VIP,信誉值：<%call vipxinyu(username)%>"  border="0"/><% end if%>
<br />
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0" />
<BR>
<%call xinyu(jifei)%>
		 </TD>
    <TD align=middle width="15%"><img src="../images/jieducm_zf.gif" width="13" height="16"><font color="#FF0000"><strong>
	<%if num*stupcang<1 then
	response.Write("0")
	response.Write(num*stupcang)
	else
	response.Write(num*stupcang)
	end if
	%>
	</strong></font>个发布点</TD>
    <TD align=middle width="15%">
<%=now1%></TD>
    <TD style="PADDING-LEFT: 50px" width="20%">	
	<STRONG><%=jieshou_num%>/<%=num%></STRONG> 
	<DIV class=speedjieducm><DIV class=rate style="WIDTH: <%=jieshou_num/num*100%>%;  float: left;"></DIV>
	</DIV>
	</TD>
    <TD style="PADDING-LEFT: 50px" width="20%"> 
<%
if  username<>session("useridname") and jieshou_num<num then%>
<a title="接手，并完成任务可获得存款和发布点" style="CURSOR: pointer" onClick="showxiao('<%=pid%>','2')">
<img src=../images/online_admin.gif><span class=anniu>接手任务</span></a>
<%elseif jieshou_num=num then
response.Write("任务已结束")
else
response.Write("不能接自己的任务")
end if%>
</TD></TD></TR></TBODY></TABLE></DIV>
<%
	  RowCount = RowCount - 1
	  i=i-1
      rs.MoveNext 
      Loop
%>  
<DIV class=tt5 >  <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></div>
</div>  
</div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
