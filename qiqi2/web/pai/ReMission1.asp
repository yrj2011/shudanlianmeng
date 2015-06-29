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
Sql = "select * from "&jieducm&"_jieshou where classid='2'  and username='"&session("useridname")&"'  order by start asc ,id desc"
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
price=rs("price")
bei=rs("bei")
now1=rs("now")
start=rs("start")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
num=rs("num")
baohu2=rs("baohu2")

Sql2 = "select * from "&jieducm&"_zhongxin where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then
play=rs2("play")
nowf=rs2("now")
usernamef=rs2("username")
tips=rs2("tips")
isprice=rs2("isprice")
jieducm_fb=rs2("jieducm_fb")
addfabu=0
end if
 %>	  
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; HEIGHT: 70px">
<DIV style="FLOAT: left; WIDTH: 125px; LINE-HEIGHT: 150%; TEXT-ALIGN: left"><%=pid%><BR>备注：
<%
if bei="马上收货好评"  then
response.Write("<IMG alt=立即收货评价 src=../images/jieducm_xuni.gif>")
j=0
elseif bei="一天后收货好评"  then
j=24
response.Write("<IMG alt=延迟收货 src=../images/jieducm_shiwu.gif>")
elseif bei="二天后收货好评"  then
j=48
response.Write("<IMG alt=延迟收货 src=../images/jieducm_shiwu.gif>")
elseif bei="三天后收货好评"  then
j=72
response.Write("<IMG alt=延迟收货 src=../images/jieducm_shiwu.gif>")
end if
response.Write("发布点")
call jieducm_fadian(jieducm_fb,addfabu)%>个<BR><%=nowf%>
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
<% call isonline(usernamef)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=usernamef%>" title="发送站内信息" target="_blank"><%=usernamef%></a>
<%if vipf="1" then %><img src="../images/jieducm_vip.gif"  alt="尊贵VIP,信誉值：<%call vipxinyu(usernamef)%>"  border="0"/><% end if%>
<br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes">
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0"  />
</a>
<%call xinyu(jifei)%><br>
[<a href="../user/tousu.asp?pid=<%=pid%>&usernamet=<%=usernamef%>" target="_blank"><font color="#FF0000">申诉</font></a>]
[<a href="../user/name.asp?heiname=<%=usernamef%>" target="_blank"><font color="#FF0000">拉黑</font></a>]<br />
[<a href="../user/sms.asp?fausername=<%=usernamef%>" target="_blank"><font color="#FF0000">发手机短信</font></a>]
</DIV>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px;  TEXT-ALIGN: center"><font color=#ff6600>平台担保<br>
  <img src="../images/j_z.gif" width="13" height="16"><%=rs("Price")%>元</font><br><%=isprice%>
   
  </DIV>
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 100%; TEXT-ALIGN: center">
    <%if baohu2="2" or baohu2="0" then%>
  <label>
  <input name="<%=pid%>" type="text" id="<%=pid%>" size="9" value="<%=prourl%>">
  <br>
  <INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=复制商品地址>
  <br>
  </label>
掌柜：
<%
response.Write(Shopkeeper)
else
response.Write("此任务受到保护，请QQ联系发布方，审核后才能看见商品地址")
end if%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 140px; LINE-HEIGHT: 150%; TEXT-ALIGN: center"><%=now1%><BR>
<font color="#FF0000">接手拍拍号：</font><br>
<%=num%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=0 then%>
<B style="COLOR: red">等待接手人...  </B>
<%elseif start=1 then%>
<B style="COLOR:#006600">已接手，<BR>等待接手方支付... </B>
<%elseif start=2 then%>
<B style="COLOR: #80b309">已支付，<BR>等待发布方发货...  </B>
<%elseif start=3 then%>
<B style="COLOR:blue">已发货，<BR>等待接手方好评...  </B>
<%elseif start=4 then%>
<B style="COLOR: #000000">已好评，<BR>等待发布方好评... </B>
<%elseif start=5 then%>
<B style="COLOR:#006600">完成！获得好评 </B>
<%end if%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 160px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
sss= DateDiff( "n",now1, Now())
if 20-sss>=1 then
if baohu2="1" then
response.Write("此任务受到保护，请QQ联系发布方，审核后才能看见商品地址")
elseif baohu2="2"  or baohu2="0" then%>

<a href="#1" onClick="javascript:showDialog('你确认你已经按任务价在拍拍上用财付通付款了吗？\n付款后受到平台保护！',1,7,'jieducm_zhifu.asp?id=<%=pid%>')" title="请在20分钟内进入商品地址用财付通支付！支付后务必点击已经支付！">
<SPAN class=anniu>已经支付</SPAN>
</A>
<BR> 
剩余<span style="color: #ff0000"><span id="min1"><%=20-sss%></span> </span>分钟可支付
<%end if%>

<BR>联系发布方可加时！<BR>
<a href="#1" onClick="javascript:showDialog('你确认要退出该任务了吗？',1,7,'jieducm_remis.asp?id=<%=pid%>')" title="你确认要退出该任务了吗？">
<SPAN class=anniu>退出任务</SPAN></A>
<%else%>
<SPAN class=anniu>
<%call jieducm_exitauto(pid,session("useridname"))%></SPAN><BR> 
联系发布方可加时！<BR>
<a href="#1" onClick="javascript:showDialog('你确认要退出该任务了吗？',1,7,'jieducm_remis.asp?id=<%=pid%>')" title="你确认要退出该任务了吗？"><SPAN class=anniu>退出任务</SPAN></A>
<%end if
elseif start=2 then%>
<a href="#1" onClick="javascript:showDialog('你确认返回未支付吗？\n任务将返回未支付状态！',1,7,'jieducm_weizhifu.asp?id=<%=pid%>')" title="如果你还未支付，请返回未支付状态！">
<SPAN class=anniu>并未支付</SPAN></A>
<BR>如果你还未支付
请点击 并未支付！
提交不真实封ID<BR>
<%elseif start=3 then
if j=0 then%>
<a href="#1" onClick="javascript:showDialog('你确认卖方已发货了吗？',1,7,'jieducm_haoping.asp?id=<%=pid%>')" title="你确认卖方已发货了吗？">
<SPAN class=anniu>我已好评</SPAN></A>
<%else
nowj=rs("now")
sss= DateDiff( "h", nowj, Now())
if cint(j)-cint(sss)<1 then%>
<a href="#1" onClick="javascript:showDialog('你确认卖方已发货了吗？',1,7,'jieducm_haoping.asp?id=<%=pid%>')" title="你确认卖方已发货了吗？">
<SPAN class=anniu>我已好评</SPAN></A>
<%else
response.Write("不能立即收货<br>还剩")
response.Write(cint(j)-cint(sss))
response.Write("小时")
end if
end if
elseif start=4 then%>
立刻通知发布方确认<br>好评！并任务完成
<%end if%></DIV>
</DIV>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 10px; PADDING-TOP: 10px;  TEXT-ALIGN: center">  
<div align="left">店铺动态打分：<%=play%>, <%=bei%>!</div> 
</DIV>

 <%if tips<>"" then%>
 <DIV style="WIDTH: 98%; LINE-HEIGHT: 10px; PADDING-TOP: 10px;  TEXT-ALIGN: center">  
 <div align="left">  
 卖家提示：<font color="#FF0000"><%=tips%></font>
 </div></DIV>
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
