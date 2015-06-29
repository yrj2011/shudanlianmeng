
<LINK href="../hs/main.css" type=text/css rel=stylesheet>
<%
dim restr
restr="index.asp"'文件必须存在当前文件夹时
path= replace(replace(server.mappath(restr),server.mappath("/"),""),restr,"")
path1=replace(path,"\","")
call alipayto() 
Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
File=Left(FileName,InstrRev(FileName,".")-1)
%>
<%
Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
call alipayto() 
File=Left(FileName,InstrRev(FileName,".")-1)
%>
<script language="JavaScript" src="/hs/min.js"></script>
<script language="JavaScript" src="/hs/mask.js"></script>
<div id="sitenav"><div class="info"><a class="db_d"  onclick='$("#sitenav").slideToggle("slow")' style="cursor:pointer"></a><div class="db_z"><span style="font-weight:800;color: #ff0000">温馨提示:</span> <MARQUEE onmouseover=this.stop()  onmouseout=this.start() scrollAmount=2 direction=left width="400" height=28><font color=#000000>公告：为了每个会员的利益和安全保证，即日起禁止小号用旺旺软件登陆，发现一次扣除5个发布点,所有小号均使用网页登陆，禁止旺旺登陆小号，旺旺登陆过的小号，禁止接手任务，谢谢合作！！</font></a>
    </MARQUEE> </div><div class="db_y"><span style="color: #F6400">当前在线人数：<%=Application("OnLine")%>人&nbsp;&nbsp;欢迎您！</span>   <%if session("useridname")="" then%>
<A href="http://www.hushuawang.com/login.asp" target="_top">登陆</A> <A  href="http://www.hushuawang.com/register.asp" target="_top"> 注册</A> 
<%else
Sql = "select fabudian,cunkuan,jifei,vip,tribeid from "&kuaishua&"_register where username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	response.Redirect("login.asp")				
Else
    fabudian=rs("fabudian")
	cunkuan=rs("cunkuan")
	jifei=rs("jifei")
	vip=Replace(Trim(rs("vip")),"'","''")
	tribeid=rs("tribeid")
end if
rs.close	
%>
 <A href="user/" target="_top"><SPAN><%=session("useridname")%></SPAN> </A>
 <%if vip="1" then%><img src="../images/vip.gif"  alt="尊贵VIP" /><%end if%>   <A onClick="return confirm('确定退出操作吗？');"  href="http://www.hushuawang.com/user/exit.asp" target="_top">[退出]</A>  
 <% end if%>
 <A href="http://www.hushuawang.com/user/" target="_top"> 管理中心 </A> <A href="#" title=添加到收藏夹 onclick="window.external.addFavorite('<%=stupurl%>','<%=stupname%>')" >收藏本站</A> </div></div></div>
<div class="menu"><div class="logo"><A href="http://www.hushuawang.com" ><img src="http://www.hushuawang.com/uploadpic/201011921135324633.jpg"></A></div><div class="sj"><embed height="30" width="85" type="application/x-shockwave-flash" pluginspage="http://www.adobe.com/go/getflashplayer" quality="high" src="/hs/clock_4.swf?TimeZone=GMT0800" wmode="transparent" style="margin:0; padding:0;"></embed></div>
<div id="quick_menu">
      <ul>
        <li id="nav_index" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.hushuawang.com');return(false);">设为首页</li>
        <li id="nav_favorite" onclick="window.open('http://shuqian.qq.com/post?from=3&title='+encodeURIComponent(document.title)+'&uri='+encodeURIComponent(document.location.href)+'&jumpback=2&noui=1','favit','');void(0)">QQ收藏夹</li>
        <li id="nav_qqgroup" onclick="window.location.href='http://www.hushuawang.com/'">平台Q群</li>
      </ul><div> <a href=""id="showKF"><img src="http://www.hushuawang.com/hs/zx.gif"></a>
      <a href="javascript:alert('恭喜你,你已成功切换到电信线路!');window.location.href='http://www.hushuawang.com/'"><img src="/hs/06.png"/></a>  <img src="/hs/split.gif"/>  <a href="javascript:alert('恭喜你,你已成功切换到网通线路!');window.location.href='http://www.hushuawang.com/'"><img src="/hs/07.png"/></a>  </div>
</div></div>

  <div class="nav">
  <div class="menul"></div><div class="menuc">
  
  <div class="menutab">
<ul>
            <li class="current" ><a href="/"><span>网站首页</span></a></li>
            <li><a href="/taobao/"><span>淘宝互动</span></a></li>
            <li><a href="/pai/"><span>拍拍互动</span></a></li>
            <li><a href="/you/"><span>有啊互动</span></a></li>
            <li><a href="http://www.hushuawang.com//user/MoneyOrPush.asp"><span>赠送兑换</span></a></li>
            <li><a href="/user/login.asp"><span>加入VIP</span></a></li>
            <li><a href="/union/"><span>推广赚钱</span></a></li>
            <li><a href="/tribe/"><span>部落交流</span></a></li> 
            <li style="position:relative;"><a href="/new.asp"><span>新手入门</span></a><img src="/hs/Nddz_04.gif" style="position:absolute; top:-10px; right:0px; z-index:999;" /> </li>
           <li class="kuaisu" onclick="alert('待开放')"> </li>
          </ul>
        </div>
  
  
  </div><div class="menur"></div>
  </div>

 <div class="menud">
  <div class="menudl"></div><div class="menudc">
  <div class="userstau">亲爱的：<FONT color=#ff0000><%=session("useridname")%></FONT>&nbsp; <%if vip="1" then%><img src="../images/vip.gif"  alt="尊贵VIP" /><%end if%>您好！<font color="#FF0000"><%call tribename(tribeid)%> </font>积分:
<FONT color=#ff0000><%=jifei%></FONT> <%call xinyu(jifei)%>  
发布点:<FONT color=#ff0000><%
if (fabudian)=0 then
response.Write("0.0")
else
fabudian=formatnumber((fabudian),1)
response.Write(fabudian)
end  if
%></FONT> 存款:<FONT color=#ff0000><%
if (cunkuan)=0 then
response.Write("0.00")
else
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
end  if
%>
</FONT> 元&nbsp;<font color="#FF0000"><%=weidu%> </font>封私人站内信</div>
  <div class="menud2">您的位置：<a href="#" class="span2">七七互刷平台</a> > <a href="http://www.778892.com/" class="span2">七七首页</a></div>
  
  </div><div class="menudr"></div>
  </div>
	
	


