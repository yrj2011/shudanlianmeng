<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#include file="user/checklogin1.asp"-->

<!--#include file="md5.asp"-->
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
action=request.Form("action")
if action="ok" then
call login()
end if
 %>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE><%=stupname%>-<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link rel="shortcut icon" href="/favicon.ico">
<link href="css1/layout.css" type="text/css" rel="stylesheet" />
<link href="css1/reset.css" type="text/css" rel="stylesheet" />
<link href="css1/base.css" type="text/css" rel="stylesheet" />
<link href="cn/main.css" rel="stylesheet" type="text/css" />
<link href="/indexcss/main.css" rel="stylesheet" />
</head>
<script src="/platform/js/ShowQQ.js" language="javascript"></script>
<script src="/platform/User/TaskNotice/js/TaskDialog.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript" src="/platform/nntcn_fun/js/Dialog.js"></script>
<script type="text/javascript" src="/js/jquery.min.js"></script>

<SCRIPT src="/js/jieducm_pupu.js" type=text/javascript></SCRIPT>

<SCRIPT type=text/javascript>  
 
function zOpenD(){   
  
    var diag = new Dialog("Diag2");   
    diag.Width = 600;   
    diag.Height = 400;   
    diag.Title = "人人来网 - 新手互动教学大厅";   
    diag.URL = "/platform/User/TaskNotice/helpcenter.asp";   
    diag.OKEvent = zAlert;//点击确定后调用的方法   
    diag.show();   
   
}   
function zAlert(){   
        Dialog.alert("请点取消退出。");   
}   
    </SCRIPT>
	
<body>
	<!--头部开始-->
	<div id="m_top" style="position:relative; z-index:9999;">
	  <div class="kd">
	     <div class="kmain">
		    <div class="hy"><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1662424999&site=qq&menu=yes" class="dmhtel">在线客服</a><div style="float:left;"><span style="color:#1595DE">|</span> <span style="color:#666">您好，欢迎来到人人来网！</span><a href="/login.asp" class="chengse">[请登录]</a>，新用户？<a href="/register.asp" class="lvse">[免费注册]</a></div></div>
			<div class="top_btn"><a href="/new.asp" style="font-weight:700;" target="_blank"><font color="FE5500">新手帮助</font></a>|<a target="_blank" href="/union/">推广赚钱</a>|<a href="/union/">管理中心</a>|<a href="/user/Listtop.asp">刷客排行</a>|<a href="/user/guest_info.asp">账号设置</a></div>
		  </div>
		    <div id="service_qq" class="help_down" style="display:none;"></div>
	  </div>
	</div>
<img src="images1/br.gif" alt="br" />
	<!--header st-->
    <div class="header">
    	
        <div class="top clearfix">
        	<h1><a href="/">人人来网</a></h1>
          <p><script src="/ShowPloy.asp" type="text/javascript"></script><a href="/news.asp?/1489.html"><img src="images1/xs88.jpg" alt="" border="0" /></a></p>
        </div>
        <div class="nav clearfix">
        	<ul>
            	<li class="current"><a href="/"><em>首页</em></a></li>
                <li><a href="/taobao/"><em>淘宝互动</em></a></li>
                <li><a href="/pai/"><em>拍拍互动</em></a></li>
                <li><a href="/user/soft.asp"><em>软件下载</em></a></li>
                <li><a href="/user/tuoguan.asp"><em>托管专区</em></a></li>
                <li><a href="/user/"><em>管理中心</em></a></li>
                
                <li><a href="/chongzhi"/><em>充值提现</em></a></li>
               
                <li><a href="/union/"><em>推广</em></a></li>
 <li><a href="/tribe/"><em>部落</em></a></li>
                <li><a href="/new.asp"><em>新手入门</em></a></li>
            </ul>
        </div>

    </div>
    <!--header end-->
    <!--S-banner st-->
    <div class="layout-index">
    <!-- 首页幻灯片 -->
    <div class="index-focus">
        <div class="focus-crop">
            <ul class="focus-cnt">
                <li class="panel"><a href="/register.asp" style="background-image:url(images1/S-banner-01.jpg);"></a></li>
                <li class="panel"><a href="/user/" style="background-image:url(images1/S-banner-02.jpg);"></a></li>
                <li class="panel"><a href="/union/" style="background-image:url(images1/S-banner-03.jpg);"></a></li>
            </ul>
        </div>
    </div>
    <script src="js/kissy-min.js"></script>
     <script src="js/switchable-min.js"></script>
   
    <script>
        KISSY.use("event,switchable", function(S, Event,Switchable) {
        S.ready(function(S) {
            var carousel = new Switchable.Carousel('.index-focus', {
                effect: 'scrolly',//切换方式
                easing: 'easeOutStrong',
                navCls: 'focus-nav',
                activeTriggerCls: 'current',
                contentCls: 'focus-cnt',
                prevBtnCls: 'prev',
                nextBtnCls: 'next',
                interval: '3',//3秒切换
                duration: '.3',
                circular: true,//循环播放
                autoplay: true
            });
        });
    });
    </script>

    <!-- 登录/注册框 -->
    <div class="S-login">
            
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
            
            
                <%if session("useridname")="" then%>
            	<p style="margin-bottom:15px;"><a href="/register.asp"><img src="images1/S-reg-btn.gif" alt="" /></a></p>

                 <form name="form" method="post" action=""><input name="action" type="hidden" value="ok">
                <p class="S-login-input">账&nbsp;&nbsp;&nbsp;&nbsp;户：<input name="name" type="text" /></p>
                <p class="S-login-input">密&nbsp;&nbsp;&nbsp;&nbsp;码：<input name="pass" type="password"/></p>
				<p class="S-login-input">验&nbsp;证&nbsp;码：<input  type=text style="width:100px;" maxlength=4 name=CheckCode class=code  onKeyUp="if(isNaN(value))execCommand('undo')">&nbsp;<img src="jieducm_code.asp" alt="验证码" onClick="this.src='jieducm_code.asp?rnd=' + Math.random();" />
                <p class="S-login-btn"><input type="image" src="images1/S-login-btn.jpg"  style="margin-right:10px;"/><input name="" type="checkbox" value="" /> 记住我  <a href="/getpwd.asp" style="color:#3282c1; text-decoration:underline;">忘记密码？</a></p>
                <p class="S-login-txt">一个真实的淘宝信誉互动平台期待你的加入！</p>
                </form>
<%else%>

			  <table class=LeftNews cellSpacing=0 cellPadding=2 width="100%"  border=0>

 <tr>
    <td height="25" colspan="2">会员：<font color=#ff0000><%=session("useridname")%></font>
<%if vip="1" then%><img src="../images/vip.gif"  alt="尊贵VIP，信誉值:<%call vipxinyu(session("useridname"))%>" /><%end if%> 您好！</td>
  </tr>
  
  <tr>
    <td height="25" colspan="2"> 您拥有：<font color=#ff0000>
    <%
if (cunkuan)=0 then
response.Write("0.00")
elseif cunkuan<1 then
response.Write("0")
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
else
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
end if
%>
    </font>
元 发布点：<font color=#ff0000>
<%
if (fabudian)=0 then
response.Write("0.0")
elseif fabudian<1 then
response.Write("0")
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
else
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
end  if%>
</font>点<br>
<tr>
    <td height="25" colspan="2">积分：<font color=#ff0000><%=jifei%></font>
	&nbsp;&nbsp;收藏点：<font color=#ff0000><%
if (cang)=0 then
response.Write("0.00")
elseif cang<1 then
response.Write("0")
cang=formatnumber((cang),2)
response.Write(cang)
else
cang=formatnumber((cang),2)
response.Write(cang)
end  if%></font></td></td>
  </tr>
  </tr>
  <tr>
    <td width="94" height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="user/guest_info.asp"><font color=#ff0000>修改个人信息</font></a></td>
    <td width="102" height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="user/md5_pay.asp"><font color=#ff0000>帐号充值</font></a></td>
  </tr>
   <tr>
     <td height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="user/ReMoney.asp"><font color=#ff0000>我要提现&nbsp; &nbsp; </font></a></td>
     <td height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="user/MoneyOrPush.asp"><font color=#ff0000>我要兑换</font></a></td>
   </tr>
   <tr>
    <td height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="taobao/"><font color=#ff0000>免费刷钻区&nbsp; </font></a></td>
    <td height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a onClick="return confirm('确定退出操作吗？');" href="user/exit.asp"><font color=#ff0000>安全退出</font></a></td>
   </tr></table>   
<%end if%>			  


    </div>

</div>
    
    <!--S-banner end-->
    <!--main st-->
    <div class="main">
	
    	
<div class="g-layout" style="margin-bottom:20px;">
    <div class="bricks">
        <a class="itm1 itm-L" href="union/" tppabs="/Spread/" target="_blank"><img src="images/bricks-itm1.jpg" alt="推广赚钱" border="0" tppabs="/images/bricks-itm1.jpg"></a>      <img src="images_new/bricks-itm2.jpg" tppabs="/images_new/bricks-itm2.jpg" alt="全网借脑">      <img src="images/bricks-itm3.jpg" tppabs="/images/bricks-itm3.jpg" alt="关注我们">
        <a class="itm4 itm-S" href="new.asp" tppabs="/Service/tuoguan.asp" target="_blank"><img src="images_new/bricks-itm4.jpg" alt="新店没成交" border="0" tppabs="/images_new/bricks-itm4.jpg"></a>
        <a class="itm5 itm-S" href="new.asp" tppabs="/Service/peixun.asp" target="_blank"><img src="images_new/bricks-itm5.jpg" alt="总要会一点" border="0" tppabs="/images_new/bricks-itm5.jpg"></a>
        <a class="itm6 itm-M" href="teach.asp" tppabs="/liucheng/liu25.Asp" target="_blank"><img src="images/bricks-itm6.jpg" alt="点击这里看视频教程，轻松学习掌握技能" border="0" tppabs="/images/bricks-itm6.jpg"></a>
        <a class="itm7 itm-M" href="teach.asp" tppabs="/help/sp/" target="_blank"><img src="images/bricks-itm7.jpg" alt="查看图文教程，快速了解谷得网运作模式" border="0" tppabs="/images/bricks-itm7.jpg"></a>    </div>
</div>

<div class="g-layout a-mb">
    <div class="home-section">
        <div class="home-section-hd">
            <h2>服务交易</h2>
        </div>
        <div class="home-section-bd">
            <img src="images_new/home-section2.jpg" tppabs="/images_new/home-section2.jpg" usemap="#map1" style="display:block;"/>
      <map name="map1">
                <area shape="rect" coords="372,106,618,170" alt="发布任务获得精准营销" tppabs="/wk/" target="_blank"/>
                <area shape="rect" coords="666,106,910,170" alt="接手任务赚取佣金" tppabs="/wk/" target="_blank"/>
            <area shape="rect" coords="6,12,303,221" tppabs="/youshi/" target="_blank" />
      </map>
      </div>
  </div>
</div>
<div class="g-layout a-mb">
    <div class="home-section">
        <div class="home-section-hd">
            <h2>卖家服务</h2>
        </div>
        <div class="home-section-bd">
            <a href="help/lianxi.asp" tppabs="/Service/" target="_blank"><img src="images_new/home-section3.jpg" alt="" border="0" style="display:block;" tppabs="/images_new/home-section3.jpg"/></a>        </div>
  </div>
</div>
<div class="g-layout a-mb">
    <div class="home-section">
        <div class="home-section-hd">
            <h2>选择我们</h2>
        </div>
        <div class="home-section-bd">
            <img src="images_new/home-section4.jpg" tppabs="/images_new/home-section4.jpg" usemap="#map3" style="display:block;"/>
<map name="map3">
                <area shape="rect" coords="0,0,378,237" href="help/lianxi.asp" tppabs="/Jiejue/" alt="我们的优势" target="_blank"/>
                <area shape="rect" coords="379,0,960,237" href="help/lianxi.asp" tppabs="/Jiejue/" alt="解决方案" target="_blank"/>
          </map>
      </div>
  </div>
</div>

		
		
        <!--help end-->
    </div>
    <!--main end-->

    <!--footer st-->
    <div class="footer footer-margin">
    	<div style="width:960px; margin:0 auto; text-align:center;">
    	<p><a href="/" style="color:#3282c1;">首页</a> | <a href="/help/about.asp" style="color:#3282c1;">关于我们</a> | <a href="Javascript:openNewWindows('newDiv','/help/lianxi.asp','675','257');" style="color:#3282c1;">联系我们</a> | <a href="/news.asp?/1454.html" style="color:#3282c1;">人人来网规则</a> | <a href="/news.asp?/1410.html" style="color:#3282c1;">人人来网服务协议</a><font style="color:#b9b9b9;">       <a href="http://www.miitbeian.gov.cn/" target="_blank">粤ICP备13034299号</a>    Copyright &copy; 2008 - 2018  ALL Rights Reserved 
</font>
</p>

        </div>
    </div>

<SCRIPT src="/Images/floatadv.js" type=text/javascript></SCRIPT>
<SCRIPT type=text/javascript>
theFloaters.addItem('floatAdv1',6,'document.documentElement.clientHeight-340','<div style="position: absolute; right: 6px; top: 6px;"><img src=\"/Images/gudeF.jpg\" width=\"100\" height=\"232\" border=\"0\" usemap=\"#Map\" /><map name=\"Map\" id=\"Map\"><area shape=\"rect\" coords=\"13,9,80,48\" href=\"#\" /><area shape=\"rect\" coords=\"23,65,66,84\" href=\"#\" onclick=\"openNewWindows(\'newDiv\',\'service/Gude.html\',\'675\',\'257\');\"/><area shape=\"rect\" coords=\"22,86,64,106\" href=\"/platform/index.html\" target=\"_blank\" /></map><br /><img src=\"/images/advclose.gif\" onMouseOver=\"this.style.cursor=\'pointer\'\" onClick=\"closeBanner();\"></div>');theFloaters.play();
</SCRIPT>
    
<script type="text/javascript" src="/js/jquery.slider.js"></script>

    <!-- 下方悬浮工具栏 -->
    <style>
    #bottomBar { display:none; position: fixed; left: 0; bottom: 0; width: 100%; height: 100px; background: rgba(0,0,0,.7); filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#B4000000', EndColorStr='#B4000000');}
    :root #bottomBar { filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#00000000', EndColorStr='#00000000');}
    * html #bottomBar { position: absolute; bottom:auto; top:expression(eval(document.documentElement.scrollTop+document.documentElement.clientHeight-this.offsetHeight-(parseInt(this.currentStyle.marginTop,10)||0)-(parseInt(this.currentStyle.marginBottom,10)||0)));}
    #bottomBar .bottomBar-wrap { width: 960px; margin: 0 auto; padding-top: 30px;}
    #bottomBar .bottomBar-wrap:after { display: block; clear: both; content: '\20';}
    #bottomBar .signin, #bottomBar .signup { _display: inline; overflow: hidden; float: left; width: 131px; height: 50px; margin-left: 25px; background: url(/images/bottomBar.png) no-repeat; text-indent: -99em;}
    #bottomBar .signup { background-position: -131px 0;}
    #bottomBar .flow { overflow: hidden; float: right; width: 382px; height: 39px; margin-right: 45px; background: url(/images/bottomBar.png) no-repeat 0 -50px;}
    #closeBottomBar { display:block; overflow:hidden; position: absolute; right: 42px; top: 21px; width: 59px; height: 59px; background: url(/images/bottomBar-close.png); line-height: 600px; cursor: pointer;}
    </style>
    <div id="bottomBar">
        <div class="bottomBar-wrap">
            <a class="signin" href="login.asp">登录</a>
            <a class="signup" href="register.asp">免费注册</a>
            <span class="flow"></span>
        </div>
        <span id="closeBottomBar" onClick="closeBottomBar()">关闭</span>
    </div>
    <script>
    var closeBottomBar = function () {
        document.getElementById("bottomBar").style.display="none";
    }
    </script>
    <!-- /下方悬浮工具栏 -->
<script>
$(document).ready(function(){
        //首先将#back-to-top隐藏
        //$("#bottomBar").hide();
        //当滚动条的位置处于距顶部100像素以下时，跳转链接出现，否则消失
        $(function () {
            $(window).scroll(function(){
                if ($(window).scrollTop()>100){
                    $("#bottomBar").show();
                }
                else
                {
                    //$("#back-to-top").fadeOut(1500);
                }
            });
        });
    });

</script>



<script>
$(function(){
        //首先将#back-to-top隐藏
	$('#b-stage').cycle({fx:'scrollLeft',pager:'#b-nav',speed:700 });
})
</script>


<style>
#sideBuoy { position: fixed; _position: absolute; z-index: 9; top: 205px; _top: expression(205+documentElement.scrollTop +"px"); left: 50%; width: 81px; height: 173px; margin-left: 500px; background: #82cbec url(/images1/sideBuoy-qq.jpg) no-repeat 0 30px;}
.sideBuoy-jy { display: block; width: 100%; height: 30px; text-align: center; line-height: 30px;}
.sideBuoy-qq { display: block; width: 100%; height: 112px;}
.sideBuoy-top { display: block; width: 100%; height: 31px; text-align: center; line-height: 30px;}
.sideBuoy-detail { display: none; position: absolute; left: -156px; top: 0; overflow: hidden; width: 156px; height: 173px; background: #ffeeda; text-align: center;}
.sideBuoy-detail .qq-list { padding: 8px 0 10px;}
.sideBuoy-detail .qq { height: 22px; padding: 5px 0; line-height: 22px;}
.sideBuoy-detail .qq img { vertical-align: middle;}
</style>
<div id="sideBuoy">
    <a class="sideBuoy-jy" href="/help/viewreturn.asp" tppabs="/help/viewreturn.asp" target="_blank">我有建议</a>
    <a class="sideBuoy-qq" href="#"></a>
    <a class="sideBuoy-top" href="#">返回顶部</a>
    <div class="sideBuoy-detail">
        <div class="qq-list" ID="qqAllList"><div class="qq">新手在线咨询</div>
<div class="qq">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="middle"> 
	   <%
			  	Sql = "select Top 10 * from "&jieducm&"_qq where class=1 order by ID desc"
				Set rs = Server.CreateObject("Adodb.RecordSet")
				rs.Open Sql,conn,1,1
				IF rs.Eof Then
					
				Else
			
					Do While Not rs.Eof
						
			  %>
                       <div class="qq">
		   
		 <a target="_blank"  href="http://wpa.qq.com/msgrd?v=3&uin=<%=rs("qq")%>&site=qq&menu=yes" ><img src="http://wpa.qq.com/pa?p=1:<%=rs("qq")%>:41"  border="0"  /> </A><br />
                       </div>
                    <%
			  	rs.MoveNext
				Loop
				End IF
				rs.close
			  %>
              
	</td>
  </tr>
</table>
</div>
</div>
<br/><br/><div class="service-time">工作时间：9:00~21:00<br>
非工作时间请给我们留言
</div>
    </div>
</div>
<script>
var flag,
    timer;
$('.sideBuoy-qq').hover(function(){
    flag = false;
    //createqq();
    $('.sideBuoy-detail').show();
},function(){
    flag = true;
    timer = setTimeout(function(){
        if(flag){
            $('.sideBuoy-detail').hide();
        }
    },200) 
})
$('.sideBuoy-detail').hover(function(){
    flag = false;
    $(this).show();
},function(){
    var $self = $(this);
    flag = true;
    timer = setTimeout(function(){
        if(flag){
            $self.hide();
        }
    },200) 
})
document.getElementById("t"+"e"+"s"+"i").style.display='none';</script>