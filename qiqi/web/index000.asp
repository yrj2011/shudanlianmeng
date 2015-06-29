<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
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
</head>
<script src="/platform/js/ShowQQ.js" language="javascript"></script>
<script src="/platform/User/TaskNotice/js/TaskDialog.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript" src="/platform/nntcn_fun/js/Dialog.js"></script>
<script type="text/javascript" src="/js/jquery.min.js"></script>
<SCRIPT type=text/javascript>  
 
function zOpenD(){   
  
    var diag = new Dialog("Diag2");   
    diag.Width = 600;   
    diag.Height = 400;   
    diag.Title = "小七网 - 新手互动教学大厅";   
    diag.URL = "/platform/User/TaskNotice/helpcenter.asp";   
    diag.OKEvent = zAlert;//点击确定后调用的方法   
    diag.show();   
   
}   
function zAlert(){   
        Dialog.alert("请点取消退出。");   
}   
    </SCRIPT>

<body>
	<!--header st-->
    <div class="header">
    	<div class="quick-channel">
        	<p><script src="/platform/ShowMyInfo.asp" type="text/javascript"></script></p>
        </div>
        <div class="top clearfix">
        	<h1><a href="/">小七网</a></h1>
            <p><script src="/ShowPloy.asp" type="text/javascript"></script></p>
        </div>
        <div class="nav clearfix">
        	<ul>
            	<li class="current"><a href="/"><em>首页</em></a></li>
                <li><a href="/taobao/"><em>淘宝互动</em></a></li>
                <li><a href="/pai/"><em>拍拍互动</em></a></li>
                <li><a href="/cang/"><em>收藏区</em></a></li>
                <li><a href="/user/tuoguan.asp"><em>托管专区</em></a></li>
                <li><a href="/user/"><em>管理中心</em></a></li>
                <li><a href="/chongzhi"/"><em>充值提现</em></a></li>
               
                <li><a href="/union/"><em>淘宝推广</em></a></li>
 <li><a href="/tribe/"><em>部落</em></a></li>
                <li><a href="/new.asp"><em>新手入门</em></a></li>
            </ul>
        </div>

    </div>
    <!--header end-->
    <!--S-banner st-->
    <div class="S-banner">
    	<div class="S-banner-con">
    		<div class="S-banner-list">
			<div id="slider" class="fl">
				<div id="b-stage" class="b-stage">
			        <div><img src="images1/S-banner-01.jpg" /></div>
			        <div><a href='/register.asp' target="_blank"><img src="images1/S-banner-02.jpg" border='0' /></a></div>
			    </div>
    			<div id="b-nav" class="b-nav"></div>
			</div>
        		
        	</div>
        	<div class="S-login">
                
            	<p style="margin-bottom:15px;"><a href="/register.asp"><img src="images1/S-reg-btn.jpg" alt="" /></a></p>

                 <form name="form" method="post" action=""><input name="action" type="hidden" value="ok">
                <p class="S-login-input">账&nbsp;&nbsp;&nbsp;&nbsp;户：<input name="name" type="text" /></p>
                <p class="S-login-input">密&nbsp;&nbsp;&nbsp;&nbsp;码：<input name="pass" type="password"/></p>
                <p class="S-login-btn"><input type="image" src="images1/S-login-btn.jpg"  style="margin-right:10px;"/><input name="" type="checkbox" value="" /> 记住我  <a href="/platform/login.asp?chuli=zhaomima&urlfan=x" style="color:#3282c1; text-decoration:underline;">忘记密码？</a></p>
                <p class="S-login-txt">一个真实的淘宝信誉互动平台期待你的加入！</p>
                </form>

       	    </div>
            <div class="clear"></div>
        </div>
    </div>
    <!--S-banner end-->
    <!--main st-->
    <div class="main">
    	<div class="S-zhuan"> <a href="/platform/index.html"><img src="images1/S-zhuan.jpg" alt="小七网" /></a>
      </div>
        <div class="S-jiao">
        	<div class="S-jiao-part S-jiao-con">
            	<dl class="S-jiao-video">
                	<dd><a ="/gude18/Contract.asp"><img src="images1/S-video.jpg" alt="淘宝刷钻平台" /></a></dd>
                    <dt><a "/gude18/Contract.asp">视频操作指南（入门篇）</a></dt>
                </dl>
                <dl class="S-jiao-gong">
                	<dd><a "/liucheng/liu25.Asp"><img src="images1/S-jiao.jpg" alt="淘宝刷钻" /></a></dd>
                    <dt><a ="/jiayuan/view-15-1.html" style="color:#3282c1; text-decoration:underline;">更多..</a></dt>
                </dl>
                <div class="clear"></div>
            </div>
            <div class="S-jiao-part S-jiao-xin">
            	<h3>最新加入</h3>
                <table>
                	<tr>
                    	<td width="248" style="color:#898989;">掌柜的</td>
                        <td style="color:#898989;">加入时间</td>
                    </tr>
                    
                    <tr>
                    	<td> - <a "style="color:#3282c1;">5321</a></td>
                        <td style="color:#8aab01;">最近加入</td>
                    </tr>

                    <tr>
                    	<td> - <a "style="color:#3282c1;">xuegang</a></td>
                        <td style="color:#8aab01;">最近加入</td>
                    </tr>

                    <tr>
                    	<td> - <a "href="style="color:#3282c1;">fly6947</a></td>
                        <td style="color:#8aab01;">最近加入</td>
                    </tr>

                    <tr>
                    	<td> - <a "style="color:#3282c1;">xielong3344</a></td>
                        <td style="color:#8aab01;">最近加入</td>
                    </tr>

                    <tr>
                    	<td> - <a "style="color:#3282c1;">wuzeben1988</a></td>
                        <td style="color:#8aab01;">最近加入</td>
                    </tr>

                    <tr>
                    	<td> - <a " style="color:#3282c1;">hello2</a></td>
                        <td style="color:#8aab01;">最近加入</td>
                    </tr>
 

                
                </table>
            </div>
            <div class="clear"></div>
        </div>
        <!--help st-->
        <div class="help-new">
        	<div class="help-tuiguang">
            	<a href="union/">
            	<p style="font-weight:bold; color:#0086cb; margin-bottom:10px;">推广会员就能赚钱，高额提成，模式多选，多重收益。<p>
				<p>推荐一个会员就可产生丰厚收入，长期收入。<br />
做小七网淘宝刷钻推广，轻松赚大钱！</p>
				</a>
            </div>
            <div class="help-list">
            	<div class="help-list-part helppart01">
                	<h4><a "><b style="color:#0086cb;">常见问题</b></a></h4>
                    <ul>
                    <li>·<a '><span style='margin-left:0px;'>教您如何预防骗子</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>淘宝E客服，子帐号申请规则</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>资金流向图以及平台任务生命周</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>我一天发布多少个任务才安全？</span></a> </li>
                    </ul>
                    <p><a href="/cjwt/"style="color:#0086cb; margin-left:15px;">更多</a></p>
				</div>
                
                <div class="help-list-part helppart02">
                	<h4><a "><b style="color:#0086cb;">热门问题</b></a></h4>
                    <ul>
<li>·<a ' target='_blank'><span style='margin-left:0px;'>如何获得谷粒？</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>谷粒扣除规则</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>关于平台贡献，接手任务后谷粒</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>什么是平台担保金</span></a> </li>
                    </ul>
                    <p><a href="/hotwt/"style="color:#0086cb; margin-left:15px;">更多</a></p>
				</div>
                
                <div class="help-list-part helppart03">
                	<h4><a "><b style="color:#0086cb;">开店经验</b></a></h4>
                    <ul>

<li>·<a " target='_blank'>淘宝新手卖家必备攻略一</a></li>

<li>·<a " target='_blank'>网店经营中的几个小技巧,使销售</a></li>

<li>·<a " target='_blank'>【数据分析】破译消费者做精准营</a></li>

<li>·<a " target='_blank'>在天猫举报什么样的商品？需要哪</a></li>
 
                    </ul>
                    <p><a href="/jiayuan/list-6-1.html"style="color:#0086cb; margin-left:15px;">更多</a></p>
				</div>
                
                <div class="help-list-part helppart04">
                	<h4><a "><b style="color:#0086cb;">卖家帮助</b></a></h4>
                    <ul>
<li>·<a ' target='_blank'><span style='margin-left:0px;'>如何查看托管进度？</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>安全提升手册，每天要发布任务</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>发布商保任务教程</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>新手卖家都需要淡定，只因沉默</span></a> </li>
                    </ul>
                    <p><a href="/fabu/"style="color:#0086cb; margin-left:15px;">更多</a></p>
				</div>
                
                <div class="help-list-part helppart05">
                	<h4><a "><b style="color:#0086cb;">买家帮助</b></a></h4>
                    <ul>
<li>·<a ' target='_blank'><span style='margin-left:0px;'>加入小七网商家保护教程</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>平台任务是不可以使用信用卡付</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>关于不使用绑定买号购买任务商</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>如果淘宝买号被封了要怎么办？</span></a> </li>
                    </ul>
                    <p><a href="/jieshou/"style="color:#0086cb; margin-left:15px;">更多</a></p>
				</div>
                
                <div class="help-list-part helppart06">
                	<h4><a "><b style="color:#0086cb;">家园交流</b></a></h4>
                    <ul>
<li>·<a ' target='_blank'><span style='margin-left:0px;'>快递单号为什么查不到信息呢</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>那位大哥知道，分销商怎么刷信</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>淘宝新手卖家必备攻略一</span></a> </li><li>·<a ' target='_blank'><span style='margin-left:0px;'>网店经营中的几个小技巧,使销</span></a> </li>
                    </ul>
                    <p><a href="/jiayuan/"style="color:#0086cb; margin-left:15px;">更多</a></p>
				</div>
                


            
            </div>
            <div class="clear"></div>
        </div>
        <!--help end-->
    </div>
    <!--main end-->

    <!--footer st-->
    <div class="footer footer-margin">
    	<div style="width:960px; margin:0 auto; text-align:center;">
    	<p><a href="/" style="color:#3282c1;">首页</a> | <a href="/gude18/About.asp" style="color:#3282c1;">关于我们</a> | <a href="Javascript:openNewWindows('newDiv','/service/Gude.html','675','257');" style="color:#3282c1;">联系我们</a> | <a href="/Note.asp" style="color:#3282c1;">小七网规则</a> | <a href="#" style="color:#3282c1;">小七网服务协议</a><font style="color:#b9b9b9;">       Copyright &copy; 2009 - 2013 xiaoqishua.com. ALL Rights Reserved <script type="text/javascript">
var _bdhmProtocol = (("https:" == document.location.protocol) ? " https://" : " http://");
document.write(unescape("%3Cscript src='" + _bdhmProtocol + ");
</script></font>
</p>
<script type="text/javascript">var cnzz_protocol = (("https:" == document.location.protocol) ? " https://" : " http://");document.write(unescape("%3Cspan id='cnzz_stat_icon_5298292'%3E%3C/span%3E%3Cscript src='" + cnzz_protocol + "s11.cnzz.com/stat.php%3Fid%3D5298292%26show%3Dpic' type='text/javascript'%3E%3C/script%3E"));</script>
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
    <!--footer end-->
</body>
</html>
<SCRIPT language=javascript src="js/ServiceQQ.js"></SCRIPT> <DIV id=affiche>
<DIV><SPAN id=closeAffice><I>×</I>关闭</SPAN> 