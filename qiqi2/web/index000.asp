<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ**************************************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
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
    diag.Title = "С���� - ���ֻ�����ѧ����";   
    diag.URL = "/platform/User/TaskNotice/helpcenter.asp";   
    diag.OKEvent = zAlert;//���ȷ������õķ���   
    diag.show();   
   
}   
function zAlert(){   
        Dialog.alert("���ȡ���˳���");   
}   
    </SCRIPT>

<body>
	<!--header st-->
    <div class="header">
    	<div class="quick-channel">
        	<p><script src="/platform/ShowMyInfo.asp" type="text/javascript"></script></p>
        </div>
        <div class="top clearfix">
        	<h1><a href="/">С����</a></h1>
            <p><script src="/ShowPloy.asp" type="text/javascript"></script></p>
        </div>
        <div class="nav clearfix">
        	<ul>
            	<li class="current"><a href="/"><em>��ҳ</em></a></li>
                <li><a href="/taobao/"><em>�Ա�����</em></a></li>
                <li><a href="/pai/"><em>���Ļ���</em></a></li>
                <li><a href="/cang/"><em>�ղ���</em></a></li>
                <li><a href="/user/tuoguan.asp"><em>�й�ר��</em></a></li>
                <li><a href="/user/"><em>��������</em></a></li>
                <li><a href="/chongzhi"/"><em>��ֵ����</em></a></li>
               
                <li><a href="/union/"><em>�Ա��ƹ�</em></a></li>
 <li><a href="/tribe/"><em>����</em></a></li>
                <li><a href="/new.asp"><em>��������</em></a></li>
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
                <p class="S-login-input">��&nbsp;&nbsp;&nbsp;&nbsp;����<input name="name" type="text" /></p>
                <p class="S-login-input">��&nbsp;&nbsp;&nbsp;&nbsp;�룺<input name="pass" type="password"/></p>
                <p class="S-login-btn"><input type="image" src="images1/S-login-btn.jpg"  style="margin-right:10px;"/><input name="" type="checkbox" value="" /> ��ס��  <a href="/platform/login.asp?chuli=zhaomima&urlfan=x" style="color:#3282c1; text-decoration:underline;">�������룿</a></p>
                <p class="S-login-txt">һ����ʵ���Ա���������ƽ̨�ڴ���ļ��룡</p>
                </form>

       	    </div>
            <div class="clear"></div>
        </div>
    </div>
    <!--S-banner end-->
    <!--main st-->
    <div class="main">
    	<div class="S-zhuan"> <a href="/platform/index.html"><img src="images1/S-zhuan.jpg" alt="С����" /></a>
      </div>
        <div class="S-jiao">
        	<div class="S-jiao-part S-jiao-con">
            	<dl class="S-jiao-video">
                	<dd><a ="/gude18/Contract.asp"><img src="images1/S-video.jpg" alt="�Ա�ˢ��ƽ̨" /></a></dd>
                    <dt><a "/gude18/Contract.asp">��Ƶ����ָ�ϣ�����ƪ��</a></dt>
                </dl>
                <dl class="S-jiao-gong">
                	<dd><a "/liucheng/liu25.Asp"><img src="images1/S-jiao.jpg" alt="�Ա�ˢ��" /></a></dd>
                    <dt><a ="/jiayuan/view-15-1.html" style="color:#3282c1; text-decoration:underline;">����..</a></dt>
                </dl>
                <div class="clear"></div>
            </div>
            <div class="S-jiao-part S-jiao-xin">
            	<h3>���¼���</h3>
                <table>
                	<tr>
                    	<td width="248" style="color:#898989;">�ƹ��</td>
                        <td style="color:#898989;">����ʱ��</td>
                    </tr>
                    
                    <tr>
                    	<td> - <a "style="color:#3282c1;">5321</a></td>
                        <td style="color:#8aab01;">�������</td>
                    </tr>

                    <tr>
                    	<td> - <a "style="color:#3282c1;">xuegang</a></td>
                        <td style="color:#8aab01;">�������</td>
                    </tr>

                    <tr>
                    	<td> - <a "href="style="color:#3282c1;">fly6947</a></td>
                        <td style="color:#8aab01;">�������</td>
                    </tr>

                    <tr>
                    	<td> - <a "style="color:#3282c1;">xielong3344</a></td>
                        <td style="color:#8aab01;">�������</td>
                    </tr>

                    <tr>
                    	<td> - <a "style="color:#3282c1;">wuzeben1988</a></td>
                        <td style="color:#8aab01;">�������</td>
                    </tr>

                    <tr>
                    	<td> - <a " style="color:#3282c1;">hello2</a></td>
                        <td style="color:#8aab01;">�������</td>
                    </tr>
 

                
                </table>
            </div>
            <div class="clear"></div>
        </div>
        <!--help st-->
        <div class="help-new">
        	<div class="help-tuiguang">
            	<a href="union/">
            	<p style="font-weight:bold; color:#0086cb; margin-bottom:10px;">�ƹ��Ա����׬Ǯ���߶���ɣ�ģʽ��ѡ���������档<p>
				<p>�Ƽ�һ����Ա�Ϳɲ���������룬�������롣<br />
��С�����Ա�ˢ���ƹ㣬����׬��Ǯ��</p>
				</a>
            </div>
            <div class="help-list">
            	<div class="help-list-part helppart01">
                	<h4><a "><b style="color:#0086cb;">��������</b></a></h4>
                    <ul>
                    <li>��<a '><span style='margin-left:0px;'>�������Ԥ��ƭ��</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>�Ա�E�ͷ������ʺ��������</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>�ʽ�����ͼ�Լ�ƽ̨����������</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>��һ�췢�����ٸ�����Ű�ȫ��</span></a> </li>
                    </ul>
                    <p><a href="/cjwt/"style="color:#0086cb; margin-left:15px;">����</a></p>
				</div>
                
                <div class="help-list-part helppart02">
                	<h4><a "><b style="color:#0086cb;">��������</b></a></h4>
                    <ul>
<li>��<a ' target='_blank'><span style='margin-left:0px;'>��λ�ù�����</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>�����۳�����</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>����ƽ̨���ף�������������</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>ʲô��ƽ̨������</span></a> </li>
                    </ul>
                    <p><a href="/hotwt/"style="color:#0086cb; margin-left:15px;">����</a></p>
				</div>
                
                <div class="help-list-part helppart03">
                	<h4><a "><b style="color:#0086cb;">���꾭��</b></a></h4>
                    <ul>

<li>��<a " target='_blank'>�Ա��������ұر�����һ</a></li>

<li>��<a " target='_blank'>���꾭Ӫ�еļ���С����,ʹ����</a></li>

<li>��<a " target='_blank'>�����ݷ�������������������׼Ӫ</a></li>

<li>��<a " target='_blank'>����è�ٱ�ʲô������Ʒ����Ҫ��</a></li>
 
                    </ul>
                    <p><a href="/jiayuan/list-6-1.html"style="color:#0086cb; margin-left:15px;">����</a></p>
				</div>
                
                <div class="help-list-part helppart04">
                	<h4><a "><b style="color:#0086cb;">���Ұ���</b></a></h4>
                    <ul>
<li>��<a ' target='_blank'><span style='margin-left:0px;'>��β鿴�йܽ��ȣ�</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>��ȫ�����ֲᣬÿ��Ҫ��������</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>�����̱�����̳�</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>�������Ҷ���Ҫ������ֻ���Ĭ</span></a> </li>
                    </ul>
                    <p><a href="/fabu/"style="color:#0086cb; margin-left:15px;">����</a></p>
				</div>
                
                <div class="help-list-part helppart05">
                	<h4><a "><b style="color:#0086cb;">��Ұ���</b></a></h4>
                    <ul>
<li>��<a ' target='_blank'><span style='margin-left:0px;'>����С�����̼ұ����̳�</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>ƽ̨�����ǲ�����ʹ�����ÿ���</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>���ڲ�ʹ�ð���Ź���������</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>����Ա���ű�����Ҫ��ô�죿</span></a> </li>
                    </ul>
                    <p><a href="/jieshou/"style="color:#0086cb; margin-left:15px;">����</a></p>
				</div>
                
                <div class="help-list-part helppart06">
                	<h4><a "><b style="color:#0086cb;">��԰����</b></a></h4>
                    <ul>
<li>��<a ' target='_blank'><span style='margin-left:0px;'>��ݵ���Ϊʲô�鲻����Ϣ��</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>��λ���֪������������ôˢ��</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>�Ա��������ұر�����һ</span></a> </li><li>��<a ' target='_blank'><span style='margin-left:0px;'>���꾭Ӫ�еļ���С����,ʹ��</span></a> </li>
                    </ul>
                    <p><a href="/jiayuan/"style="color:#0086cb; margin-left:15px;">����</a></p>
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
    	<p><a href="/" style="color:#3282c1;">��ҳ</a> | <a href="/gude18/About.asp" style="color:#3282c1;">��������</a> | <a href="Javascript:openNewWindows('newDiv','/service/Gude.html','675','257');" style="color:#3282c1;">��ϵ����</a> | <a href="/Note.asp" style="color:#3282c1;">С��������</a> | <a href="#" style="color:#3282c1;">С��������Э��</a><font style="color:#b9b9b9;">       Copyright &copy; 2009 - 2013 xiaoqishua.com. ALL Rights Reserved <script type="text/javascript">
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

    <!-- �·����������� -->
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
            <a class="signin" href="login.asp">��¼</a>
            <a class="signup" href="register.asp">���ע��</a>
            <span class="flow"></span>
        </div>
        <span id="closeBottomBar" onClick="closeBottomBar()">�ر�</span>
    </div>
    <script>
    var closeBottomBar = function () {
        document.getElementById("bottomBar").style.display="none";
    }
    </script>
    <!-- /�·����������� -->
<script>
$(document).ready(function(){
        //���Ƚ�#back-to-top����
        //$("#bottomBar").hide();
        //����������λ�ô��ھඥ��100��������ʱ����ת���ӳ��֣�������ʧ
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
        //���Ƚ�#back-to-top����
	$('#b-stage').cycle({fx:'scrollLeft',pager:'#b-nav',speed:700 });
})
</script>
    <!--footer end-->
</body>
</html>
<SCRIPT language=javascript src="js/ServiceQQ.js"></SCRIPT> <DIV id=affiche>
<DIV><SPAN id=closeAffice><I>��</I>�ر�</SPAN> 