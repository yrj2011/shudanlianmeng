<!--#include file="alipayconn.asp"-->
<%
	  
	  	show_url        = stupurl    '商户网站的网址。
	seller_email	= stupzfb	 '请填写签约支付宝账号，
	partner			= stupwushi	 '填写签约支付宝账号对应的partnerID，
	key			    = stupyibai	 '填写签约账号对应的安全校验码

      notify_url			= ""&show_url&"/alipay/Alipay_Notify.asp"	'付完款后服务器通知的页面 要用 http://格式的完整路径
	  return_url			= ""&show_url&"/alipay/return_Alipay_Notify.asp"	'付完款后跳转的页面 要用 http://格式的完整路径
 
	 
'登陆 www.alipay.com 后, 点商家服务,可以看到支付宝安全校验码和合作id,导航栏的下面
	 
%>