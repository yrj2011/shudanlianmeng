<!--#include file="alipayto/alipay_payto.asp"-->
<%
   shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '客户网站订单号，（现取系统时间，可改成网站自己的变量）
	price1=request.form("v_amounta")
	
	subject			=	"在线充值"		'商品名称
	body			=	"平台在线充值"		'body			商品描述
	out_trade_no    =   dingdan       
	price		    =	price1				'price商品单价			0.01～50000.00
    quantity        =   "1"               '商品数量,如果走购物车默认为1
	discount        =   "0"               '商品折扣
    seller_email    =    seller_email   '卖家的支付宝帐号
	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email)
	response.Redirect(itemUrl)
%>
