<!--#include file="alipayto/alipay_payto.asp"-->
<%
   shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	price1=request.form("v_amounta")
	
	subject			=	"���߳�ֵ"		'��Ʒ����
	body			=	"ƽ̨���߳�ֵ"		'body			��Ʒ����
	out_trade_no    =   dingdan       
	price		    =	price1				'price��Ʒ����			0.01��50000.00
    quantity        =   "1"               '��Ʒ����,����߹��ﳵĬ��Ϊ1
	discount        =   "0"               '��Ʒ�ۿ�
    seller_email    =    seller_email   '���ҵ�֧�����ʺ�
	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email)
	response.Redirect(itemUrl)
%>
