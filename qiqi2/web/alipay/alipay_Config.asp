<!--#include file="alipayconn.asp"-->
<%
	  
	  	show_url        = stupurl    '�̻���վ����ַ��
	seller_email	= stupzfb	 '����дǩԼ֧�����˺ţ�
	partner			= stupwushi	 '��дǩԼ֧�����˺Ŷ�Ӧ��partnerID��
	key			    = stupyibai	 '��дǩԼ�˺Ŷ�Ӧ�İ�ȫУ����

      notify_url			= ""&show_url&"/alipay/Alipay_Notify.asp"	'�����������֪ͨ��ҳ�� Ҫ�� http://��ʽ������·��
	  return_url			= ""&show_url&"/alipay/return_Alipay_Notify.asp"	'��������ת��ҳ�� Ҫ�� http://��ʽ������·��
 
	 
'��½ www.alipay.com ��, ���̼ҷ���,���Կ���֧������ȫУ����ͺ���id,������������
	 
%>