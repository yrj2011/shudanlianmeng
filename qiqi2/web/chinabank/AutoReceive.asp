<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="MD5.asp"-->
<%
'****************************************	' MD5��ԿҪ�������ύҳ��ͬ����Send.asp��� key = "test" ,�޸�""���� test Ϊ������Կ
											' �������û������MD5��Կ���½����Ϊ���ṩ�̻���̨����ַ��https://merchant3.chinabank.com.cn/
	key = stupshi							' ��½��������ĵ�����������ҵ���B2C�����ڶ������������С�MD5��Կ���á� 
											' ����������һ��16λ���ϵ���Կ����ߣ���Կ���64λ��������16λ�Ѿ��㹻��
'****************************************

v_oid=request("v_oid")'������
v_pmode=request("v_pmode")'�������� �磺��������
v_pstatus=request("v_pstatus")'֧��״̬ �磺20 ֧���ɹ���30 ֧��ʧ��
v_pstring=request("v_pstring")'֧��״̬˵�� �磺֧���ɹ�
v_amount=request("v_amount")'֧�����
v_moneytype=request("v_moneytype")'���� �磺CNY

v_md5str=request("v_md5str")'MD5Ч����

remark1=request("remark1")'��ע1
remark2=request("remark2")'��ע2


if v_md5str = "" then
	
	response.write("error")

	response.end '�жϳ���

end if

text = v_oid&v_pstatus&v_amount&v_moneytype&key 'ƴ�ռ��ܴ�

md5text = Ucase(trim(md5(text))) '����MD5Ч����

if md5text<>v_md5str then '���������߷��͹�����MD5Ч����Աȣ�ȷ�����������߷��͵���Ϣ

	response.write("error") '���߷�������֤ʧ�ܣ�Ҫ���ط�

	response.end '�жϳ���

else

	response.write("ok") '���߷������Ѿ���ȷ�����Լ���֤������ȷ��Ҫ��ֹͣ����

	if v_pstatus = "20" then

		'֧���Ѿ��ɹ�
		'�˴������̻�ϵͳ���߼������������жϽ��ж�֧��״̬(20�ɹ�,30ʧ��)�����¶���״̬�ȵȣ�......

	end if

end if
%>