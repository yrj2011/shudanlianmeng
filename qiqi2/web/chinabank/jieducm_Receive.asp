<!--#include file="MD5.asp"-->
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<%  

'****************************************	' MD5��ԿҪ�������ύҳ��ͬ����Send.asp��� key = "test" ,�޸�""���� test Ϊ������Կ
											' �������û������MD5��Կ���½����Ϊ���ṩ�̻���̨����ַ��https://merchant3.chinabank.com.cn/
	key = stupshi						' ��½��������ĵ�����������ҵ���B2C�����ڶ������������С�MD5��Կ���á�
											' ����������һ��16λ���ϵ���Կ����ߣ���Կ���64λ��������16λ�Ѿ��㹻��
'****************************************

' ȡ�÷��ز���ֵ
	v_oid=request("v_oid")                               ' �̻����͵�v_oid�������
	v_pmode=request("v_pmode")                           ' ֧����ʽ���ַ����� 
	v_pstatus=request("v_pstatus")                       ' ֧��״̬ 20��֧���ɹ���;30��֧��ʧ�ܣ�
	v_pstring=request("v_pstring")                       ' ֧�������Ϣ ֧����ɣ���v_pstatus=20ʱ����ʧ��ԭ�򣨵�v_pstatus=30ʱ����
	v_amount=request("v_amount")                         ' ����ʵ��֧�����
	v_moneytype=request("v_moneytype")                   ' ����ʵ��֧������
	remark1=request("remark1")                           ' ��ע�ֶ�1
	remark2=request("remark2")                           ' ��ע�ֶ�2
	v_md5str=request("v_md5str")                         ' ��������ƴ�յ�Md5У�鴮

text = v_oid&v_pstatus&v_amount&v_moneytype&key
md5text =Ucase(trim(md5(text)))    '�̻�ƴ�յ�Md5У�鴮

if v_pstatus<>20 then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('������!�벻Ҫ���ظ�����!ƽ̨���¼�����Ϊ��');window.location.href='../user/index.asp';</script>")
response.End()
end if
if key="" then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('������!�벻Ҫ���ظ�����!ƽ̨���¼�����Ϊ��');window.location.href='../user/index.asp';</script>")
response.End()
end if
if stupwangyin="0" then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('������!�벻Ҫ���ظ�����!ƽ̨���¼�����Ϊ��');window.location.href='../user/index.asp';</script>")
response.End()
end if
if int(session("v_amount"))<>int(v_amount) then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('������!�벻Ҫ���ظ�����!ƽ̨���¼�����Ϊ��');window.location.href='../user/index.asp';</script>")
response.End()
end if

if md5text<>v_md5str then	
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('������!�벻Ҫ���ظ�����!ƽ̨���¼�����Ϊ��');window.location.href='../user/index.asp';</script>")
response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_ding where v_oid='"&v_oid&"'",conn,1,1
if not (rs.eof) then 
     Response.Write("<script>alert('������!��Ĵ���ѵ���!�벻Ҫ���ظ�����7!');window.location.href='../user/index.asp';</script>")
	 response.End()
end if

       Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select * From "&jieducm&"_ding " ,Conn,3,3  
	    rs.addnew
		rs("class")=1
		rs("v_oid")=v_oid
    	rs("v_pmode")=v_pmode
		rs("v_pstatus")=v_pstatus
		rs("v_pstring")=v_pstring
		rs("v_amount")=v_amount
		rs("v_moneytype")=v_moneytype
		rs("remark1")=remark1
		rs("remark2")=remark2	
    	rs.update
		rs.close

total=int(v_amount)*0.01
price_kou=int(v_amount)-total
	
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select cunkuan from "&jieducm&"_register where username='"&session("useridname")&"'",conn,3,3
if not (rs.eof) then 
  rs("cunkuan")=rs("cunkuan")+price_kou
  rs.update
  rs.close
else
  Response.Write("<script>alert('������!�����ֵ���û���Զ����ˡ�����ϵ����Ա��');window.location.href='../index.asp';</script>")
  response.end()
end if
		
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu" ,Conn,3,3  
	    rs.addnew
		rs("class")="������ֵ"
		rs("username")=session("useridname")
    	rs("cunkuan")=v_amount
    	rs.update
    	rs.close  

	num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="������ֵ"
		rs("content")=session("useridname")&"ͨ��������ֵ"&v_amount&"Ԫ,��"&total&"Ԫ������"
		rs("price")=v_amount
		rs("jiegou")="��ֵ�ɹ�"
		rs("now")=now
    	rs.update
    	rs.close

response.write("��Ĵ��"+request("v_amount")+"�ѵ���")
Response.Write("<script>alert('��Ĵ���ѵ���!');window.location.href='../user/';</script>")
response.End()
%>

