<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!-- #include file="yeepayCommon.asp" -->
<%
'**************************************************
'* @Description �ױ�֧����Ʒͨ��֧���ӿڷ���  
'* @V3.0
'* @Author rui.xin
'**************************************************

	'	ֻ��֧���ɹ�ʱ�ױ�֧���Ż�֪ͨ�̻�.
	''֧���ɹ��ص������Σ�����֪ͨ������֧����������е�p8_Url �ϣ�������ض���;��������Ե�ͨѶ.
	
	Dim r0_Cmd
	Dim r1_Code
	Dim r2_TrxId
	Dim r3_Amt
	Dim r4_Cur
	Dim r5_Pid
	Dim r6_Order
	Dim r7_Uid
	Dim r8_MP
	Dim r9_BType
	Dim hmac
	
	Dim bRet
	Dim returnMsg
	
	'�������ز���
	Call getCallBackValue(r0_Cmd,r1_Code,r2_TrxId,r3_Amt,r4_Cur,r5_Pid,r6_Order,r7_Uid,r8_MP,r9_BType,hmac)
	Call createLog(logName)
	'�жϷ���ǩ���Ƿ���ȷ��True/False��
	bRet = CheckHmac(r0_Cmd,r1_Code,r2_TrxId,r3_Amt,r4_Cur,r5_Pid,r6_Order,r7_Uid,r8_MP,r9_BType,hmac)
	'���ϴ���ͱ�������Ҫ�޸�.
	
	
	'У������ȷ
	returnMsg	= ""
	If bRet = True Then
	  If(r1_Code="1") Then
	  	
		'��Ҫ�ȽϷ��صĽ�����̼����ݿ��ж����Ľ���Ƿ���ȣ�ֻ����ȵ�����²���Ϊ�ǽ��׳ɹ�.
		'������Ҫ�Է��صĴ������������ƣ����м�¼�������Դ����ڽ��յ�֧�����֪ͨ���ж��Ƿ���й�ҵ���߼�������Ҫ�ظ�����ҵ���߼�������ֹ��ͬһ�������ظ��������������.	  	      	  
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_ding where v_oid='"&r6_Order&"'",conn,1,1
if not (rs.eof) then 
     Response.Write("<script>alert('������!��Ĵ���ѵ���!�벻Ҫ���ظ�����!');window.location.href='../user/index.asp';</script>")
	 response.End()
end if

       Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select * From "&jieducm&"_ding " ,Conn,3,3  
	    rs.addnew
		rs("class")=3
		rs("v_oid")=r6_Order
    	rs("v_pmode")=v_pmode
		rs("v_pstatus")=v_pstatus
		rs("v_pstring")=v_pstring
		rs("v_amount")=r3_Amt
		rs("v_moneytype")=v_moneytype
		rs("remark1")=remark1
		rs("remark2")=remark2	
    	rs.update
		rs.close

total=int(r3_Amt)*0.01
price_kou=int(r3_Amt)-total
	
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
		rs("class")="�ױ���ֵ"
		rs("username")=session("useridname")
    	rs("cunkuan")=r3_Amt
    	rs.update
    	rs.close  

	num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�ױ���ֵ"
		rs("content")=session("useridname")&"ͨ���ױ���ֵ"&r3_Amt&"Ԫ,��"&total&"Ԫ������"
		rs("price")=price_kou
		rs("jiegou")="��ֵ�ɹ�"
		rs("now")=now
    	rs.update
    	rs.close

response.write("��Ĵ��"+price_kou+"�ѵ���")
Response.Write("<script>alert('��Ĵ���ѵ���!');window.location.href='../user/';</script>")
response.End()



			 
	  End IF
	Else
		returnMsg	= returnMsg	&	"������Ϣ���۸�"
	End If


%>

