<!--#include file="alipayto/Alipay_md5.asp"-->
<!--#include file="alipayconn.asp"-->
<%


 partner			= stupwushi	 '��дǩԼ֧�����˺Ŷ�Ӧ��partnerID��
	key			    = stupyibai	 '��дǩԼ�˺Ŷ�Ӧ�İ�ȫУ����


	out_trade_no		= DelStr(Request("out_trade_no")) '��ȡ������
    total_fee		    = DelStr(Request("total_fee")) '��ȡ֧�����ܼ۸�
    receive_name    =DelStr(Request("receive_name"))   '��ȡ�ջ�������
	receive_address =DelStr(Request("receive_address")) '��ȡ�ջ��˵�ַ
	receive_zip     =DelStr(Request("receive_zip"))   '��ȡ�ջ����ʱ�
	receive_phone   =DelStr(Request("receive_phone")) '��ȡ�ջ��˵绰
	receive_mobile  =DelStr(Request("receive_mobile")) '��ȡ�ջ����ֻ�

'******************************************�ж���Ϣ�ǲ���֧��������
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*****************************************
'��ȡ֧����GET����֪ͨ��Ϣ,�ж���Ϣ�ǲ��Ǳ��޸Ĺ�
For Each varItem in Request.QueryString
mystr=varItem&"="&Request(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
mystr=Left(mystr,Len(mystr)-1)
End If 

mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'�Բ�������
For i = Count TO 0 Step -1
minmax = mystr( 0 )
minmaxSlot = 0
For j = 1 To i
mark = (mystr( j ) > minmax)
If mark Then 
minmax = mystr( j )
minmaxSlot = j
End If 
Next
    
If minmaxSlot <> i Then 
temp = mystr( minmaxSlot )
mystr( minmaxSlot ) = mystr( i )
mystr( i ) = temp
End If
Next
'����md5ժҪ�ַ���
 For j = 0 To Count Step 1
 value = SPLIT(mystr( j ), "=")
 If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
 If j=Count Then
 md5str= md5str&mystr( j )
 Else 
 md5str= md5str&mystr( j )&"&"
 End If 
 End If 
 Next
md5str=md5str&key
 mysign=md5(md5str)
'********************************************************

If mysign=Request("sign") and ResponseTxt="true"   Then 	

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_ding where v_oid='"&out_trade_no&"'",conn,3,3
if not (rs.eof) then 
     Response.Write("<script>alert('������!��Ĵ���ѵ���!�벻Ҫ���ظ�����!');window.location.href='../user/';</script>")
	 response.End()
else
rs.addnew
rs("class")=2
rs("v_oid")=out_trade_no
rs("v_pstatus")=trade_status
rs("v_amount")=total_fee		
rs.update
rs.close
end if

total=int(total_fee)*0.01
price=int(total_fee)-total

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select cunkuan from "&jieducm&"_register where username='"&session("useridname")&"'",conn,3,3
if not (rs.eof) then 
  rs("cunkuan")=rs("cunkuan")+price
  rs.update
  rs.close
else
  Response.Write("<script>alert('������!�����ֵ���û���Զ����ˡ�����ϵ����Ա��');window.location.href='../index.asp';</script>")
  response.end()
end if

 

	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="֧������ֵ"
		rs("content")=session("useridname")&"ͨ��֧������ֵ"&total_fee&"Ԫ,��"&total&"Ԫ������"
		rs("price")=price
		rs("jiegou")="��ֵ�ɹ�"
		rs("now")=now
    	rs.update
    	rs.close
		
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu" ,Conn,3,3  
	    rs.addnew
		rs("class")="֧������ֵ"
		rs("username")=session("useridname")
		rs("cunkuan")=price
		rs("now")=now()
    	rs.update
    	rs.close 
		
  Response.Write("<script>alert('��Ĵ���ѵ���!');window.location.href='../user/';</script>")
  response.End()
Else
response.write "��תʧ��"          '�������ָ������Ҫ��ʾ������
response.End()
End If 


	
	

Function DelStr(Str)
		If IsNull(Str) Or IsEmpty(Str) Then
			Str	= ""
		End If
		DelStr	= Replace(Str,";","")
		DelStr	= Replace(DelStr,"'","")
		DelStr	= Replace(DelStr,"&","")
		DelStr	= Replace(DelStr," ","")
		DelStr	= Replace(DelStr,"��","")
		DelStr	= Replace(DelStr,"%20","")
		DelStr	= Replace(DelStr,"--","")
		DelStr	= Replace(DelStr,"==","")
		DelStr	= Replace(DelStr,"<","")
		DelStr	= Replace(DelStr,">","")
		DelStr	= Replace(DelStr,"%","")
	End Function
%>