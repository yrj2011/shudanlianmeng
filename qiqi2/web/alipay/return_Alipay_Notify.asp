<!--#include file="alipayto/Alipay_md5.asp"-->
<!--#include file="alipayconn.asp"-->
<%


 partner			= stupwushi	 '填写签约支付宝账号对应的partnerID，
	key			    = stupyibai	 '填写签约账号对应的安全校验码


	out_trade_no		= DelStr(Request("out_trade_no")) '获取定单号
    total_fee		    = DelStr(Request("total_fee")) '获取支付的总价格
    receive_name    =DelStr(Request("receive_name"))   '获取收货人姓名
	receive_address =DelStr(Request("receive_address")) '获取收货人地址
	receive_zip     =DelStr(Request("receive_zip"))   '获取收货人邮编
	receive_phone   =DelStr(Request("receive_phone")) '获取收货人电话
	receive_mobile  =DelStr(Request("receive_mobile")) '获取收货人手机

'******************************************判断消息是不是支付宝发出
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*****************************************
'获取支付宝GET过来通知消息,判断消息是不是被修改过
For Each varItem in Request.QueryString
mystr=varItem&"="&Request(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
mystr=Left(mystr,Len(mystr)-1)
End If 

mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'对参数排序
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
'构造md5摘要字符串
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
     Response.Write("<script>alert('出错了!你的存款已到账!请不要再重复操作!');window.location.href='../user/';</script>")
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
  Response.Write("<script>alert('出错了!如果充值金额没有自动到账。请联系管理员！');window.location.href='../index.asp';</script>")
  response.end()
end if

 

	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="支付宝充值"
		rs("content")=session("useridname")&"通过支付宝充值"&total_fee&"元,扣"&total&"元手续费"
		rs("price")=price
		rs("jiegou")="充值成功"
		rs("now")=now
    	rs.update
    	rs.close
		
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu" ,Conn,3,3  
	    rs.addnew
		rs("class")="支付宝充值"
		rs("username")=session("useridname")
		rs("cunkuan")=price
		rs("now")=now()
    	rs.update
    	rs.close 
		
  Response.Write("<script>alert('你的存款已到账!');window.location.href='../user/';</script>")
  response.End()
Else
response.write "跳转失败"          '这里可以指定你需要显示的内容
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
		DelStr	= Replace(DelStr,"　","")
		DelStr	= Replace(DelStr,"%20","")
		DelStr	= Replace(DelStr,"--","")
		DelStr	= Replace(DelStr,"==","")
		DelStr	= Replace(DelStr,"<","")
		DelStr	= Replace(DelStr,">","")
		DelStr	= Replace(DelStr,"%","")
	End Function
%>