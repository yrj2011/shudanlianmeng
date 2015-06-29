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
'*******************************************************************
'获取支付宝POST过来通知消息
For Each varItem in Request.Form
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
 '*************************************************
 
If mysign=request.Form("sign") And ResponseTxt="true" Then 	

   If trade_status = "TRADE_FINISHED" Then
      
	 '付款成功,更新数据库  
	 returnTxt	= "success"
	Else
	returnTxt	= "success"
	End If

	Response.Write returnTxt

Else
response.write "fail"
End If 

'*******************************************************************
 '写文本，方便测试（看网站需求，也可以改成存入数据库）
'TOEXCELLR=TOEXCELLR&md5str&"MD5结果:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
'set fs= createobject("scripting.filesystemobject") 
'set ts=fs.createtextfile(server.MapPath("alipayto/Notify_DATA/"&replace(now(),":","")&".txt"),true)

' ts.writeline(TOEXCELLR)
 'ts.close
' set ts=Nothing
' set fs=Nothing



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
	End Function%>