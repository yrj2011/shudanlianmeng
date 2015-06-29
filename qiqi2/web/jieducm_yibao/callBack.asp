<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!-- #include file="yeepayCommon.asp" -->
<%
'**************************************************
'* @Description 易宝支付产品通用支付接口范例  
'* @V3.0
'* @Author rui.xin
'**************************************************

	'	只有支付成功时易宝支付才会通知商户.
	''支付成功回调有两次，都会通知到在线支付请求参数中的p8_Url 上：浏览器重定向;服务器点对点通讯.
	
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
	
	'解析返回参数
	Call getCallBackValue(r0_Cmd,r1_Code,r2_TrxId,r3_Amt,r4_Cur,r5_Pid,r6_Order,r7_Uid,r8_MP,r9_BType,hmac)
	Call createLog(logName)
	'判断返回签名是否正确（True/False）
	bRet = CheckHmac(r0_Cmd,r1_Code,r2_TrxId,r3_Amt,r4_Cur,r5_Pid,r6_Order,r7_Uid,r8_MP,r9_BType,hmac)
	'以上代码和变量不需要修改.
	
	
	'校验码正确
	returnMsg	= ""
	If bRet = True Then
	  If(r1_Code="1") Then
	  	
		'需要比较返回的金额与商家数据库中订单的金额是否相等，只有相等的情况下才认为是交易成功.
		'并且需要对返回的处理进行事务控制，进行记录的排它性处理，在接收到支付结果通知后，判断是否进行过业务逻辑处理，不要重复进行业务逻辑处理，防止对同一条交易重复发货的情况发生.	  	      	  
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_ding where v_oid='"&r6_Order&"'",conn,1,1
if not (rs.eof) then 
     Response.Write("<script>alert('出错了!你的存款已到账!请不要再重复操作!');window.location.href='../user/index.asp';</script>")
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
  Response.Write("<script>alert('出错了!如果充值金额没有自动到账。请联系管理员！');window.location.href='../index.asp';</script>")
  response.end()
end if

        Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu" ,Conn,3,3  
	    rs.addnew
		rs("class")="易宝充值"
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
		rs("class")="易宝充值"
		rs("content")=session("useridname")&"通过易宝充值"&r3_Amt&"元,扣"&total&"元手续费"
		rs("price")=price_kou
		rs("jiegou")="充值成功"
		rs("now")=now
    	rs.update
    	rs.close

response.write("你的存款"+price_kou+"已到账")
Response.Write("<script>alert('你的存款已到账!');window.location.href='../user/';</script>")
response.End()



			 
	  End IF
	Else
		returnMsg	= returnMsg	&	"交易信息被篡改"
	End If


%>

