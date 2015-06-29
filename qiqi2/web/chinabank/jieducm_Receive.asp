<!--#include file="MD5.asp"-->
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<%  

'****************************************	' MD5密钥要跟订单提交页相同，如Send.asp里的 key = "test" ,修改""号内 test 为您的密钥
											' 如果您还没有设置MD5密钥请登陆我们为您提供商户后台，地址：https://merchant3.chinabank.com.cn/
	key = stupshi						' 登陆后在上面的导航栏里可能找到“B2C”，在二级导航栏里有“MD5密钥设置”
											' 建议您设置一个16位以上的密钥或更高，密钥最多64位，但设置16位已经足够了
'****************************************

' 取得返回参数值
	v_oid=request("v_oid")                               ' 商户发送的v_oid定单编号
	v_pmode=request("v_pmode")                           ' 支付方式（字符串） 
	v_pstatus=request("v_pstatus")                       ' 支付状态 20（支付成功）;30（支付失败）
	v_pstring=request("v_pstring")                       ' 支付结果信息 支付完成（当v_pstatus=20时）；失败原因（当v_pstatus=30时）；
	v_amount=request("v_amount")                         ' 订单实际支付金额
	v_moneytype=request("v_moneytype")                   ' 订单实际支付币种
	remark1=request("remark1")                           ' 备注字段1
	remark2=request("remark2")                           ' 备注字段2
	v_md5str=request("v_md5str")                         ' 网银在线拼凑的Md5校验串

text = v_oid&v_pstatus&v_amount&v_moneytype&key
md5text =Ucase(trim(md5(text)))    '商户拼凑的Md5校验串

if v_pstatus<>20 then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('出错了!请不要再重复操作!平台会记录你的行为！');window.location.href='../user/index.asp';</script>")
response.End()
end if
if key="" then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('出错了!请不要再重复操作!平台会记录你的行为！');window.location.href='../user/index.asp';</script>")
response.End()
end if
if stupwangyin="0" then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('出错了!请不要再重复操作!平台会记录你的行为！');window.location.href='../user/index.asp';</script>")
response.End()
end if
if int(session("v_amount"))<>int(v_amount) then
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('出错了!请不要再重复操作!平台会记录你的行为！');window.location.href='../user/index.asp';</script>")
response.End()
end if

if md5text<>v_md5str then	
sqlr="update "&jieducm&"_register set  level1='0' where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('出错了!请不要再重复操作!平台会记录你的行为！');window.location.href='../user/index.asp';</script>")
response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_ding where v_oid='"&v_oid&"'",conn,1,1
if not (rs.eof) then 
     Response.Write("<script>alert('出错了!你的存款已到账!请不要再重复操作7!');window.location.href='../user/index.asp';</script>")
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
  Response.Write("<script>alert('出错了!如果充值金额没有自动到账。请联系管理员！');window.location.href='../index.asp';</script>")
  response.end()
end if
		
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu" ,Conn,3,3  
	    rs.addnew
		rs("class")="网银充值"
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
		rs("class")="网银充值"
		rs("content")=session("useridname")&"通过网银充值"&v_amount&"元,扣"&total&"元手续费"
		rs("price")=v_amount
		rs("jiegou")="充值成功"
		rs("now")=now
    	rs.update
    	rs.close

response.write("你的存款"+request("v_amount")+"已到账")
Response.Write("<script>alert('你的存款已到账!');window.location.href='../user/';</script>")
response.End()
%>

