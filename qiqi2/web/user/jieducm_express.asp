<%
'请保留此声明信息,这段声明不并会影响你的速度!
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn http://www.jieducm.com
'软件销售平台：http://www.258shua.com
'演示平台地址：http://www.258shua.net/
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'注: 请不要擅自更改此标识,更不要私下贩卖本程序,否则视为自动放弃售后服务
'*****************************************************************

From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "请求超时,请关闭本页面并重新登陆即可!"
response.end
end If
Response.Charset="gb2312" 
Server.ScriptTimeout = 999999999
rq=request.form("jieducm_rq")
dq=request.form("jieducm_dq")
vip=request.form("jieducm_vip")
jieducm_type=request.form("type")
if vip<>1 then
userName=Request.Cookies("jieducm_name")("visiter") 
userName2=Request.Cookies("jieducm_name2")("visiter") 
Response.Cookies("jieducm_name")("visiter")=userName+1 
Response.Cookies("jieducm_name2")("visiter")=userName2+1 
Response.Cookies("jieducm_name").Expires= (now()+1) 
if (userName)>=(rq) and jieducm_type=1 then
Response.Write("普通会员每天可以使用"&rq&"次随机生成")
response.End
end if
if (userName2)>=(dq) and jieducm_type=2 then
Response.Write("普通会员每天可以使用"&dq&"次自定义生成")
response.End
end if
end if
RefreshIntervalTime = 10 '防止刷新的时间秒数，0表示不防止
If Not IsEmpty(Session("QYCvisit")) and isnumeric(Session("QYCvisit")) and int(RefreshIntervalTime) > 0 Then
if (timer()-int(Session("QYCvisit")))*1000 < RefreshIntervalTime * 1000 then
Response.write ("<meta http-equiv=""refresh"" content="""& RefreshIntervalTime &""" />")
Response.write ("<b>刷新时间间隔为"& RefreshIntervalTime &"秒，请稍候重试！</b>")
Session("QYCvisit") = timer()
Response.end
end if
End If
Session("QYCvisit") = timer()

Dim Retrieval,NumArr,AppKey,SendURL,ResponseTxt,url,nu
AppKey = "39232c200e901622" 
com = Request.form("com")'公司
jieducm_kdnu=request.form("nu")

if com="yuantong" then
jieducm_ji=0
jieducm_lennums=10
elseif com="yunda" then
jieducm_ji=1
jieducm_lennums=13
elseif com="zhongtong" then
jieducm_ji=2
jieducm_lennums=12
elseif com="huitongkuaidi" then
jieducm_ji=3
jieducm_lennums=12
elseif com="tiantian" then
jieducm_ji=4
jieducm_lennums=14
elseif com="zhaijisong" then
jieducm_ji=5
jieducm_lennums=10
end if
'各快递运单号生成模型 (只需要输入单号前几位即可，可自行增加，多个用短号分隔)
jieducm_nu_NumArr = Array("17323,17713,17527,24913,24923,24951,25073,604984,61474,61489,61531,61546,61638,61691,61743,61801,61817,618170,61846,61894,7007,70079",_
"190013722,190013723,190013724,190013725,12002428,1200414888,12004369,12004339,1200437,15000283,16001051,19001166,19000936,19000946,19000956,19000966,190009605,19000971,80000148,8000014",_
"7170087,7170088,7170081,7170082,7170089,618181,618183,628015,628018,628020,628021,7010007,717000,7170007,718002,728002,757100,761003,762003,762004",_
"21005092,21005093,21005097,21005098,00001527,21005096,02000,21005094,21005099,280007",_
"0000112024,000001479,00001120245,00001120242,000001691,0000036212,0000037012,000008416,000001997,00001120247",_
"2990982,299438,540611,544039,576484,577609,647,6412005,642065,64297,6429765,642065,644087,644849,647975,671476,722274,7827788,782889,821115,9101187,911244,911930,912372,931646")

jieducm_num_rnd=Array("","9","99","999","9999","99999","999999","9999999","99999999","999999999","9999999999","99999999999","999999999999")

if jieducm_type="1" then

jieducm_nu = split(jieducm_nu_NumArr(jieducm_ji),",")
jieducm_dsk=Ubound(jieducm_nu)

for i=1 to 5
randomize
jieducm_mid=int(jieducm_dsk*rnd)
jieducm_nu_shen=int(jieducm_lennums)-int(len(jieducm_nu(jieducm_mid)))
jieducm_d=jieducm_num_rnd(jieducm_nu_shen)
jieducm_Cid=int(jieducm_d*rnd)
nu=jieducm_nu(jieducm_mid)&jieducm_Cid
SendURL ="http://api.kuaidi100.com/api?id="&AppKey&"&com="&com&"&nu="&nu&"&show=2&muti=1&order=asc"
ResponseTxt=GetHTTPPage(SendURL) '//获取源代码的函数发送数据
response.Write("<font color=red><strong>运单号:"&nu&"</strong></font>")'输入查询结果
Response.Write ResponseTxt
next
elseif jieducm_type="2" then
jieducm_kdadd=jieducm_lennums-len(jieducm_kdnu)
jieducm_d=jieducm_num_rnd(jieducm_kdadd)
for jieducm_i=1 to 5
randomize
jieducm_Cid=int(jieducm_d*rnd)
jieducm_kdnu2=jieducm_kdnu&jieducm_Cid
SendURL ="http://api.kuaidi100.com/api?id="&AppKey&"&com="&com&"&nu="&jieducm_kdnu2&"&show=2&muti=1&order=asc"
ResponseTxt=GetHTTPPage(SendURL) '//获取源代码的函数发送数据
response.Write("<font color=red><strong>运单号:"&jieducm_kdnu2&"</strong></font>")'输入查询结果
Response.Write ResponseTxt
next
elseif jieducm_type="3" then
SendURL ="http://api.kuaidi100.com/api?id="&AppKey&"&com="&com&"&nu="&jieducm_kdnu&"&show=2&muti=1&order=asc"
ResponseTxt=GetHTTPPage(SendURL) '//获取源代码的函数发送数据
response.Write("<font color=red><strong>运单号:"&jieducm_kdnu&"</strong></font>")'输入查询结果
Response.Write ResponseTxt
end if
'调用发送数据组件
Function GetHTTPPage(URL) 
    Dim objXML 
    Set objXML=CreateObject("MSXML2.SERVERXMLHTTP.3.0")  '调用XMLHTTP组件，测试空间是否支持XMLHTTP，如果服务不支持，请测试下面两个。
	'Set objXML=Server.CreateObject("Microsoft.XMLHTTP") 
	'Set objXML=Server.CreateObject("MSXML2.XMLHTTP.4.0") 
	objXML.SetTimeouts 5000, 5000, 30000, 10000' 
    objXML.Open "GET",URL,False 'false表示以同步的方式获取网页代码，了解什么是同步？什么是异步？
    objXML.Send() '发送
	If objXML.Readystate<>4 Then
		Exit Function 
	End If
	GetHTTPPage=BytesToBstr(objXML.ResponseBody)'返回信息，同时用函数定义编码，如果您需要转码请选择 
    Set objXML=Nothing'关闭 
	If Err.number<>0 Then 
		Response.Write "<p align='center'><font color='red'><b>服务器获取文件内容出错，请稍后再试！</b></font></p>" 
		Err.Clear
	End If
End Function

'页面编码转换
Function BytesToBstr(body) 
    Dim objstream 
    Set objstream = Server.CreateObject("Adodb.Stream") '//调用adodb.stream组件
    objstream.Type = 1 
    objstream.Mode =3 
    objstream.Open 
    objstream.Write body 
    objstream.Position = 0 
    objstream.Type = 2 
    objstream.Charset = "UTF-8" '转换原来默认的编码转换成GB2312编码，否则直接用XMLHTTP调用有中文字符的网页得到的将是乱码 
    BytesToBstr = objstream.ReadText 
    objstream.Close 
    Set objstream = Nothing 
End Function
%>