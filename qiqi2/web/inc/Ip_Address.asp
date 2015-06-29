<!-- #Include File="config.asp" -->
<!-- #Include File="Cls_IpSys.asp" -->
<%


dim ppip,ip_ok,useres_ip,useres_port,useres_address,useres_url
ppip=trim(request.querystring("ip"))
if ip_true(ppip)="no" then
  ip_ok="no"
else
  useres_ip=ip_ip(ppip)
  useres_port=ip_port(ppip)
  useres_address=useres_ip
  useres_address=ip_address(useres_address)
  ip_ok="yes"
End If

call web_clear(0)

useres_url=trim(request.servervariables("http_referer"))
if useres_url="" then
  useres_url="parent.location='/'"
else
  useres_url="history.back(1)"
End If
if useres_port<>"yes" and useres_port<>"no" and isnumeric(useres_port) then
  useres_port="端 口 号：" & useres_port & "\n"
else
  useres_port=""
End If

if ip_ok="yes" then
  Response.Write "<script language=javascript>" & _
				 vbcrlf & "alert(""您要查询的 IP 信息如下：\n\nIP 地址：" & useres_ip & "\n" & useres_port & "来源地区：" & useres_address & """);" & _
				 vbcrlf & useres_url & _
				 vbcrlf & "</script>"
else
  Response.Write "<script language=javascript>" & _
				 vbcrlf & "alert(""您的可能进行了非法操作！\n\n系统将自动返回论坛首页。"");" & _
				 vbcrlf & "parent.location='main.asp'" & _
				 vbcrlf & "</script>"
End If

%>