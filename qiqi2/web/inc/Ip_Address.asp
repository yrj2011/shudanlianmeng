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
  useres_port="�� �� �ţ�" & useres_port & "\n"
else
  useres_port=""
End If

if ip_ok="yes" then
  Response.Write "<script language=javascript>" & _
				 vbcrlf & "alert(""��Ҫ��ѯ�� IP ��Ϣ���£�\n\nIP ��ַ��" & useres_ip & "\n" & useres_port & "��Դ������" & useres_address & """);" & _
				 vbcrlf & useres_url & _
				 vbcrlf & "</script>"
else
  Response.Write "<script language=javascript>" & _
				 vbcrlf & "alert(""���Ŀ��ܽ����˷Ƿ�������\n\nϵͳ���Զ�������̳��ҳ��"");" & _
				 vbcrlf & "parent.location='main.asp'" & _
				 vbcrlf & "</script>"
End If

%>