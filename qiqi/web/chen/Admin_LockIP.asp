<%@language=vbscript codepage=936 %>
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
option explicit
response.buffer=true	
Const PurviewLevel=1    '����Ȩ��
Const CheckChannelID=0    '����Ƶ����0Ϊ���������Ƶ��
Const PurviewLevel_Others="ip"
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->

<!--#include file="inc/md5.asp"-->
<script language="JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
function CheckAdd()
{
  if(document.AddIP.ip1.value=="")
    {
      alert("��ʼIP����Ϊ�գ�");
	  document.AddIP.ip1.focus();
      return false;
    }
    
  if(document.AddIP.ip2.value=="")
    {
      alert("��βIP����Ϊ�գ�");
	  document.AddIP.ip2.focus();
      return false;
    }
}
</script>
<%
dim rs, sql,i
dim Action,FoundErr,ErrMsg
Action=Trim(request("Action"))
%>
<html>
<head>
<title>IP�����޶�����ҳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center" class="title"><strong>IP�����޶�����</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="Admin_LockIP.asp?Action=Add">IP�����޶����</a> | <a href="Admin_LockIP.asp">IP�����޶�����</a> 
	</td>
  </tr>
</table>
<%
dim userip,ips,GetIp1,GetIp2
if request("userip")<>"" then
userip=request("userip")
ips=Split(userIP,".")
GetIp1=ips(0)&"."&ips(1)&"."&ips(2)&".1"
GetIp2=ips(0)&"."&ips(1)&"."&ips(2)&".255"
else 
userip=""
GetIp1=""
GetIp2=""
end if

if Action="Add" then
	call Add()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Del" then
	call DelIP()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
''call CloseConn() 'shiyu

sub main()
%>
  <table width="100%" border="0" cellspacing="1" cellpadding="0" class="border">
<FORM METHOD=POST ACTION="?action=Del" onSubmit="return confirm('ȷ��Ҫɾ��ѡ�е�IP��Χ��');">
    <tr class="title" align="center"> 
      <td height="22" align="center" colspan="4" class="title"><b>IP�������ƣ�������</b></td>
    </tr>
    <tr valign="top" class="tdbg" align="center"> 
      <td><b>ID��</b></td>
      <td><b>��ʼIP</b></td>
      <td><b>��βIP</b></td>
      <td><b>����</b></td>
    </tr>
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	sql="select id,sip1,sip2 from LockIP order by id desc"
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
<tr class="tdbg"> 
<td width="100%" colspan=4 align="center">��û���κ�IP�������ݡ�</td>
</tr>
<%
	else
		rs.PageSize = 20
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = 20)
%>
<tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
<td width="15%" align="center" height="20" ><%=rs("id")%></td>
<td width="30%" align="center" ><%=rs("sip1")%></td>
<td width="30%" align="center" ><%=rs("sip2")%></td>
<td width="25%" align="center" ><input type=checkbox name="delid" value="<%=cstr(rs("ID"))%>" style="border: 0px;background-color: #eeeeee;"></td>
</tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
%>
<tr align="center" class="tdbg"> 
<td colspan=3  align=center>��ҳ��<font color="#FF6600">
<%Pcount=rs.PageCount
	if currentpage > 4 then
	response.write "<a href=""?page=1"">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a>"
	end if
%></font>
</td>
<td align="left" >
 <input type=submit name=submit value=" ɾ&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;"> 
 <input type=checkbox value="on" name="chkall" onClick="CheckAll(this.form)" style="border: 0px;background-color: #eeeeee;">ȫѡ&nbsp;&nbsp;&nbsp;</td>
</tr>
<%
	end if
	rs.close
	set rs=nothing
%>
</FORM>
</table>
<%
end sub

sub Add()
%>
  <table width="100%" border="0" cellspacing="1" cellpadding="0" class="border">
<form action="admin_LockIP.asp?action=SaveAdd" method="post" name="AddIP" onSubmit="javascript:return CheckAdd();">
  <tr class="title" align="center"> 
    <td width="100%" colspan=2 height="22" class="title"><b>IP�������ƣ������</b></td>
  </tr>
  <tr class="tdbg"> 
    <td width="35%" align="center" >��ʼI&nbsp;P</td>
    <td width="65%" > 
      <input type="text" name="ip1" size="30" value="<%=GetIp1%>">
      &nbsp;��202.152.12.1</td>
  </tr>
  <tr class="tdbg"> 
    <td width="35%" align="center" >��βI&nbsp;P</td>
    <td width="65%" > 
      <input type="text" name="ip2" size="30" value="<%=GetIp2%>">
      &nbsp;��202.152.12.255</td>
  </tr>
  <tr align="center" class="tdbg"> 
    <td colspan="2" > 
      <input type="submit" name="Submit" value=" ��&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;">
    </td>
  </tr>
  </form>
</table>
<%
end sub

sub SaveAdd()
dim sip,str1,str2,str3,str4,num_1,num_2
sip=cstr(request.form("ip1"))
'dot=instr(ip,".")-1
'response.write dot
str1=left(sip,cint(instr(sip,".")-1))
sip=mid(sip,cint(instr(sip,"."))+1)
str2=left(sip,cint(instr(sip,"."))-1)
sip=mid(sip,cint(instr(sip,"."))+1)
str3=left(sip,cint(instr(sip,"."))-1)
str4=mid(sip,cint(instr(sip,"."))+1)
num_1=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1

sip=cstr(request.form("ip2"))
str1=left(sip,instr(sip,".")-1)
sip=mid(sip,instr(sip,".")+1)
str2=left(sip,instr(sip,".")-1)
sip=mid(sip,instr(sip,".")+1)
str3=left(sip,instr(sip,".")-1)
str4=mid(sip,instr(sip,".")+1)
num_2=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
'response.write num_1 &","& num_2
'response.end

set rs = server.CreateObject ("adodb.recordset")
sql="select * from LockIP"
rs.open sql,conn,1,3
rs.addnew
rs("ip1")=num_1
rs("ip2")=num_2
rs("sip1")=request.form("ip1")
rs("sip2")=request.form("ip2")
rs.update
rs.close
set rs=nothing
%>
<table width="100%" border="0" cellspacing="1" cellpadding="0" class="border">
<tr class="tdbg">
<td width="100%" colspan=2 >��ӳɹ���</td>
</tr>
</table>
<%
end sub

sub DelIP()
	conn.execute("delete from lockip where id in ("&request.form("delid")&")")
	call main()
end sub
%>
<!--#include file="Admin_fooder.asp"-->
