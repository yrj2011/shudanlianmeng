<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
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
if session("useridname")="" then
Response.Write("<script>alert('����û�е�¼��������,���ȵ�¼!');window.location.href='../login.asp';</script>")
response.End()
end if
if request.form("name")="" then
response.Write "<script LANGUAGE='javascript'>alert('����д��������');history.go(-1);</script>"
response.End
end if
if request.form("content")="" then
response.Write "<script LANGUAGE='javascript'>alert('����д���������뽨��');history.go(-1);</script>"
response.End
end if

if request.form("temp")="" then
response.write "<script language='javascript'>" & VbCRlf
response.write "alert('�Ƿ�����!');" & VbCrlf
response.write "history.go(-1);" & vbCrlf
response.write "</script>" & VbCRLF
'���ڱ����ظ��ύ(��־Ϊsession("antry")Ϊ��)����
else
iname=HtmlEncode(request.Form("name"))
iqq=HtmlEncode(request.Form("qq"))
iemail=HtmlEncode(request.Form("email"))
iurl=HtmlEncode(request.form("url"))
isex="boy"
icontent=HtmlEncode(request.form("content"))

set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_guestbook where (id is null)",conn,1,3
rs.addnew
rs("name")=iname
rs("sex")=isex
rs("content")=icontent
rs("time")=now()
rs.update
rs.close
set rs=nothing
session("antry")="" '�ύ�ɹ������session("antry")���Է��ظ��ύ����
end if
%>
<meta http-equiv="refresh" content="1;URL=viewreturn.asp">
<title>���Գɹ�</title>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="62%" border="0" align="center">
  <tr>
    <td align="center"><p> <img src="images/loading.gif" width="94" height="27"><br>
        <br>
        <font size="2">���Գɹ���1���Ӻ󷵻أ�</font></p>
      </td>
  </tr>
</table>