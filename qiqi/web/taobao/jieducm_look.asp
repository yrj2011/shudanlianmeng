<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<SCRIPT src="../js/jieducm_zDialog.js" type=text/javascript></SCRIPT>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>

<%
pid=Replace(Trim(request("taobaoid")),"'","''")
Sql = "select * from "&jieducm&"_zhongxin where  pid='"&pid&"' and IfComeSet>0"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
IfComeSet=rs("IfComeSet")
ComeKey	=rs("ComeKey")
ComeNote=rs("ComeNote")
prourl=rs("prourl")
pidy=mid(prourl,instr(1,prourl,"id=")+3,len(prourl)-instr(1,prourl,"id=")-2)
Shopkeeper=rs("Shopkeeper")
else
 Response.Write("<script>alert('�������뷵��!');history.back();</script>")
 response.End()
end if
 
Sql= "select * from "&jieducm&"_jieshou where pid='"&pid&"'  and username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then			
 Response.Write("<script>alert('�Բ���,��������!');history.back();</script>")
 response.End()
end if

if request.Form("action")="jieducm_ok" then
prourl_jieducm=Replace(Trim(Request.Form("prourl_jieducm")),"'","''")
if prourl_jieducm="" then
 Response.Write("<script>alert('������Ҫ��֤����Ʒ����id!');history.back();</script>")
 response.End()
end if
 
'if prourl_jieducm<>prourl then

if prourl_jieducm<>pidy then
 Response.Write("<script>alert('" & prourl_jieducm & "��֤ʧ�ܣ�');history.back();</script>")
 response.End()
else
 Response.Write("<script>alert('��֤�ɹ�!����Ʒ��ַ�Ƿ�������������Ʒ��ַ������Ĺ���');history.back();</script>")
 response.End()
end if
end if
jieducm_taobao="http://www.taobao.com"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE>������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
 <LINK href="../css/top_bottom.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.STYLE1 {font-size: 14px}
.STYLE2 {font-size: 14px}
-->
</style>

<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
	color: #333;
}
body {
	margin-left: 20px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
}
-->
</style>
</HEAD>
<BODY>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 
      borderColor=#dddddd cellPadding=3 width=100% align=center>
  <tr height="30">
    <td colspan="3" bgcolor="#CCCCCC" height="30"><div align="center" class="STYLE1">������������Ʒ��·���񣬽��շ���Ҫ������е���·������ܲ鿴����Ʒ����ַ</div></td>
  </tr>
  <tr>
    <td colspan="3" height="20"><div align="left" class="STYLE2"><font color="#CC9900"><strong>��·��Ʒ����</strong></font></div></td>
  </tr>
  <tr height="30">
    <td width="12%" height="30"><font color="#669900"><strong>��һ����</strong></font></td>
    <td height="30" colspan="2">
	  <div align="left">
	 <%if IfComeSet=1 then%>
	�����Ա���ҳ�������̣�
	<%elseif IfComeSet=2 then%>
	�����Ա���ҳ������Ʒ��
	<%elseif IfComeSet=3 then%>
	�򿪸������������۵�ַ��
	<%end if%>
	<%=ComeKey%>	  </div></td>
  </tr>
  <tr height="30">
    <td height="30">&nbsp;</td>
    <td height="30" colspan="2"><a onClick="javascript:clipboardData.setData('text','<%=ComeKey%>');alert('���Ƴɹ�');return false;" href="javascript:void(0)"><img src="../images/key_lai.jpg" width="84" height="18" border="0"></a>&nbsp;
	<a  onClick="jsWindon('<%=jieducm_taobao%>')" href="javascript:void(0)">
	 <img src="../images/taobao_in.jpg" width="96" height="19" border="0"></a></td>
  </tr>
  <tr height="30">
    <td height="30"><font color="#669900"><strong>�ڶ�����</strong></font></td>
    <td height="30" colspan="2">������ʾ��Ϣ��<%=ComeNote%></td>
  </tr>
  <tr height="30">
    <td height="30"><font color="#669900"><strong>��������</strong></font></td>
    <td height="30" colspan="2">������Ʒҳ���ַ��������id����ճ�������棬Ȼ��������֤��Ʒ����ť</td>
  </tr>
  <tr>
    <td height="30"><font color="#669900"><strong>�ƹ�����</strong></font></td>
    <td colspan="2"><%=Shopkeeper%></td>
  </tr>
  <tr>
    <td height="30"><font color="#669900"><strong>����id��</strong></font></td>
    <td colspan="2"><form name="form1" method="post" action="">
	<input name="action" type="hidden" value="jieducm_ok">
	<input name="taobaoid" type="hidden" value="<%=pid%>">
      <label>
        <input type="text" name="prourl_jieducm" style="WIDTH: 260px">
        </label>
      <label>
      <input type="submit" name="Submit" value="��֤��Ʒ">
      </label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="/news.asp?/1482.html" target="_blank"><font color="#FF0000"><strong>ʲô������id��</strong></font></a>
    </form>    </td>
  </tr>
  
  <tr>
    <td colspan="3"><div align="center"> <a href="javascript:void(0)" onClick="Dialog.close()"><img src="../images/close.jpg" border="0"></a></div></td>
  </tr>
</table>
<p>
  <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</p>
</BODY></HTML>
