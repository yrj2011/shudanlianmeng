<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
if session("czm")="" then
 response.Redirect("code.asp")
  response.End()
end if
Sql = "select * from "&jieducm&"_keeper where username='"&session("useridname")&"' and class=1"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then
 Response.Write("<script>alert('���Ȱ󶨵��̺Ų��ܷ�������!');window.location.href='../taobao/myww.asp';</script>")	
 response.End()
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>��Ա����-<%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/base.css" type=text/css rel=stylesheet>
<LINK href="../css/layout.css" type=text/css rel=stylesheet>

</head>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->
<div align="center">  
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 700px">
<UL id=task_nav>
  <LI><A  href="index.asp">��ֵ������</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="pushmission.asp">��Ҫ����</A> </LI>
  <LI><A href="ReMission.asp">�ѽӳ�ֵ</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
   <LI><A href="../user/md5_pay.asp">������ֵ</A> </LI>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>
</div>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
        <DIV class=page>
<DIV class=notes>
<DIV class=inner>
<P class=s><div align="left">������������ǰ��֪��</div> </P>
<UL>
  <LI><div align="left">1.����һ���������񲻿۳����ķ����㣡</div> 
  <LI><div align="left">2. �뱣֤���㹻100Ԫ���ϵĴ�����������������۸���Ӵ������۳������˽ӵ����������Ϊ��Ʒ��ظ��㡣</div>  </LI></UL></DIV></DIV>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form  action="jieducm_faok.asp"  method=post>
  <INPUT type=hidden value=ok name=fa> 
  <TR>
    <TH>�����ƹ�����</TH>
    <TD><div align="left">
      <%call shopname(1)%>
    </div></TD></TR>
  <TR>
    <TH>���񷢲��۸�</TH>
    <TD ><div align="left">
                            <select name="Price1">
                              <option value="100" selected>100</option>
                              <option value="200">200</option>
                              <option value="300">300</option>
                              <option value="400">400</option>
                              <option value="500">500</option>
                              <option value="600">600</option>
                              <option value="1000">1000</option>
                            </select> 
                            Ԫ      </div></TD>
  </TR>
  <TR>
    <TH>����Ʒ�ĵ�ַ��</TH>
    <TD ><div align="left">
      <INPUT class=txtinput id=ProUrl1 style="WIDTH: 350px" maxLength=355 name=ProUrl1> 
      ����д��Ʒҳ����ȷ��ַ</div></TD></TR>
  
  <TR>
    <TH></TH>
    <TD><A href="#" ><FONT color=#ff0000><STRONG>* 
      ������ʱ�ջ�������ƽ̨����ṩ�������ţ���ǿ�������ʱ�ջ�</STRONG></FONT></A></TD></TR>
  <TR>
    <TH>�������ͣ�</TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<div align="left"><LABEL><INPUT name="bei" type="radio" id="bei"  value="�����ջ�����" checked> 
	<SPAN class=font12l>�����ջ�����</SPAN></LABEL>
&nbsp;&nbsp;	<LABEL><INPUT type="radio" name="bei" id="bei"  value="һ����ջ�����"> <SPAN class=font12l>һ����ջ�����</SPAN></LABEL><BR>
    <LABEL><input type="radio" name="bei" id="bei" value="������ջ�����"> <span class="font12l">������ջ�����</span></LABEL>
  <LABEL><input type="radio" name="bei" id="bei" value="������ջ�����"> <span class="font12l">������ջ�����</span></LABEL>
  <BR>
  </div></TD></TR>
  <TR>
    <TH>���������</TH>
    <TD><div align="left">
      <INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100">
      ��������ҿɼ���100������</div></TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD align="left"><div align="left"> <LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> ����·���� </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> �����걣��</LABEL>
	  </div>
	   </TD></TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD align="left"><div align="left">
      <INPUT id=baohu3 disabled type=checkbox CHECKED value=1 
      name=baohu3>
      �Զ���ⱦ����ַ���ƹ���</div></TD></TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD align="left"><LABEL>
    <div align="left">
      <INPUT id=baohu2 type=checkbox value=1 name=baohu2> 
      ���񱣻��������߽��������Ҫ����˲ſ��Կ�����Ʒ��ַ��</div>
    </LABEL></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 30px; PADDING-TOP: 30px" 
    colSpan=2>
      <div align="left">
        <input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button"<% if cunkuan<100 then%> disabled <%end if%> value="��������" onClick="this.disabled=true;document.form.submit();"> 
        </div></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>	    </td>
	    </tr>
  </table>
  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
   call CloseConn()%>
</body>
</html>