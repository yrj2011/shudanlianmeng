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
 response.Redirect("code.asp?id=fa")
end if

Sql = "select * from "&jieducm&"_keeper where username='"&session("useridname")&"' and class=2"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if rs.eof then
	 rs.close
	 set rs=nothing
	 Response.Write("<script>alert('���Ȱ󶨵��̺Ų��ܷ�������!');window.location.href='../pai/myww.asp';</script>")	
	 response.End()
	 end if
	
response.charset="gb2312"
action=request.QueryString("go")
if action="check" then
url = request("url")
nickname = trim(request("nickname"))

content = GetNickName(url)
if nickname<>content then
echo "no"
else
echo "yes"
end if

response.End()
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh" lang="zh" dir="ltr">
<HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/style.css" type=text/css rel=stylesheet>
<LINK href="../css/base.css" type=text/css rel=stylesheet>
<LINK href="../css/layout.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
 <!--#include file=../inc/jieducm_top.asp-->
 <!--#include file=../taobao/jieducm_toptao.asp-->
<div align="center"> 
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 700px">
<UL id=paipai_task_nav>
  <LI><A  href="index.asp">���Ļ�����</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/paipai_task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff"  href="pushmission.asp">��������</A> </LI>
  <LI><A href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
   <LI><A href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI></UL>
  </DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG 
src="../images/paipai_task_nav_line.gif"></DIV>
</DIV>
</div>

<table width="910"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr height="1">
                  <td height="5"></td>
      </tr>
  </table>
              <DIV class=page>
<DIV class=addtask-wrap>
<DIV class=inner>
��������ǰ��֪�� 
<UL>
  <LI>1. ����һ�������۳����Ĵ��뱣֤����ƽ̨�����㣻 
  <LI>2. �뱣֤���ڱ�վ���㹻�Ĵ�������������۸���Ӵ������۳������˽ӵ����������Ϊ��Ʒ��ظ��㡣 </LI>
  <li>3. �벻Ҫ����˷������񣬷���ᵼ�����������޷��ڱ�վ������</li>
  </UL></DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form id="form"  action="jieducm_faok.asp"  method=post>
  <INPUT type=hidden value=ok name=action> 
  <TR>
    <TH><div align="right">��Ʒ����ۣ�</div></TH>
    <TD><input name=Price1 id=Price1 size="10" maxlength=6   onKeyUp="if(isNaN(value))execCommand('undo')" >Ԫע��(�ü۸��ǰ����ʷѵ��ܼ۸�)1-40Ԫ����һ�������㣻41-100Ԫ�����������㣻101-200Ԫ�������������㣻201-500Ԫ�����ĸ������㣻501-1000Ԫ�����������</TD></TR>
  <TR>
    <TH><div align="right">��&nbsp; ��&nbsp; ����</div></TH>
    <TD ><%call shopname(2)%></TD>
  </TR>
  <TR>
    <TH><div align="right">��Ʒ��ַ��</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl1 style="WIDTH: 350px" maxLength=355 name=ProUrl1> 
    ����д��Ʒҳ����ȷ��ַ</TD></TR>
  
 
    <TH><div align="right">�۸��Ƿ���ȣ�</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="isprice" type="radio" id="isprice" value="������" checked>
                              <span class="font12l"> ������</span> <input type="radio" name="isprice" id="isprice" value="���޸ļ۸�">
                               <span class="font12l">���޸ļ۸�</span>&nbsp; ����۸�Ͱ����˷ѵ���Ʒ�ܼ۸��Ƿ���ȣ�</TD>
  </TR>
  <TR>
    <TH><div align="right">��̬���֣�</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="play" type="radio" value="ȫ����5��" checked>
                              <span class="font12l">ȫ����5��</span> <input type="radio" name="play" value="ȫ�������">
                             <span class="font12l"> ȫ�������</span></TD>
  </TR>
   <TR>
    <TH></TH>
    <TD><FONT color=#ff0000><STRONG>* 
      ������ʱ�ջ�������ƽ̨����ṩ�������ţ���ǿ�������ʱ�ջ�</STRONG></FONT></TD></TR>
  <TR>
  <TR>
    <TH><div align="right">��ע��</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<LABEL><INPUT name="bei" type="radio" id="bei"  value="�����ջ�����" checked> 
	<SPAN class=font12l>�����ջ�����</SPAN></LABEL>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	<LABEL><INPUT type="radio" name="bei" id="bei"  value="һ����ջ�����"> <SPAN class=font12l>һ����ջ�����</SPAN></LABEL>(��x*2��������)<BR>
    <LABEL><input type="radio" name="bei" id="bei" value="������ջ�����"> <span class="font12l">������ջ�����</span></LABEL>(��x*2+1��������)
  <LABEL><input type="radio" name="bei" id="bei" value="������ջ�����"> <span class="font12l">������ջ�����</span></LABEL>(��x*2+2��������)<BR>	</TD></TR>
  <TR>
    <TH><div align="right">���������ƣ�</div></TH>
    <TD><select name="Limit" >
                                <option value="1" selected>������</option>
                                <option value="2">����100��������</option>
                                <option value="3">����300��������</option>
                                <option value="4">����ֻ����VIP��Ա</option>
                                                                                          </select></TD>
  </TR>
  <TR>
    <TH><div align="right">���������</div></TH>
    <TD><INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100">
    ��������ҿɼ���100������</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> ����·���� </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> �����걣��</LABEL>	  <input id=baohu32 disabled type=checkbox checked value=1 
      name=baohu32>
      �Զ���ⱦ����ַ���ƹ���</TD>
  </TR>
  
  <TR>
    <TH>&nbsp;</TH>
    <TD><input  name="baohu2" type="checkbox" id="baohu2" value="1"  />  
                  ���񱣻��������߽��������Ҫ����˲ſ��Կ�����Ʒ��ַ��</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><input  name="depot" type="checkbox" id="depot" value="1" />˳����ӵ��ҵ�����ֿ�</LABEL>&nbsp; �ֿⱸע(�������Լ��ֱ���Ʒ)�� 
                               <label>
                               <input name="title" type="text" maxlength="20">
                               </label></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px" 
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button" <% if cunkuan<1 or fabudian<1 then%> disabled <%end if%> value="��������" onClick="ajax()"><% if cunkuan<1 or fabudian<1 then%>���򷢲������1ʱ�����ٷ�����<%end if%></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>	<script>
function ajax(){
document.form.button.disabled="disabled";
document.form.button.value="ϵͳ���ڴ�����";
document.form.submit();
}
</script>	
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
