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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
if session("czm")="" then
 response.Redirect("code.asp")
end if
call myww(1)

%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh" lang="zh" dir="ltr">
<head>
<title>��Ա����-<%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/style.css" type=text/css rel=stylesheet>
<LINK href="../css/base.css" type=text/css rel=stylesheet>
<LINK href="../css/layout.css" type=text/css rel=stylesheet>
<SCRIPT language=JavaScript src="../js/jquery.min.js"></SCRIPT>

</head>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->
<SCRIPT type=text/javascript>
  function changPrice()
  {
  var price, number, flag;
  number = document.getElementById('product[number]').value;
  if   ((number.indexOf("-")   ==   0)||!(number.indexOf(".")   ==   -1)){   
  document.getElementById('product[number]').value=""
  alert("�ղ���������ΪС������");   
  return   false;   
  }
number = number*(<%=stupcang%>*10)/10;
  $('#collect_price').html(number);
  }
  
  </SCRIPT>
<div align="center">  
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
<UL id=task_nav>
 <li><a  href="index.asp">�Ա��ղ���</a> </li>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="pushmission.asp">���ղ�����</A> </LI>
  <LI><A href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
  <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>
</div>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
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
�����ղ�������֪�� 
<UL>
  <LI>1. ���������ղ�����ʱ����ý��鲻Ҫ�ö���������ʹ��ϵͳ����ĵ��̵�ַ��
  <LI>2. �Ͻ�������������Ʒ����Ȥ��Ʒ�����޷������ղص����ӣ�</LI>
  <li>3. ��������Ʒ���뱣֤���ղ�����δ���ǰ�����ܱ����ϼ�״̬�����̵�ַʵʱ�ܷ��ʣ�</li>
  </UL>
</DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form id="form"  action="jieducm_faok.asp"  method=post>
  <INPUT type=hidden value=ok name=fa> 
  
  <TR>
    <TH><div align="right">��&nbsp; ��&nbsp; ����</div></TH>
    <TD ><%call shopname(1)%></TD>
  </TR>
  
 
    <TH><div align="right">��ע��</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<input name="info[remarks]" type="radio" id="info[remarks]" value="�����ղ�����" checked>
                              <span class="font12l"> �����ղ�����</span> <input type="radio" name="info[remarks]" id="info[remarks]" value="�����ղ�����">
                               <span class="font12l">�����ղ�����</span>&nbsp;</TD>
  </TR>
  <TR>
    <TH><div align="right">��������̻򱦱���ַ��</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl style="WIDTH: 165px" maxLength=355 name=ProUrl> 
    �����ղ�����ֻ�ܷ����������ӵ�ַ�������ղ�����ֻ�ܷ������̵�ַ</TD>
  </TR>
  
  <TR>
    <TH><div align="right">�ղ�������</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"> 
      <label>
        <INPUT id=product[number] maxLength=255 onchange=changPrice() name=product[number] onKeyUp="if(isNaN(value))execCommand('undo')"> 
        </label>
     
   <SPAN class="blue bold" style="FONT-SIZE: 14px">��������<SPAN style="color:#FF0000" id=collect_price>0</SPAN>��������</SPAN> </TD>
  </TR>
   
  <TR>
  <TR>
    <TH><div align="right">�ղر�ǩ��</div></TH>
    <TD><INPUT id=tips maxLength=10 name=tips>
      ��Ҫ����8������</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> ����·���� </LABEL>
      <LABEL>	  <input id=baohu32 disabled type=checkbox checked value=1 
      name=baohu32>
      �Զ���ⱦ����ַ���ƹ���</LABEL></TD>
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
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button"  value="��������" onClick="ajax()"></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>	    </td>
	    </tr>
  </table>
<script>
function ajax(){
document.form.button.disabled="disabled";
document.form.button.value="ϵͳ���ڴ�����";
document.form.submit();
}
</script>
  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing%>
</body>
</html>