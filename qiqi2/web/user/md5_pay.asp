<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../md5.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
Select Case Request("Action")
	Case "SaveAlipPayCZ"
		Call DoSaveAlipPayCZ()
	Case "CZDel"
		Call DoCZDel()
End Select

if request.Form<>"" then
num=HtmlEncode(trim(request.form("num")))
pwd=HtmlEncode(trim(request.form("pwd")))

Sql = "select * from  "&jieducm&"_dj586_Jicar where carid='"&num&"' and carpws='"&md5(pwd)&"' "
Set Rsm = Server.CreateObject("Adodb.RecordSet")
Rsm.Open Sql,conn,1,1
if Rsm.eof  Then
	Response.Write("<script>alert('�Բ���,���Ż��������!');history.back();</script>")
	response.End()
Else
       		    
   if stupka=2 and rsm("ok")=0 then				
	  Response.Write("<script>alert('�˿���û�б��������ϵ����Ա����!');history.back();</script>")
	  response.End()	
   elseif   Rsm("ok")=int(1) then
	 Response.Write("<script>alert('�˿�����ʹ�á�');history.back();</script>")
	 response.End()	
   end if
  paymoney =rsm("paymoney")
  paymoney1=paymoney
  day1=date()
  sqlr="update "&jieducm&"_register set cunkuan=cunkuan+"&paymoney1&" where username='"&session("useridname")&"'"
  conn.execute sqlr
  sqlr="update "&jieducm&"_dj586_Jicar set ok=1,userid='"&session("useridname")&"',adddate='"&day1&"' where carid='"&num&"'"
  conn.execute sqlr
			
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="��ֵ����ֵ"
		rs("username")=session("useridname")
    	rs("cunkuan")=paymoney1
    	rs.update
    	rs.close  
			   
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="��ֵ����ֵ"
		rs("content")=session("useridname")&"�������߳�ֵ"&paymoney&"Ԫ"
		rs("price")="+"&paymoney
		rs("jiegou")="��ֵ�ɹ�"
    	rs.update
    	rs.close  	
		call check_jieducm_name(session("useridname"))   			   
    Response.Write("<script>alert('��ϲ��!��ֵ�ɹ�!\n"&paymoney&"Ԫ�Ѵ浽����˻�!');window.location.href='index.asp';</script>")
	response.End()
	end if				  
end if
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
<SCRIPT language=javascript>
function  CheckForm()
{

  var name=document.form.num.value;
  if(name=="")
  {
    alert("��ֵ���Ų��ܿգ�");
	document.form.num.focus();
	return false;
	}
 var password=document.form.pwd.value;
  if(password=="")
  {
    alert("��ֵ�����벻��Ϊ�գ�");
	document.form.pwd.focus();
	return false;
	}
  return true;  
}
function  save_onclick()
{

    var Price=document.formbill.v_amount.value;
  if(Price=="")
  {
    alert("�������ֵ��");
	document.formbill.v_amount.focus();
	return false;
	}
if   ((document.formbill.v_amount.value.indexOf("-")   ==   0)||!(document.formbill.v_amount.value.indexOf(".")   ==   -1)){   
  alert("��ֵ����ΪС������");   
  document.formbill.v_amount.focus();   
  return   false;   
  }
  
  return true;  
}
function save_onclicka()
{
    var Pricea=document.forma.v_amounta.value;
  if(Pricea=="")
  {
    alert("�������ֵ��");
	document.forma.v_amounta.focus();
	return false;
	}
if   ((document.forma.v_amounta.value.indexOf("-")   ==   0)||!(document.forma.v_amounta.value.indexOf(".")   ==   -1)){   
  alert("��ֵ����ΪС������");   
  document.forma.v_amounta.focus();   
  return   false;   
  }
  return true;  
}

function save_onclicke()
{
    var Pricea=document.forme.p3_Amt.value;
  if(Pricea=="")
  {
    alert("�������ֵ��");
	document.forme.p3_Amt.focus();
	return false;
	}
if   ((document.forme.p3_Amt.value.indexOf("-")   ==   0)||!(document.forme.p3_Amt.value.indexOf(".")   ==   -1)){   
  alert("��ֵ����ΪС������");   
  document.forme.p3_Amt.focus();   
  return   false;   
  }
  return true;  
}
//����Ϊ����֧�����Զ���ֵ
function Confirm()
		 {
		  if (document.myform.TradeNo.value=="")
		  {
		   alert('������֧�������׺�!')
		   document.myform.TradeNo.focus();
		   return false;
		  }

		  if (document.myform.Money.value=="")
		  {
		   alert('������֧���Ľ��!')
		   document.myform.Money.focus();
		   return false;
		  }
		  if (document.myform.Money.value<1)
		  {
		   alert('֧���Ľ������Ǵ���0')
		   document.myform.Money.focus();
		   return false;
		  }
		  if (document.myform.Verifycode.value=="")
		  {
		   alert('��������֤�룡')
		   document.myform.Verifycode.focus();
		   return false;
		  }
		  return true;
		  }
function CheckCZ(){
		   if (document.getElementById("CZMoney").value == "" || document.getElementById("CZMoney").value <= 0){
			   alert('���������0��������лл��');
			   document.getElementById("CZMoney").focus();
			   return false;
			   }
		   var url="https://lab.alipay.com/life/payment/fill.htm?_ad=c&_adType=alipay_my_home_aide01&optEmail=" + document.getElementById("AlipayAccount").innerHTML + "&totleFee=" + document.getElementById("CZMoney").value + "&title=��ֵ:" + document.getElementById("UserName").value;
		   window.open(url);
		  }
</SCRIPT>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>



<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
 <DIV class=pp9>
      <DIV style="PADDING-BOTTOM: 15px; WIDTH: 97%">
      <DIV class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; �˺ų�ֵ &gt;&gt; </DIV>
      <DIV class=pp8><STRONG>�˺ų�ֵ</STRONG></DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <UL>
        <LI>* �ʺų�ֵ��Ҫ���ڷ������񣬱����ֺ���ЩǮ��ת�ص����ĵ���֧�����ˣ���Ҳ����ͨ��ȥ�����񣬻�ô��ͷ�����</LI>
        <LI>* ƽ̨�����ṩ�����¼��ּ򵥵ĳ�ֵ��ʽ��ͨ������һ�ַ�ʽ������ʵ�֡���ֵ��ƽ̨���� </LI>
      </UL></DIV>
	 <DIV class=pp8 style="color:red; font-weight:bold">֧����ת�˳�ֵ �����������Ƽ���</DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD align=middle></TD>
        </TR>
        <TR>
          <td align=left><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="18%" align="center"><img src="../images/jieducm_03.gif" width="109" height="55"></td>
              <td width="82%"><p>1�����¿�����Ҫ��ֵ�Ľ����Ŀ(ֻ��������������100)��Ȼ����&ldquo;ȷ�ϳ�ֵ&rdquo;��ť����֧������վ��<br />
2����֧������վҳ�������֧�����ٵ�֧������վ������&ldquo;���Ѽ�¼&rdquo;���ҵ���֧����ɵ�&ldquo;���׺�&rdquo;�����ƽ��׺ŵ����֣�����&ldquo;201207280249XXXX&rdquo;<br />
<font color="#FF0000">3�������׺źͳ�ֵ������뵽�����&ldquo;ת��֧����Ϣ&rdquo;�¿��в��ύ��ϵͳ����1���������Զ���ֵ��ɡ�</font></p>
                <p style="color:#F63; font-size:16px; font-weight:bold;">�տ�֧����Ψһ�˻���<span id="AlipayAccount"><%=stupzfb%></span>&nbsp; ������*ռ��</p>
                <table width="100%" border="0">
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left"><input type="hidden" name="UserName" id="UserName" value="<%=session("useridname")%>" readonly="readonly" /></td>
                </tr>
                <tr>
                  <td width="20%" align="right">��ֵ��</td>
                  <td width="80%" align="left"><input name="CZMoney" id="CZMoney" type="text" class="textbox" size="25" maxlength="8" onKeyUp="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" onBlur="this.value=this.value.replace(/\D/g,'')" value="100" style="color:#060; font-weight:bold; font-size:14px;" /></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left" height=30><input  id="Chongzhi" type="button" value=" ȷ����ֵ "  name="Chongzhi" onClick="CheckCZ()" /></td>
                </tr>
              </table>
              <br /></td>
            </tr>
          </table>
            
         </td>
        </TR></TBODY></TABLE></DIV>
		
      <DIV class=pp8 style="color:blue; font-weight:bold">��дת��֧����Ϣ</DIV>
      <TABLE width="100%" border=0 cellPadding=0 cellSpacing=0>
			<FORM name=myform action="?" method="post">
			<tr  height="30">
           	  <td  align=right width=271>���׺ţ�&nbsp; </td>
              <td width="497" align=left>
                <input name="TradeNo" id="TradeNo" type="text" class="textbox" size="25" maxlength="38" onKeyUp="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" onBlur="this.value=this.value.replace(/\D/g,'')" />
             </td>
			</tr>
			<tr height="30">
           	  <td align=right width=271>��ֵ��&nbsp; </td>
              <td align=left>
                <input name="Money" id="Money" type="text" class="textbox" size="25" maxlength="8" onKeyUp="this.value=this.value.replace(/[^0123456789.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^0123456789.]/g,'')" onBlur="this.value=this.value.replace(/[^0123456789.]/g,'')" />
            </td>
			</tr>
			<tr class=tdbg>
           	  <td height=30 align=right width=271>��֤�룺&nbsp; </td>
              <td align=left><input type="text" maxlength="6" name="Verifycode" id="Verifycode" onFocus="this.value='';" class="textbox" style="width:80px;" autocomplete="off" />&nbsp; <img src="../jieducm_code.asp" alt="��֤��" onClick="this.src='../jieducm_code.asp?rnd=' + Math.random();" /></td>
			</tr>
			<tr class=tdbg>
			  <td align=left height=30></td>
			  <td align=left height=30><Input id=Action type=hidden value="SaveAlipPayCZ" name="Action"> 
				<Input  id=Submit type=submit value=" ȷ���ύ " onClick="return(Confirm())" name=Submit></td>
			</tr>
		 	</FORM>
		  </table>
		  
		   <DIV class=pp8 style="color:red; font-weight:bold">�Ƹ�ͨת�˳�ֵ �����������Ƽ���</DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD align=middle></TD>
        </TR>
        <TR>
          <td align=left><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="18%" align="center"><img src="../images/cft.gif" width="109" height="55"></td>
              <td width="82%"><p>1�����¿�����Ҫ��ֵ�Ľ����Ŀ(ֻ��������������100)��Ȼ����&ldquo;ȷ�ϳ�ֵ&rdquo;��ť�����Ƹ�ͨ��վ��<br />
			  2���ڲƸ�ͨ��վҳ�������֧�����ٵ��Ƹ�ͨ��վ�����ġ����ײ�ѯ�����ҵ�����֧��ϸ����ġ����׵��š������ƽ��׺ŵ����֣�����&ldquo;1000000000201309210268769XXXX&rdquo;<br />
<font color="#FF0000">3�������׺źͳ�ֵ������뵽�����&ldquo;ת��֧����Ϣ&rdquo;�¿��в��ύ��ϵͳ����1���������Զ���ֵ��ɡ�</font></p>
                <p style="color:#F63; font-size:16px; font-weight:bold;">�տ�Ƹ�ͨΨһ�˻���<span id="AlipayAccount"><%=stupcft%></span>&nbsp; ������*ռ��</p>
                <table width="100%" border="0">
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left"><input type="hidden" name="UserName" id="UserName" value="<%=session("useridname")%>" readonly="readonly" /></td>
                </tr>
                <tr>
                  <td width="20%" align="right">��ֵ��</td>
                  <td width="80%" align="left"><input name="CZMoney" id="CZMoney" type="text" class="textbox" size="25" maxlength="8" onKeyUp="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" onBlur="this.value=this.value.replace(/\D/g,'')" value="100" style="color:#060; font-weight:bold; font-size:14px;" /></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left" height=30> 
				  <A href="http://www.tenpay.com/" target=_blank><input  id="Chongzhi" type="button" value=" ȷ����ֵ "  name="Chongzhi" /></A></td>
                </tr>
              </table>
              <br /></td>
            </tr>
          </table>
            
         </td>
        </TR></TBODY></TABLE></DIV>
		  
		    <%if stupcar=1 then%>
      <DIV class=pp8><STRONG>���߳�ֵ����</STRONG>�����Զ���ֵ�����������ѣ�</DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>��һ��������ٷ��Ա����̵ĳ�ֵ����һ�οɹ�����ţ�<BR>
            ��������ֵ����Ʒ��ַ  
			<input name="jieducm_address1" type="hidden" id="jieducm_address1" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address1')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="50Ԫ" >
             | <input name="jieducm_address1" type="hidden" id="jieducm_address2" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address2')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="100Ԫ" > | <input name="jieducm_address3" type="hidden" id="jieducm_address3" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address3')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="200Ԫ" > | <input name="jieducm_address4" type="hidden" id="jieducm_address4" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address4')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="300Ԫ" > | <input name="jieducm_address5" type="hidden" id="jieducm_address5" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address5')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="500Ԫ" > | 
            &lt;--������Ҫ��ֵ������</TD></TR>
        <TR>
          <TD>�ڶ��������������빺��õ��Ŀ��ź�������д���ֵ��</TD></TR>
        <TR>
          <TD>
		  <form name="form" action="" method="post" onSubmit="return CheckForm();">
            <TABLE cellSpacing=0 cellPadding=0 width="90%" border=0>
              <TBODY>
              <TR>
                <TD class=inputtext>���ţ�</TD>
                <TD><INPUT id="num" style="WIDTH: 145px; HEIGHT: 15px" maxLength=34 name="num"></TD></TR>
              <TR>
                <TD class=inputtext>���룺</TD>
                <TD><INPUT id="pwd" style="WIDTH: 145px; HEIGHT: 15px" type=password maxLength=16 name="pwd"></TD></TR>
              <TR>
                <TD class=inputtext>
                 </TD>
                <TD><INPUT id=submitadd  type=submit value=��ֵ����ֵ <%if stupcar=0 then%> disabled="disabled" <% end if%>></TD></TR></TBODY></TABLE>
				</form>
				
				</TD></TR>
        <TR>
          <TD>ע�����³�ֵ����Ʒ��֧��������ȷ���ջ�����ϵ�ͷ�������룬�����¡�֧����ȷ���ջ���ƽ̨��ֵ</TD></TR></TBODY></TABLE></DIV>
     <%end if%>
		  
		  
          
          <%   Dim RS: Set RS=Server.CreateObject("ADODB.RECORDSET")
             RS.Open "Select * From AutoCZ Where UserName = '" & session("useridname") & "' And DateDiff(d,AddDate,Getdate()) < 3 And Status < 3 Order By ID Desc",Conn,1,1
             If Not RS.Eof Then
		%>
        <br />
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
            <tr class=pp8 align=middle>
              <td nowrap width=70 height="25">���</td>
              <td nowrap width=160>���׺�</td>
              <td nowrap width=110>��ֵ���</td>
              <td nowrap width=160>�ύʱ��</td>
              <td nowrap width=100>״̬</td>
              <td nowrap>����</td>
            </tr>
			<%   Dim i: i = 1
                 Do While Not(RS.Eof) %>
                <tr class=tdbg onMouseOver="this.style.backgroundColor='#EEE'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
                  <td class="splittd" align="center"><%=i%></td>
                  <td class="splittd" align="center"><%=RS("TradeNo")%></td>
                  <td class="splittd" align="center"><%=RS("Money")%> Ԫ</td>
                  <td class="splittd" align="center"><%=RS("AddDate")%></td>
                  <td class="splittd" align="center"><%
				  If RS("Status")=0 Then
				  		Response.Write "<span style=""color:green"">�ȴ���ֵ</span>"
				  ElseIf RS("Status")=1 Then
				  		Response.Write "��ֵ�ɹ�"
				  ElseIf RS("Status")=2 Then
				  		Response.Write "<span style=""color:red"">��ֵʧ��</span>"
				  End If
				  %></td>
                  <td class="splittd" align="center"><%If RS("Status")=0 Then%><a href="?Action=CZDel&ID=<%=RS("ID")%>" onClick="return confirm('��ȷ��Ҫɾ����')">ɾ��</a><%End If%>&nbsp;</td>
              	</tr>
            <%	 RS.MoveNext
                 i = i + 1
                 loop
			%>
            <tr>
			  <td colSpan=6 height=40> &nbsp; ��ʾ��ֻ��ʾ3��֮���ύ�ĳ�ֵ��¼</td>
			</tr>
        </table>
        <%
             End If
             RS.Close
	         Set RS=Nothing
	   %>
          
          
 </DIV>
 
	  <%if stupwangyin=1 then%>
	  <DIV  class=pp8><STRONG>�����������߳�ֵ��</STRONG>�����Զ�����</DIV>
      <TABLE cellSpacing=0 cellPadding=0 width=100% align=center bgColor=#666666 border=0>
<FORM name=formbill action="../chinabank/Send.asp" method=post target="_blank" onSubmit="return save_onclick()">
    <TBODY>
      <TR>
        <TD bgColor=#ffffff>&nbsp;</TD>
      </TR>
      <TR> 
        <TD bgColor=#ffffff>
        <!--���ز���-->
            <!--������--><input name="v_oid" type="hidden" maxlength="64" value="" />
            <!--�ջ�������--><input name="v_rcvname" type="hidden" value="">
            <!--�ջ��˵�ַ--><input name="v_rcvaddr" type="hidden" id="v_rcvaddr"  >
            <!--�ջ��˵绰--><input name="v_rcvtel" type="hidden" id="v_rcvtel"  >
            <!--�ջ����ʱ�--><input name="v_rcvpost" type="hidden" id="v_rcvpost"  value="">
            <!--�ջ����ʼ�--><input name="v_rcvemail" type="hidden" >
            <!--�ջ����ֻ�����--><input type="hidden" name="v_rcvmobile" value="">
            <!---->
            <!--����������--><input name="v_ordername" type="hidden" value="">
            <!--�����˵�ַ--><input name="v_orderaddr" type="hidden" id="v_orderaddr"  value="">
            <!--�����˵绰--><input name="v_ordertel" type="hidden" id="v_ordertel"  value="">
            <!--�������ʱ�--><input name="v_orderpost" type="hidden" id="v_orderpost"  value="">
            <!--�������ʼ�--><input name="v_orderemail" type="hidden" value="">
            <!--�������ֻ�����--><input type="hidden" name="v_ordermobile" value="">
            <!--��ע2--><input name="remark2" type="hidden" id="remark2" value="">
			
                  <table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse">
                    <TBODY>
                     <TR>
                        <TD align=right height=24>��Ա����</TD>
                        <TD valign="top"><input name="remark1" type=text id="remark1" value="<%=session("useridname")%>" readOnly
>
&nbsp;&nbsp; <font color="#FF0000">//</font>��ȷ���û���</TD>
                      </TR>
                      <TR>
                        <TD align=right height=24>��ֵ��Ԫ����</TD>
                        <TD height=24><input name="v_amount" type=text   onkeyup="if(isNaN(value))execCommand('undo')" value="100">
                          &nbsp;&nbsp; <font color="#FF0000">*</font>�����Ʃ�磺<font color="#FF0000">100.00</font><br></TD>
                      </TR>
                      <TR> 
                        <TD align=right height=24>��</TD>
                        <TD valign="top"><input type="submit" name="Submit" value=" ��������֧�� " <%if stupwangyin="0" then %> disabled<% end if%>>
                        &nbsp; </TD>
                      </TR>
                  </TABLE>        </TD>
      </TR>
  </FORM></TBODY>
</TABLE><%end if%>
<%if stupalipay=1 then%>
      <DIV class=pp8><STRONG>֧�������߳�ֵ��</STRONG>�����Զ�����</DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD align=middle></TD>
        </TR>
        <TR>
          <TD align=middle><table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse">
 <FORM name=forma action="../alipay/index.asp" method=post target="_blank" onSubmit="return save_onclicka()">

                    <TBODY>
                     <TR>
                        <TD align=right height=24>��Ա����</TD>
                        <TD valign="top"><div align="left">
  <input name="remark1" type=text id="remark1" value="<%=session("useridname")%>" readOnly
>
  &nbsp;&nbsp; <font color="#FF0000">//</font>��ȷ���û���</div></TD>
                      </TR>
                      <TR>
                        <TD align=right height=24>��ֵ��Ԫ����</TD>
                        <TD height=24><div align="left">
  <input name="v_amounta" type=text   onkeyup="if(isNaN(value))execCommand('undo')" value="100">
  &nbsp;&nbsp; <font color="#FF0000">*</font>�����Ʃ�磺<font color="#FF0000">100.00</font><br>
                        </div></TD>
                      </TR> 
                      <TR> 
                        <TD align=right height=24>��
                          <div align="right"></div></TD>
                        <TD valign="top"><div align="left">
                          <input type="submit" name="Submit" value=" ֧��������֧��" <%if stupalipay="0" then %> disabled<% end if%>>
                          <input type="hidden" name="subject" id="subject" value="���߳�ֵ">
                        </div></TD>
                      </TR>
                  </TBODY>
                  </FORM>
                  </TABLE></TD>
        </TR></TBODY></TABLE></DIV>
		<%end if%>
		<%if stupyibao=1 then%>
		 <DIV class=pp8><STRONG>�ױ����߳�ֵ��</STRONG>�����Զ�����&nbsp;&nbsp;��1%�����ѡ�</DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD align=middle></TD>
        </TR>
        <TR>
          <TD align=middle><table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse">
 <FORM name=forme action="../jieducm_yibao/req.asp" method=post target="_blank" onSubmit="return save_onclicke()">

                    <TBODY>
                     <TR>
                        <TD align=right height=24>��Ա����</TD>
                        <TD valign="top"><div align="left">
  <input name="remark1" type=text id="remark1" value="<%=session("useridname")%>" readOnly
>
  &nbsp;&nbsp; <font color="#FF0000">//</font>��ȷ���û���</div></TD>
                      </TR>
                      <TR>
                        <TD align=right height=24>��ֵ��Ԫ����</TD>
                        <TD height=24><div align="left">
  <input name="p3_Amt" type=text   onkeyup="if(isNaN(value))execCommand('undo')"  value="100">
  &nbsp;&nbsp; <font color="#FF0000">*</font>�����Ʃ�磺<font color="#FF0000">100.00</font><br>
                        </div></TD>
                      </TR> 
                      <TR> 
                        <TD align=right height=24>��
                          <div align="right"></div></TD>
                        <TD valign="top"><div align="left">
                          <input type="submit" name="Submit" value=" �ױ�����֧��"  <%if stupyibao="0" then %> disabled<% end if%>>
                          <input type="hidden" name="subject" id="subject" value="���߳�ֵ">
                        </div></TD>
                      </TR>
                  </TBODY>
                  </FORM>
                  </TABLE></TD>
        </TR></TBODY></TABLE></DIV>
		
		<%end if%>
		 </DIV>
 </DIV></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
<%

Function AlertHistory(SuccessStr, N)
	Response.Write("<script language=""Javascript""> alert('" & SuccessStr & "');history.back(" & N & ");</script>")
	Response.End()
End Function
	
Sub DoCZDel()
	Dim ID: ID = HtmlEncode(Request("ID"))
	If ID = "" Then Call AlertHistory("�ύ��������",-1)
	If Not Conn.Execute ("SELECT ID FROM AutoCZ Where UserName = '" & session("useridname") & "' And Status = 0 And ID = " & ID).Eof Then
		Conn.Execute ("DELETE FROM AutoCZ Where UserName = '" & session("useridname") & "' And Status = 0 And ID = " & ID)
		Response.Write("<script language=""Javascript"">alert('ɾ���ɹ���');window.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
	Else
		Call AlertHistory("��Ǹ��ϵͳ�������ˣ�������ɾ����",-1)
	End If
End Sub

Sub DoSaveAlipPayCZ()
	Dim TradeNo: TradeNo=Left(Trim(HtmlEncode(Request("TradeNo"))),38)
	Dim Money:Money=Left(Trim(HtmlEncode(Request("Money"))),8)
	jieducm_Len=Len(TradeNo)
	If TradeNo="" Or Money="" Then 
		Call AlertHistory("������֧�����Ľ��׺Ż�֧����",-1)
		Exit Sub
	End If
	 
	If Not IsNumeric(Money) Then
		Call AlertHistory("֧���Ľ����������",-1)
		Exit Sub
	End if
	If Round(Money,2) <= 0 Then
		Call AlertHistory("֧���Ľ������Ǵ���0",-1)
		Exit Sub
	End if
	IF Trim(Request("Verifycode")) <> CStr(session("GetCode")) Then
		Call AlertHistory("��֤���������������룡",-1)
		Exit Sub
	End IF
	If Session("AutoCZ") >= 6 Then Call Alert("ϵͳ��⵽���ж����Ƶ���ύ��Ϊ���ﵽһ������ϵͳ�Զ������˺Ŵ���","") 
	Dim RS: Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM AutoCZ WHERE TradeNo = '" & TradeNo & "'", Conn, 1, 3
	If Not Rs.EOF Then
		If Rs.RecordCount >= 2 then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("��Ǹ��ϵͳ����ͬһ���׺Ų����ύ����2����\n\n������ǲ�С����������ϵ�ͷ���ʵ����лл��",-1)
			Exit Sub
		ElseIf Rs("Status") = 0 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("�ý��׺��Ѿ����ڣ��벻Ҫ�ظ��ύ��лл��",-1)
			Exit Sub
		ElseIf Rs("Status") = 1 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("�ý��׺��Ѿ���ֵ�ɹ����벻Ҫ�ظ��ύ��лл��",-1)
			Exit Sub
		ElseIf Rs("Status") = 3 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("�ý��׺�ϵͳ�ѽ�ֹ��ֵ���벻Ҫ�����ύ��лл��",-1)
			Exit Sub
		End IF
	End If
	'��������ͬ���׺ŵ� �� ��һ�γ�ֵʧ��(Status=2)���������ύһ��
	Rs.AddNew
	Rs("UserName") = session("useridname")
	RS("RealName") = ""
	RS("TradeNo") = TradeNo
	Rs("Money") = Round(Money,2)
	Rs("Status") = 0
	Rs("AddDate") = Now
	Rs("IP") = Request.ServerVariables("REMOTE_ADDR")
	Rs.Update
	Rs.Close
	Set Rs = Nothing
	Session("AutoCZ") = Session("AutoCZ") + 1
	Response.Write("<script language=""Javascript"">alert('�ύ�ɹ���\n\nϵͳ������1�������ҳ�ֵ��ϣ����Ժ�ˢ�²鿴��');window.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
End Sub
%>