<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../md5.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
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
	Response.Write("<script>alert('对不起,卡号或密码错误!');history.back();</script>")
	response.End()
Else
       		    
   if stupka=2 and rsm("ok")=0 then				
	  Response.Write("<script>alert('此卡还没有被激活，请联系管理员激活!');history.back();</script>")
	  response.End()	
   elseif   Rsm("ok")=int(1) then
	 Response.Write("<script>alert('此卡正在使用。');history.back();</script>")
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
		rs("class")="充值卡充值"
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
		rs("class")="充值卡充值"
		rs("content")=session("useridname")&"进行在线充值"&paymoney&"元"
		rs("price")="+"&paymoney
		rs("jiegou")="充值成功"
    	rs.update
    	rs.close  	
		call check_jieducm_name(session("useridname"))   			   
    Response.Write("<script>alert('恭喜您!充值成功!\n"&paymoney&"元已存到你的账户!');window.location.href='index.asp';</script>")
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
    alert("充值卡号不能空！");
	document.form.num.focus();
	return false;
	}
 var password=document.form.pwd.value;
  if(password=="")
  {
    alert("充值卡密码不能为空！");
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
    alert("请输入充值金额！");
	document.formbill.v_amount.focus();
	return false;
	}
if   ((document.formbill.v_amount.value.indexOf("-")   ==   0)||!(document.formbill.v_amount.value.indexOf(".")   ==   -1)){   
  alert("充值金额不能为小数或负数");   
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
    alert("请输入充值金额！");
	document.forma.v_amounta.focus();
	return false;
	}
if   ((document.forma.v_amounta.value.indexOf("-")   ==   0)||!(document.forma.v_amounta.value.indexOf(".")   ==   -1)){   
  alert("充值金额不能为小数或负数");   
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
    alert("请输入充值金额！");
	document.forme.p3_Amt.focus();
	return false;
	}
if   ((document.forme.p3_Amt.value.indexOf("-")   ==   0)||!(document.forme.p3_Amt.value.indexOf(".")   ==   -1)){   
  alert("充值金额不能为小数或负数");   
  document.forme.p3_Amt.focus();   
  return   false;   
  }
  return true;  
}
//以下为新增支付宝自动充值
function Confirm()
		 {
		  if (document.myform.TradeNo.value=="")
		  {
		   alert('请输入支付宝交易号!')
		   document.myform.TradeNo.focus();
		   return false;
		  }

		  if (document.myform.Money.value=="")
		  {
		   alert('请输入支付的金额!')
		   document.myform.Money.focus();
		   return false;
		  }
		  if (document.myform.Money.value<1)
		  {
		   alert('支付的金额必需是大于0')
		   document.myform.Money.focus();
		   return false;
		  }
		  if (document.myform.Verifycode.value=="")
		  {
		   alert('请输入验证码！')
		   document.myform.Verifycode.focus();
		   return false;
		  }
		  return true;
		  }
function CheckCZ(){
		   if (document.getElementById("CZMoney").value == "" || document.getElementById("CZMoney").value <= 0){
			   alert('请输入大于0的整数，谢谢！');
			   document.getElementById("CZMoney").focus();
			   return false;
			   }
		   var url="https://lab.alipay.com/life/payment/fill.htm?_ad=c&_adType=alipay_my_home_aide01&optEmail=" + document.getElementById("AlipayAccount").innerHTML + "&totleFee=" + document.getElementById("CZMoney").value + "&title=充值:" + document.getElementById("UserName").value;
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
      <DIV class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 账号充值 &gt;&gt; </DIV>
      <DIV class=pp8><STRONG>账号充值</STRONG></DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <UL>
        <LI>* 帐号充值主要用于发布任务，被接手后，这些钱就转回到您的店铺支付宝了，您也可以通过去接任务，获得存款和发布点</LI>
        <LI>* 平台给您提供了以下几种简单的充值方式，通过任意一种方式均可以实现“充值到平台”！ </LI>
      </UL></DIV>
	 <DIV class=pp8 style="color:red; font-weight:bold">支付宝转账充值 （免手续费推荐）</DIV>
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
              <td width="82%"><p>1、在下框输入要充值的金额数目(只能是整数，例：100)，然后点击&ldquo;确认充值&rdquo;按钮跳到支付宝网站。<br />
2、在支付宝网站页面中完成支付，再到支付宝网站顶部的&ldquo;消费记录&rdquo;，找到刚支付完成的&ldquo;交易号&rdquo;，复制交易号的数字，类似&ldquo;201207280249XXXX&rdquo;<br />
<font color="#FF0000">3、将交易号和充值金额输入到下面的&ldquo;转账支付信息&rdquo;下框中并提交，系统会在1分钟左右自动充值完成。</font></p>
                <p style="color:#F63; font-size:16px; font-weight:bold;">收款支付宝唯一账户：<span id="AlipayAccount"><%=stupzfb%></span>&nbsp; 姓名：*占光</p>
                <table width="100%" border="0">
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left"><input type="hidden" name="UserName" id="UserName" value="<%=session("useridname")%>" readonly="readonly" /></td>
                </tr>
                <tr>
                  <td width="20%" align="right">充值金额：</td>
                  <td width="80%" align="left"><input name="CZMoney" id="CZMoney" type="text" class="textbox" size="25" maxlength="8" onKeyUp="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" onBlur="this.value=this.value.replace(/\D/g,'')" value="100" style="color:#060; font-weight:bold; font-size:14px;" /></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left" height=30><input  id="Chongzhi" type="button" value=" 确定充值 "  name="Chongzhi" onClick="CheckCZ()" /></td>
                </tr>
              </table>
              <br /></td>
            </tr>
          </table>
            
         </td>
        </TR></TBODY></TABLE></DIV>
		
      <DIV class=pp8 style="color:blue; font-weight:bold">填写转账支付信息</DIV>
      <TABLE width="100%" border=0 cellPadding=0 cellSpacing=0>
			<FORM name=myform action="?" method="post">
			<tr  height="30">
           	  <td  align=right width=271>交易号：&nbsp; </td>
              <td width="497" align=left>
                <input name="TradeNo" id="TradeNo" type="text" class="textbox" size="25" maxlength="38" onKeyUp="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" onBlur="this.value=this.value.replace(/\D/g,'')" />
             </td>
			</tr>
			<tr height="30">
           	  <td align=right width=271>充值金额：&nbsp; </td>
              <td align=left>
                <input name="Money" id="Money" type="text" class="textbox" size="25" maxlength="8" onKeyUp="this.value=this.value.replace(/[^0123456789.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^0123456789.]/g,'')" onBlur="this.value=this.value.replace(/[^0123456789.]/g,'')" />
            </td>
			</tr>
			<tr class=tdbg>
           	  <td height=30 align=right width=271>验证码：&nbsp; </td>
              <td align=left><input type="text" maxlength="6" name="Verifycode" id="Verifycode" onFocus="this.value='';" class="textbox" style="width:80px;" autocomplete="off" />&nbsp; <img src="../jieducm_code.asp" alt="验证码" onClick="this.src='../jieducm_code.asp?rnd=' + Math.random();" /></td>
			</tr>
			<tr class=tdbg>
			  <td align=left height=30></td>
			  <td align=left height=30><Input id=Action type=hidden value="SaveAlipPayCZ" name="Action"> 
				<Input  id=Submit type=submit value=" 确定提交 " onClick="return(Confirm())" name=Submit></td>
			</tr>
		 	</FORM>
		  </table>
		  
		   <DIV class=pp8 style="color:red; font-weight:bold">财付通转账充值 （免手续费推荐）</DIV>
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
              <td width="82%"><p>1、在下框输入要充值的金额数目(只能是整数，例：100)，然后点击&ldquo;确认充值&rdquo;按钮跳到财付通网站。<br />
			  2、在财付通网站页面中完成支付，再到财付通网站顶部的“交易查询”，找到刚收支明细详情的“交易单号”，复制交易号的数字，类似&ldquo;1000000000201309210268769XXXX&rdquo;<br />
<font color="#FF0000">3、将交易号和充值金额输入到上面的&ldquo;转账支付信息&rdquo;下框中并提交，系统会在1分钟左右自动充值完成。</font></p>
                <p style="color:#F63; font-size:16px; font-weight:bold;">收款财付通唯一账户：<span id="AlipayAccount"><%=stupcft%></span>&nbsp; 姓名：*占光</p>
                <table width="100%" border="0">
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left"><input type="hidden" name="UserName" id="UserName" value="<%=session("useridname")%>" readonly="readonly" /></td>
                </tr>
                <tr>
                  <td width="20%" align="right">充值金额：</td>
                  <td width="80%" align="left"><input name="CZMoney" id="CZMoney" type="text" class="textbox" size="25" maxlength="8" onKeyUp="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" onBlur="this.value=this.value.replace(/\D/g,'')" value="100" style="color:#060; font-weight:bold; font-size:14px;" /></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left" height=30> 
				  <A href="http://www.tenpay.com/" target=_blank><input  id="Chongzhi" type="button" value=" 确定充值 "  name="Chongzhi" /></A></td>
                </tr>
              </table>
              <br /></td>
            </tr>
          </table>
            
         </td>
        </TR></TBODY></TABLE></DIV>
		  
		    <%if stupcar=1 then%>
      <DIV class=pp8><STRONG>在线充值卡：</STRONG>在线自动充值！（免手续费）</DIV>
      <DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>第一步：购买官方淘宝店铺的充值卡，一次可购买多张；<BR>
            点击进入充值卡商品地址  
			<input name="jieducm_address1" type="hidden" id="jieducm_address1" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address1')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="50元" >
             | <input name="jieducm_address1" type="hidden" id="jieducm_address2" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address2')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="100元" > | <input name="jieducm_address3" type="hidden" id="jieducm_address3" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address3')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="200元" > | <input name="jieducm_address4" type="hidden" id="jieducm_address4" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address4')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="300元" > | <input name="jieducm_address5" type="hidden" id="jieducm_address5" size="9" value="https://me.alipay.com/xxsh">
            <input name="button" type="button" onClick="jsCopy('jieducm_address5')" style="FONT-WEIGHT: bold; WIDTH: 50px; CURSOR: pointer; COLOR: #000000; HEIGHT: 20px" value="500元" > | 
            &lt;--请点击您要充值的数额</TD></TR>
        <TR>
          <TD>第二步：在下面输入购买得到的卡号和密码进行存款充值；</TD></TR>
        <TR>
          <TD>
		  <form name="form" action="" method="post" onSubmit="return CheckForm();">
            <TABLE cellSpacing=0 cellPadding=0 width="90%" border=0>
              <TBODY>
              <TR>
                <TD class=inputtext>卡号：</TD>
                <TD><INPUT id="num" style="WIDTH: 145px; HEIGHT: 15px" maxLength=34 name="num"></TD></TR>
              <TR>
                <TD class=inputtext>密码：</TD>
                <TD><INPUT id="pwd" style="WIDTH: 145px; HEIGHT: 15px" type=password maxLength=16 name="pwd"></TD></TR>
              <TR>
                <TD class=inputtext>
                 </TD>
                <TD><INPUT id=submitadd  type=submit value=充值卡充值 <%if stupcar=0 then%> disabled="disabled" <% end if%>></TD></TR></TBODY></TABLE>
				</form>
				
				</TD></TR>
        <TR>
          <TD>注：拍下充值卡商品并支付，在您确认收货后，联系客服获得密码，请拍下、支付、确认收货、平台充值</TD></TR></TBODY></TABLE></DIV>
     <%end if%>
		  
		  
          
          <%   Dim RS: Set RS=Server.CreateObject("ADODB.RECORDSET")
             RS.Open "Select * From AutoCZ Where UserName = '" & session("useridname") & "' And DateDiff(d,AddDate,Getdate()) < 3 And Status < 3 Order By ID Desc",Conn,1,1
             If Not RS.Eof Then
		%>
        <br />
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
            <tr class=pp8 align=middle>
              <td nowrap width=70 height="25">序号</td>
              <td nowrap width=160>交易号</td>
              <td nowrap width=110>充值金额</td>
              <td nowrap width=160>提交时间</td>
              <td nowrap width=100>状态</td>
              <td nowrap>操作</td>
            </tr>
			<%   Dim i: i = 1
                 Do While Not(RS.Eof) %>
                <tr class=tdbg onMouseOver="this.style.backgroundColor='#EEE'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
                  <td class="splittd" align="center"><%=i%></td>
                  <td class="splittd" align="center"><%=RS("TradeNo")%></td>
                  <td class="splittd" align="center"><%=RS("Money")%> 元</td>
                  <td class="splittd" align="center"><%=RS("AddDate")%></td>
                  <td class="splittd" align="center"><%
				  If RS("Status")=0 Then
				  		Response.Write "<span style=""color:green"">等待充值</span>"
				  ElseIf RS("Status")=1 Then
				  		Response.Write "充值成功"
				  ElseIf RS("Status")=2 Then
				  		Response.Write "<span style=""color:red"">充值失败</span>"
				  End If
				  %></td>
                  <td class="splittd" align="center"><%If RS("Status")=0 Then%><a href="?Action=CZDel&ID=<%=RS("ID")%>" onClick="return confirm('您确定要删除吗？')">删除</a><%End If%>&nbsp;</td>
              	</tr>
            <%	 RS.MoveNext
                 i = i + 1
                 loop
			%>
            <tr>
			  <td colSpan=6 height=40> &nbsp; 提示：只显示3天之内提交的充值记录</td>
			</tr>
        </table>
        <%
             End If
             RS.Close
	         Set RS=Nothing
	   %>
          
          
 </DIV>
 
	  <%if stupwangyin=1 then%>
	  <DIV  class=pp8><STRONG>网上银行在线充值：</STRONG>立即自动到账</DIV>
      <TABLE cellSpacing=0 cellPadding=0 width=100% align=center bgColor=#666666 border=0>
<FORM name=formbill action="../chinabank/Send.asp" method=post target="_blank" onSubmit="return save_onclick()">
    <TBODY>
      <TR>
        <TD bgColor=#ffffff>&nbsp;</TD>
      </TR>
      <TR> 
        <TD bgColor=#ffffff>
        <!--隐藏参数-->
            <!--定单号--><input name="v_oid" type="hidden" maxlength="64" value="" />
            <!--收货人姓名--><input name="v_rcvname" type="hidden" value="">
            <!--收货人地址--><input name="v_rcvaddr" type="hidden" id="v_rcvaddr"  >
            <!--收货人电话--><input name="v_rcvtel" type="hidden" id="v_rcvtel"  >
            <!--收货人邮编--><input name="v_rcvpost" type="hidden" id="v_rcvpost"  value="">
            <!--收货人邮件--><input name="v_rcvemail" type="hidden" >
            <!--收货人手机号码--><input type="hidden" name="v_rcvmobile" value="">
            <!---->
            <!--订货人姓名--><input name="v_ordername" type="hidden" value="">
            <!--订货人地址--><input name="v_orderaddr" type="hidden" id="v_orderaddr"  value="">
            <!--订货人电话--><input name="v_ordertel" type="hidden" id="v_ordertel"  value="">
            <!--订货人邮编--><input name="v_orderpost" type="hidden" id="v_orderpost"  value="">
            <!--订货人邮件--><input name="v_orderemail" type="hidden" value="">
            <!--订货人手机号码--><input type="hidden" name="v_ordermobile" value="">
            <!--备注2--><input name="remark2" type="hidden" id="remark2" value="">
			
                  <table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse">
                    <TBODY>
                     <TR>
                        <TD align=right height=24>会员名：</TD>
                        <TD valign="top"><input name="remark1" type=text id="remark1" value="<%=session("useridname")%>" readOnly
>
&nbsp;&nbsp; <font color="#FF0000">//</font>请确认用户名</TD>
                      </TR>
                      <TR>
                        <TD align=right height=24>充值金额（元）：</TD>
                        <TD height=24><input name="v_amount" type=text   onkeyup="if(isNaN(value))execCommand('undo')" value="100">
                          &nbsp;&nbsp; <font color="#FF0000">*</font>必填项，譬如：<font color="#FF0000">100.00</font><br></TD>
                      </TR>
                      <TR> 
                        <TD align=right height=24>　</TD>
                        <TD valign="top"><input type="submit" name="Submit" value=" 网银在线支付 " <%if stupwangyin="0" then %> disabled<% end if%>>
                        &nbsp; </TD>
                      </TR>
                  </TABLE>        </TD>
      </TR>
  </FORM></TBODY>
</TABLE><%end if%>
<%if stupalipay=1 then%>
      <DIV class=pp8><STRONG>支付宝在线充值：</STRONG>立即自动到账</DIV>
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
                        <TD align=right height=24>会员名：</TD>
                        <TD valign="top"><div align="left">
  <input name="remark1" type=text id="remark1" value="<%=session("useridname")%>" readOnly
>
  &nbsp;&nbsp; <font color="#FF0000">//</font>请确认用户名</div></TD>
                      </TR>
                      <TR>
                        <TD align=right height=24>充值金额（元）：</TD>
                        <TD height=24><div align="left">
  <input name="v_amounta" type=text   onkeyup="if(isNaN(value))execCommand('undo')" value="100">
  &nbsp;&nbsp; <font color="#FF0000">*</font>必填项，譬如：<font color="#FF0000">100.00</font><br>
                        </div></TD>
                      </TR> 
                      <TR> 
                        <TD align=right height=24>　
                          <div align="right"></div></TD>
                        <TD valign="top"><div align="left">
                          <input type="submit" name="Submit" value=" 支付宝在线支付" <%if stupalipay="0" then %> disabled<% end if%>>
                          <input type="hidden" name="subject" id="subject" value="在线充值">
                        </div></TD>
                      </TR>
                  </TBODY>
                  </FORM>
                  </TABLE></TD>
        </TR></TBODY></TABLE></DIV>
		<%end if%>
		<%if stupyibao=1 then%>
		 <DIV class=pp8><STRONG>易宝在线充值：</STRONG>立即自动到账&nbsp;&nbsp;【1%手续费】</DIV>
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
                        <TD align=right height=24>会员名：</TD>
                        <TD valign="top"><div align="left">
  <input name="remark1" type=text id="remark1" value="<%=session("useridname")%>" readOnly
>
  &nbsp;&nbsp; <font color="#FF0000">//</font>请确认用户名</div></TD>
                      </TR>
                      <TR>
                        <TD align=right height=24>充值金额（元）：</TD>
                        <TD height=24><div align="left">
  <input name="p3_Amt" type=text   onkeyup="if(isNaN(value))execCommand('undo')"  value="100">
  &nbsp;&nbsp; <font color="#FF0000">*</font>必填项，譬如：<font color="#FF0000">100.00</font><br>
                        </div></TD>
                      </TR> 
                      <TR> 
                        <TD align=right height=24>　
                          <div align="right"></div></TD>
                        <TD valign="top"><div align="left">
                          <input type="submit" name="Submit" value=" 易宝在线支付"  <%if stupyibao="0" then %> disabled<% end if%>>
                          <input type="hidden" name="subject" id="subject" value="在线充值">
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
	If ID = "" Then Call AlertHistory("提交参数有误！",-1)
	If Not Conn.Execute ("SELECT ID FROM AutoCZ Where UserName = '" & session("useridname") & "' And Status = 0 And ID = " & ID).Eof Then
		Conn.Execute ("DELETE FROM AutoCZ Where UserName = '" & session("useridname") & "' And Status = 0 And ID = " & ID)
		Response.Write("<script language=""Javascript"">alert('删除成功！');window.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
	Else
		Call AlertHistory("抱歉！系统已受理了，不可再删除！",-1)
	End If
End Sub

Sub DoSaveAlipPayCZ()
	Dim TradeNo: TradeNo=Left(Trim(HtmlEncode(Request("TradeNo"))),38)
	Dim Money:Money=Left(Trim(HtmlEncode(Request("Money"))),8)
	jieducm_Len=Len(TradeNo)
	If TradeNo="" Or Money="" Then 
		Call AlertHistory("请输入支付宝的交易号或支付金额！",-1)
		Exit Sub
	End If
	 
	If Not IsNumeric(Money) Then
		Call AlertHistory("支付的金额输入有误！",-1)
		Exit Sub
	End if
	If Round(Money,2) <= 0 Then
		Call AlertHistory("支付的金额必需是大于0",-1)
		Exit Sub
	End if
	IF Trim(Request("Verifycode")) <> CStr(session("GetCode")) Then
		Call AlertHistory("验证码有误，请重新输入！",-1)
		Exit Sub
	End IF
	If Session("AutoCZ") >= 6 Then Call Alert("系统检测到你有恶意或频繁提交行为，达到一定条件系统自动封锁账号处理。","") 
	Dim RS: Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM AutoCZ WHERE TradeNo = '" & TradeNo & "'", Conn, 1, 3
	If Not Rs.EOF Then
		If Rs.RecordCount >= 2 then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("抱歉！系统限制同一交易号不能提交超过2条！\n\n如果您是不小心填错金额，请联系客服核实处理！谢谢！",-1)
			Exit Sub
		ElseIf Rs("Status") = 0 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("该交易号已经存在，请不要重复提交！谢谢！",-1)
			Exit Sub
		ElseIf Rs("Status") = 1 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("该交易号已经充值成功，请不要重复提交！谢谢！",-1)
			Exit Sub
		ElseIf Rs("Status") = 3 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("该交易号系统已禁止充值，请不要恶意提交！谢谢！",-1)
			Exit Sub
		End IF
	End If
	'不存在相同交易号的 和 第一次充值失败(Status=2)的允许再提交一次
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
	Response.Write("<script language=""Javascript"">alert('提交成功！\n\n系统将会在1分钟左右充值完毕！请稍候刷新查看！');window.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
End Sub
%>