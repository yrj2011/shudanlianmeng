<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
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
 Response.Write("<script>alert('请先绑定店铺号才能发布任务!');window.location.href='../taobao/myww.asp';</script>")	
 response.End()
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
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
  <LI><A  href="index.asp">充值提现区</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="pushmission.asp">我要提现</A> </LI>
  <LI><A href="ReMission.asp">已接充值</A> </LI>
  <LI><A href="MyMission.asp">已发提现</A> </LI>
   <LI><A href="../user/md5_pay.asp">其它充值</A> </LI>
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
<P class=s><div align="left">发布提现任务前须知：</div> </P>
<UL>
  <LI><div align="left">1.发布一个提现任务不扣除您的发布点！</div> 
  <LI><div align="left">2. 请保证有足够100元以上的存款，您所发布提现任务价格金额从存款里面扣除，有人接到任务后做作为商品款返回给你。</div>  </LI></UL></DIV></DIV>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form  action="jieducm_faok.asp"  method=post>
  <INPUT type=hidden value=ok name=fa> 
  <TR>
    <TH>您的掌柜名：</TH>
    <TD><div align="left">
      <%call shopname(1)%>
    </div></TD></TR>
  <TR>
    <TH>任务发布价格：</TH>
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
                            元      </div></TD>
  </TR>
  <TR>
    <TH>您商品的地址：</TH>
    <TD ><div align="left">
      <INPUT class=txtinput id=ProUrl1 style="WIDTH: 350px" maxLength=355 name=ProUrl1> 
      请填写商品页的正确网址</div></TD></TR>
  
  <TR>
    <TH></TH>
    <TD><A href="#" ><FONT color=#ff0000><STRONG>* 
      发布延时收货的任务，平台免费提供物流单号，并强制买家延时收货</STRONG></FONT></A></TD></TR>
  <TR>
    <TH>发布类型：</TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<div align="left"><LABEL><INPUT name="bei" type="radio" id="bei"  value="马上收货好评" checked> 
	<SPAN class=font12l>马上收货好评</SPAN></LABEL>
&nbsp;&nbsp;	<LABEL><INPUT type="radio" name="bei" id="bei"  value="一天后收货好评"> <SPAN class=font12l>一天后收货好评</SPAN></LABEL><BR>
    <LABEL><input type="radio" name="bei" id="bei" value="二天后收货好评"> <span class="font12l">二天后收货好评</span></LABEL>
  <LABEL><input type="radio" name="bei" id="bei" value="三天后收货好评"> <span class="font12l">三天后收货好评</span></LABEL>
  <BR>
  </div></TD></TR>
  <TR>
    <TH>卖家提醒语：</TH>
    <TD><div align="left">
      <INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100">
      发布后买家可见，100字以内</div></TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD align="left"><div align="left"> <LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> 防来路保护 </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> 防黄钻保护</LABEL>
	  </div>
	   </TD></TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD align="left"><div align="left">
      <INPUT id=baohu3 disabled type=checkbox CHECKED value=1 
      name=baohu3>
      自动检测宝贝地址和掌柜名</div></TD></TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD align="left"><LABEL>
    <div align="left">
      <INPUT id=baohu2 type=checkbox value=1 name=baohu2> 
      任务保护，接受者接你任务后，要你审核才可以看到商品地址！</div>
    </LABEL></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 30px; PADDING-TOP: 30px" 
    colSpan=2>
      <div align="left">
        <input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button"<% if cunkuan<100 then%> disabled <%end if%> value="发布任务" onClick="this.disabled=true;document.form.submit();"> 
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