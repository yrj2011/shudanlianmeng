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
 response.Redirect("code.asp?id=fa")
end if

Sql = "select * from "&jieducm&"_keeper where username='"&session("useridname")&"' and class=2"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if rs.eof then
	 rs.close
	 set rs=nothing
	 Response.Write("<script>alert('请先绑定店铺号才能发布任务!');window.location.href='../pai/myww.asp';</script>")	
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
  <LI><A  href="index.asp">拍拍互动区</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/paipai_task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff"  href="pushmission.asp">发布任务</A> </LI>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
   <LI><A href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A href="MySave.asp">任务仓库</A> </LI></UL>
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
发布任务前须知： 
<UL>
  <LI>1. 发布一个任务会扣除您的存款，请保证您的平台存款充足； 
  <LI>2. 请保证您在本站有足够的存款，您所发布任务价格金额从存款里面扣除，有人接到任务后做作为商品款返回给你。 </LI>
  <li>3. 请不要帮别人发布任务，否则会导致您的朋友无法在本站发布。</li>
  </UL></DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form id="form"  action="jieducm_faok.asp"  method=post>
  <INPUT type=hidden value=ok name=action> 
  <TR>
    <TH><div align="right">商品任务价：</div></TH>
    <TD><input name=Price1 id=Price1 size="10" maxlength=6   onKeyUp="if(isNaN(value))execCommand('undo')" >元注：(该价格是包含邮费的总价格)1-40元，扣一个发布点；41-100元扣两个发布点；101-200元，扣三个发布点；201-500元，扣四个发布点；501-1000元扣五个发布点</TD></TR>
  <TR>
    <TH><div align="right">掌&nbsp; 柜&nbsp; 名：</div></TH>
    <TD ><%call shopname(2)%></TD>
  </TR>
  <TR>
    <TH><div align="right">商品地址：</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl1 style="WIDTH: 350px" maxLength=355 name=ProUrl1> 
    请填写商品页的正确网址</TD></TR>
  
 
    <TH><div align="right">价格是否相等：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="isprice" type="radio" id="isprice" value="金额相等" checked>
                              <span class="font12l"> 金额相等</span> <input type="radio" name="isprice" id="isprice" value="需修改价格">
                               <span class="font12l">需修改价格</span>&nbsp; 任务价格和包括运费的商品总价格是否相等！</TD>
  </TR>
  <TR>
    <TH><div align="right">动态评分：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="play" type="radio" value="全部打5分" checked>
                              <span class="font12l">全部打5分</span> <input type="radio" name="play" value="全部不打分">
                             <span class="font12l"> 全部不打分</span></TD>
  </TR>
   <TR>
    <TH></TH>
    <TD><FONT color=#ff0000><STRONG>* 
      发布延时收货的任务，平台免费提供物流单号，并强制买家延时收货</STRONG></FONT></TD></TR>
  <TR>
  <TR>
    <TH><div align="right">备注：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<LABEL><INPUT name="bei" type="radio" id="bei"  value="马上收货好评" checked> 
	<SPAN class=font12l>马上收货好评</SPAN></LABEL>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	<LABEL><INPUT type="radio" name="bei" id="bei"  value="一天后收货好评"> <SPAN class=font12l>一天后收货好评</SPAN></LABEL>(扣x*2个发布点)<BR>
    <LABEL><input type="radio" name="bei" id="bei" value="二天后收货好评"> <span class="font12l">二天后收货好评</span></LABEL>(扣x*2+1个发布点)
  <LABEL><input type="radio" name="bei" id="bei" value="三天后收货好评"> <span class="font12l">三天后收货好评</span></LABEL>(扣x*2+2个发布点)<BR>	</TD></TR>
  <TR>
    <TH><div align="right">接任务限制：</div></TH>
    <TD><select name="Limit" >
                                <option value="1" selected>不限制</option>
                                <option value="2">限制100积分以上</option>
                                <option value="3">限制300积分以上</option>
                                <option value="4">限制只能是VIP会员</option>
                                                                                          </select></TD>
  </TR>
  <TR>
    <TH><div align="right">卖家提醒语：</div></TH>
    <TD><INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100">
    发布后买家可见，100字以内</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> 防来路保护 </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> 防黄钻保护</LABEL>	  <input id=baohu32 disabled type=checkbox checked value=1 
      name=baohu32>
      自动检测宝贝地址和掌柜名</TD>
  </TR>
  
  <TR>
    <TH>&nbsp;</TH>
    <TD><input  name="baohu2" type="checkbox" id="baohu2" value="1"  />  
                  任务保护，接受者接你任务后，要你审核才可以看到商品地址！</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><input  name="depot" type="checkbox" id="depot" value="1" />顺便添加到我的任务仓库</LABEL>&nbsp; 仓库备注(用于您自己分辨商品)： 
                               <label>
                               <input name="title" type="text" maxlength="20">
                               </label></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px" 
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button" <% if cunkuan<1 or fabudian<1 then%> disabled <%end if%> value="发布任务" onClick="ajax()"><% if cunkuan<1 or fabudian<1 then%>存款或发布点低于1时不能再发任务<%end if%></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>	<script>
function ajax(){
document.form.button.disabled="disabled";
document.form.button.value="系统正在处理中";
document.form.submit();
}
</script>	
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
