<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->

<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统**************************************
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
if stupphonestart=1 and dst =0 then
	Response.Write("<script>alert('请先绑定手机才能进行其它操作!');window.location.href='../user/Center_Userlock.asp';</script>")
	response.End()
end if

 action=HtmlEncode(trim(request.form("action")))
if action="ok" then      
 ReNum=HtmlEncode(trim(request.form("ReNum1")))
 ReRName=HtmlEncode(trim(request.form("ReRName1")))
 ReZfb=HtmlEncode(trim(request.form("ReZfb1")))
 czm1=HtmlEncode(trim(request.form("czm")))
 cunkuan2=cunkuan*100
 renum2=ReNum*100
 
if renum="" then
     Response.Write("<script>alert('提现金额不能为空!');history.back();</script>")
	 response.End()
elseif czm1<>czm then
	Response.Write("<script>alert('你输入的操作码不正确!');history.back();</script>")
	 response.End()
elseif int(cunkuan2)<int(renum2) then
	Response.Write("<script>alert('你提现金额不能大于你的存款!');history.back();</script>")
    response.End()
elseif ReRName="" then
     Response.Write("<script>alert('掌柜名不能为空!');history.back();</script>")
	 response.End()
elseif ReZfb="" then
     Response.Write("<script>alert('淘宝商品地址不能为空!');history.back();</script>")
	 response.End()
elseif int(ReNum)<10 then
	Response.Write("<script>alert('低于10元的提现，请以发布任务的形式，让别人接手你的任务，存款就转到你的支付宝了!');history.back();</script>")
    response.End()
elseif not isNumeric(ReNum) then
          Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
          response.End()
end if
  
url=ReZfb
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" href=""[^""]+""[^>]*>(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk=Matches(0).SubMatches(0)

Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" title=""(.*?)"""
set Matches=objRegExp.Execute(wstr)
jieducm_sk3=Matches(0).SubMatches(0)

if (jieducm_sk=shopkeeper)  or (jieducm_sk3=shopkeeper) then
else
	Response.Write("<script>alert('此掌柜名与填写的商品地址不符!请重新输入!');history.back();</script>")
    response.End()
end if

	  Set rs=server.createobject("ADODB.RECORDSET")
     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now) 
	 	rs.open "Select * From "&jieducm&"_tixian where username='"&session("useridname")&"' and (start='0' or start='3') " ,Conn,3,3  
		if not rs.eof then
			Response.Write("<script>alert('上一次提现还未完成!');history.back();</script>")
            response.End()
		else 
	    rs.addnew
		rs("class")=1
		rs("pid")=num
		rs("ReNum")=ReNum
		rs("start")=0
    	rs("ReRName")=ReRName
		rs("ReZfb")=ReZfb
		rs("username")=session("useridname")		
    	rs.update
    	rs.close
		end if
 
 sqlr="update "&jieducm&"_register set cunkuan=cunkuan-"&renum&" where username='"&session("useridname")&"'"
 conn.execute sqlr
	call check_jieducm_name(session("useridname"))	
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="申请提现"
		rs("content")=session("useridname")&"进行申请提现"&ReNum&"元"
		rs("price")="-"&renum
		rs("jiegou")="处理中"
    	rs.update
    	rs.close
   Response.Write("<script>alert('您已成功申请提现,\n我们将在第1-2个工作日内完成将提现转进您的支付宝!');window.location.href='ReMoneyList.asp';</script>")
	response.End()
set rs=nothing
end if
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
 <SCRIPT language=javascript>
function  save_onclick22()
{

    var Price=document.formt.ReNum1.value;
  if(Price=="")
  {
    alert("请输入提现金额！");
	document.formt.ReNum1.focus();
	return false;
	}
 if(!Number(Price))
	  
  {
    alert("请您输入数字!");
	document.formt.ReNum1.focus();
	return false;
	}
	if(Price<10)
	{
	alert("低于10元的提现，请以发布任务的形式，让别人接手你的任务，存款就转到你的支付宝了!");
	document.formt.ReNum1.focus();
	return false;
	}
	
if   ((document.formt.ReNum1.value.indexOf("-")   ==   0)||!(document.formt.ReNum1.value.indexOf(".")   ==   -1)){   
  alert("提现金额不能为小数或负数");   
  document.formt.ToUser.focus();   
  return   false;   
  } 
	
  var ReRName1=document.formt.ReRName1.value;
  if(ReRName1=="")
  {
    alert("请输入您的掌柜名！");
	document.formt.ReRName1.focus();
	return false;
	}

 var ReZfb1=document.formt.ReZfb1.value;
  if(ReZfb1=="")
  {
    alert("请输入您的淘宝商品地址！");
	document.formt.ReZfb1.focus();
	return false;
	}
 var czm=document.formt.czm.value;
  if(czm=="")
  {
    alert("请输入您的操作码！");
	document.formt.czm.focus();
	return false;
	}
  
  return true;  
}

</script>  
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
              
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 申请提现 &gt;&gt; </div>
                    <div class=pp8><strong>申请提现</strong></div>
                    <div style="MARGIN-TOP: 10px; LINE-HEIGHT: 250%">
                      <ul>
                        <li>* 平台提供两种提现方式，淘宝商品提现和支付宝提现和财付通。<br>
                          * 您也可以在提现列表内撤销提现，将自动为您立即返回金额至您用户名。<br>
                          *   您的操作记录中也会有相应的记录信息。<br>
                          *   由于提现量大，商品链接低于10元、支付宝低于10元的提现，请以发布任务的形式，让别人接手您的任务，存款就转到您的店铺支付宝了！ </li>
                      </ul>
                    </div><br> 
 <table width="614" border="0" align="center" cellpadding="0" cellspacing="0"> 
  <form name="formt" action="" method="post" onSubmit="return save_onclick22();">
                <tr>
                  <td height="35" colspan="2" align="center"> 每次只能进行一次提现，本次提现完成后才能进行下一次提现！ </td>
                  </tr>
                <tr>
                  <td height="35" align="right">&nbsp;</td>
                  <td>
				  <input name="class" type="radio"  value="1" checked/>
				  淘宝商品地址提现                  
                  <input type="radio" name="class" value="2" onClick="javascript:location.href='remoney1.asp';"/>
                  支付宝提现     
				   <input type="radio" name="class" value="3" onClick="javascript:location.href='remoney2.asp';"/>
                  财付通提现               </td>
                </tr>
                <tr>
                  <td height="35" align="right">提现人：</td>
                  <td><%=session("useridname")%> </td>
                </tr>
                <tr>
                  <td width="134" height="35" align="right">提现金额：</td>
                  <td width="480"><INPUT id=ReNum1  maxLength=4 name=ReNum1 onKeyUp="if(isNaN(value))execCommand('undo')">
                  元</td>
                </tr>
                <tr>
                  <td height="35" align="right">掌柜：</td>
                  <td><INPUT id=ReRName1 maxLength=50 name=ReRName1>自动检测宝贝地址和掌柜名是否相符</td>
                </tr>
                <tr>
                  <td height="35"  align="right">淘宝商品地址：</td>
                  <td><INPUT name=ReZfb1 id=ReZfb1 size="35" maxLength=255>
                  商品价格必须和提现金额相同，否则会撤销提现</td>
                </tr>
                <tr>
                  <td height="35"><div align="right">操作码：</div></td>
                  <td><input type="password" name="czm" id="czm"></td>
                </tr>
                <tr>
                  <td height="35">&nbsp;</td>
                  <td><input type="submit" name="button" id="button" value="开始提现" />
                    <input type="hidden" name="action" value="ok"></td>
                </tr></FORM>
              </table>


               </td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
