<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../md5.asp"-->
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
action=HtmlEncode(trim(request.form("action")))
pid=request("pid")
usernamef=request("usernamet")

if action="ok" then
title=HtmlEncode(trim(request.form("title")))
name1=HtmlEncode(trim(request.form("name")))
pid=HtmlEncode(trim(request.form("pid")))
content=(trim(request.form("content")))
classid=request.form("class")
if classid="0" then
	   Response.Write("<script>alert('请先选择申诉原因!');history.back();</script>")
	   response.end()
elseif content="" then
	   Response.Write("<script>alert('申诉原因及证据不能为空!');history.back();</script>")
	   response.end()
end if
   Sql = "select * from "&jieducm&"_register where username='"&name1&"'"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof Then
	   Response.Write("<script>alert('对不起无此用户!');history.back();</script>")
	   response.end()
	end if

        Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select * From "&jieducm&"_toushu where pid='"&pid&"'" ,Conn,3,3  
		if not rs.eof then
		  Response.Write("<script>alert('此任务已提交过了!');history.back();</script>")
	      response.end() 
		else
	    rs.addnew
		rs("title")=title
		rs("name")=name1
    	rs("pid")=pid
		rs("content")=content
		rs("username")=session("useridname")
		rs("pid")=pid
		rs("bianjie")="否"
		rs("bianjie2")="否"
		rs("guang")="否"
		rs("class")=classid
    	rs.update
    	rs.close
		end if
		
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="申诉"
		rs("content")=session("useridname")&"申诉"&name1&"成功"
		rs("jiegou")="成功"
    	rs.update
    	rs.close	
		set rs=nothing
		Response.Write("<script>alert('申诉成功!');window.location.href='complaintmy.asp';</script>")
	     response.End()
end if
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</HEAD>
<SCRIPT language=javascript>
function  save_onclick12()
{

    var Price=document.formf.title.value;
  if(Price=="")
  {
    alert("请输入标题！");
	document.formf.title.focus();
	return false;
	}

    var qu=document.formf.qu.value;
  if(qu=="")
  {
    alert("请输入任务区！");
	document.formf.qu.focus();
	return false;
	}

 var ProUrl=document.formf.name.value;
  if(ProUrl=="")
  {
    alert("请输入被申诉人！");
	document.formf.name.focus();
	return false;
	}
  
  var Shopkeeper=document.formf.pid.value;
  if(Shopkeeper=="")
  {
    alert("请输入任务ID！");
	document.formf.pid.focus();
	return false;
	}
	if(!Number(Shopkeeper))
	  
  {
    alert("任务ID必须是数字!!");
	document.formf.pid.focus();
	return false;
	}
 var content=document.formf.content.value;
  if(content=="")
  {
    alert("内容不能为空！");
	
	return false;
	}
	
  return true;  
}
</SCRIPT>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
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
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;我要申诉 &gt;&gt; </div>
                    <div class=pp8><strong>我要申诉</strong></div>
					<FORM name=formf method=post action=""  onSubmit="return save_onclick12()">
					<input name="action" type="hidden" value="ok">
					<table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="110" height="40" align="center">标&nbsp;&nbsp;&nbsp;题：</td>
              <td width="490"><input name="title" type="text" id="title" size="25" /> 
              (只受理进行中的任务，已完成的任务不受理！) </td>
            </tr>
            <tr>
              <td height="40" align="center">申诉原因：</td>
              <td><select name="class">			   
               <option value="0">请点击选择申诉原因</option>
                <option value="买家未在淘宝支付，却在平台点了已经支付；">买家未在淘宝支付，却在平台点了已经支付；</option>
                <option value="买家未在淘宝确认收货好评，却在平台点了我已好评">买家未在淘宝确认收货好评，却在平台点了我已好评</option>
                <option value="在我操作2小时后，对方仍未作下一步操作(虚拟任务)">在我操作2小时后，对方仍未作下一步操作(虚拟任务)</option>
				<option value="超过收货时间12小时，买家仍未收货好评(实物任务)">超过收货时间12小时，买家仍未收货好评(实物任务)</option>
				
				<option value="买家在淘宝付款12小时后，卖家仍未发货(实物任务)">买家在淘宝付款12小时后，卖家仍未发货(实物任务)</option>			
				<option value="买家在平台确认好评12小时后，卖家仍未确认好评(实物任务)">买家在平台确认好评12小时后，卖家仍未确认好评(实物任务)</option>
				<option value="未到确认收货时间，买家却提前确认收货">未到确认收货时间，买家却提前确认收货</option>
				<option value="买家个人交易信息不完善或收货人信息填写不规范">买家个人交易信息不完善或收货人信息填写不规范</option>
				<option value="非默认匿名商品，买家却匿名购买和评价">非默认匿名商品，买家却匿名购买和评价</option>
				<option value="买家的平台接手号与淘宝买号不一致">买家的平台接手号与淘宝买号不一致</option>
				
				<option value="买家未按照我的要求写评价">买家未按照我的要求写评价</option>
				<option value="买家用黄钻买号接任务">买家用黄钻买号接任务</option>
				<option value="买家申请退款">买家申请退款</option>
				<option value="卖家发布套餐任务">卖家发布套餐任务</option>
				<option value="其他情况">其他情况</option>
				
              </select></td>
            </tr>
            <tr>
              <td height="40" align="center">被申诉人：</td>
              <td><input name="name" type="text" id="name" size="25"  value="<%=usernamef%>"/></td>
            </tr>
            <tr>
              <td height="40" align="center">任&nbsp;务&nbsp;ID：</td>
              <td><input name="pid" type="text" id="pid" size="25"  value="<%=pid%>"/></td>
            </tr>
            <tr>
              <td height="40" align="center">申诉原因及证据：</td>
              <td>
               <textarea name="Content" id="content" style="display:none"></textarea> 
            
              <iframe id="Editor" name="Editor" src="../HtmlEditor/index.html?ID=content" frameborder="0" marginheight="0"     marginwidth="0" scrolling="No" style="height:320px;width:100%"></iframe>  <span class="STYLE1">
              <p>请上传该任务的淘宝截图（先把截图上传到淘宝图片空间，然后复制地址即可）</p></span>
              </td>
            </tr>
            <tr>
              <td height="70">&nbsp;</td>
              <td valign="middle">&nbsp;&nbsp;&nbsp;
                <input type="submit" name="button" id="button" value="提交" style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px">
                <input type="reset" name="button2" id="button2" value="重置" style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px"></td>
            </tr>
          </table>
					</FORM>
                  </div>
                </div></td>
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
