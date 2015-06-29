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
'Copyright (C) 2008－2009 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
if session("czm")="" then
 response.Redirect("code.asp")
end if
call myww(1)

%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh" lang="zh" dir="ltr">
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
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
  alert("收藏数量不能为小数或负数");   
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
 <li><a  href="index.asp">淘宝收藏区</a> </li>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="pushmission.asp">发收藏任务</A> </LI>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
  <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A href="MySave.asp">任务仓库</A> </LI>
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
发布收藏任务须知： 
<UL>
  <LI>1. 发布店铺收藏任务时，最好建议不要用二级域名，使用系统分配的店铺地址！
  <LI>2. 严禁发布“成人用品，情趣用品”等无法公开收藏的链接！</LI>
  <li>3. 发布的商品，请保证在收藏任务未完成前，都能保持上架状态！店铺地址实时能访问！</li>
  </UL>
</DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form id="form"  action="jieducm_faok.asp"  method=post>
  <INPUT type=hidden value=ok name=fa> 
  
  <TR>
    <TH><div align="right">掌&nbsp; 柜&nbsp; 名：</div></TH>
    <TD ><%call shopname(1)%></TD>
  </TR>
  
 
    <TH><div align="right">备注：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<input name="info[remarks]" type="radio" id="info[remarks]" value="宝贝收藏任务" checked>
                              <span class="font12l"> 宝贝收藏任务</span> <input type="radio" name="info[remarks]" id="info[remarks]" value="店铺收藏任务">
                               <span class="font12l">店铺收藏任务</span>&nbsp;</TD>
  </TR>
  <TR>
    <TH><div align="right">请输入店铺或宝贝地址：</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl style="WIDTH: 165px" maxLength=355 name=ProUrl> 
    宝贝收藏任务只能发布宝贝链接地址，店铺收藏任务只能发布店铺地址</TD>
  </TR>
  
  <TR>
    <TH><div align="right">收藏数量：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"> 
      <label>
        <INPUT id=product[number] maxLength=255 onchange=changPrice() name=product[number] onKeyUp="if(isNaN(value))execCommand('undo')"> 
        </label>
     
   <SPAN class="blue bold" style="FONT-SIZE: 14px">该任务共需<SPAN style="color:#FF0000" id=collect_price>0</SPAN>个发布点</SPAN> </TD>
  </TR>
   
  <TR>
  <TR>
    <TH><div align="right">收藏标签：</div></TH>
    <TD><INPUT id=tips maxLength=10 name=tips>
      不要超过8个汉字</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> 防来路保护 </LABEL>
      <LABEL>	  <input id=baohu32 disabled type=checkbox checked value=1 
      name=baohu32>
      自动检测宝贝地址和掌柜名</LABEL></TD>
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
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button"  value="发布任务" onClick="ajax()"></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>	    </td>
	    </tr>
  </table>
<script>
function ajax(){
document.form.button.disabled="disabled";
document.form.button.value="系统正在处理中";
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