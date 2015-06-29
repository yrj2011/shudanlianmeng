<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
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
function  save_onclick()
{

    var Price=document.form.PushNum.value;
  if(Price=="")
  {
    alert("请输入购买个数！");
	document.form.PushNum.focus();
	return false;
	}
  if(Price<1)
	  
  {
    alert("请输入大于1的数!");
	document.form.PushNum.focus();
	return false;
	}
  if(!Number(Price))	  
  {
    alert("请您输入数字!");
	document.form.PushNum.focus();
	return false;
	}
if   ((document.form.PushNum.value.indexOf("-")   ==   0)||!(document.form.PushNum.value.indexOf(".")   ==   -1)){   
  alert("不能为小数或负数");   
  document.form.PushNum.focus();   
  return   false;   
  }
	 var code=document.form.code.value;
  if(code=="")
  {
    alert("操作码不能为空！");
	document.form.code.focus();
	return false;
	}
  return true;  
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
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 购买发布点 &gt;&gt; </div>
                    <div class=pp8><strong>购买发布点</strong></div>
                    <br>
                   
					 <DIV style="CLEAR: both; WIDTH: 710px;  BACKGROUND-COLOR: #f3f8fe">
<DIV 
style="CLEAR: both; PADDING-RIGHT: 30px; PADDING-LEFT: 30px; HEIGHT: 100%">

<DIV> </DIV>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; BACKGROUND: url(/images/shangyi.gif) no-repeat left 50%; PADDING-BOTTOM: 5px; COLOR: #004000; PADDING-TOP: 7px; BORDER-BOTTOM: #abbec8 1px dashed"><SPAN 
style="COLOR: red">增值服务卡购买(全自动购买获得) 注：想获得发布点，你也可以去接任务，不一定需要购买的！</SPAN><SPAN 
id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="FONT-SIZE: 20px; COLOR: red"></SPAN></DIV>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; PADDING-BOTTOM: 5px; COLOR: #004000; LINE-HEIGHT: 70px; PADDING-TOP: 7px; BORDER-BOTTOM: #abbec8 1px dashed; ">
  <div align="center"><a href="md5_pay.asp" >
  <IMG onMouseOver="this.src ='../images/pay_link.gif'" onMouseOut="this.src='../images/pay.gif'" src="../images/pay.gif" border=0></A></div>
</DIV>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; PADDING-BOTTOM: 5px; PADDING-TOP: 15px">
<DIV style="FLOAT: left; WIDTH: 350px"><IMG 
src="../images/jieducm_258shua.gif"></DIV>
  <form name="form" method="post" action="chongzhi.asp" id="form"  onSubmit="return save_onclick()">
  <input name="action" type="hidden" value="chongzhi">
<DIV style="FLOAT: left; WIDTH: 250px; LINE-HEIGHT: 150%; PADDING-TOP: 20px">
<SPAN style="COLOR: red">拥有发布点，就可以发布自己的任务，获得好评。 想要轻轻松松立即到钻吗？请立即购买发布点！<br>vip会员
<%if stupkuanvip <1 then
response.Write("0")
end if
response.Write(stupkuanvip)
%>
元/个,普通会员:
<%if stupkuan <1 then
response.Write("0")
end if
response.Write(stupkuan)
%>元/个</SPAN><BR>

购买数量：<INPUT id=PushNum style="WIDTH: 30px" value=10 name=PushNum onKeyUp="if(isNaN(value))execCommand('undo')">
个/每个
<%
if vip=1 then
if stupkuanvip<1 then
response.Write("0")
end if
response.Write(stupkuanvip)
else
if stupkuan<1 then
response.Write("0")
end if
response.Write(stupkuan)
end if
%>元<BR>
操作码：<INPUT name=code type="password" id=code size="10">  <A style="COLOR: red; TEXT-DECORATION: underline" class=renwu-link  href="../user/paypoint.asp">去交易市场看看</A>
<BR>

<INPUT id=ImageButton1  onClick="return confirm('您确定购买发布点吗？');" type=image src="../images/buy1.png" name=ImageButton1></DIV>
</form>
</DIV>
<%
   	sql="SELECT * FROM "&jieducm&"_car  order by car_sort"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	do while not rs.eof
%>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; PADDING-BOTTOM: 5px; PADDING-TOP: 15px">
<DIV style="FLOAT: left; WIDTH: 350px"><IMG src="../<%=rs("car_pic")%>"></DIV>
<form name="Form2" method="post" action="chongzhi.asp" id="Form2">
<input name="action" type="hidden" value="carchong">
<input name="id" type="hidden" value="<%=rs("id")%>">
<DIV style="FLOAT: left; WIDTH: 200px; LINE-HEIGHT: 150%; PADDING-TOP: 20px">
<SPAN style="COLOR: red"><%=rs("car_content")%></SPAN><BR>
发布点：<%=rs("car_fabudian")%>个<BR>
<%=rs("car_name")%>：<%=rs("car_price")%>元<BR>
<INPUT id=ImageButton2  onClick="return confirm('您确定要购买<%=rs("car_name")%>吗？');" type=image src="../images/buy1.png" name=ImageButton2>
</DIV>
</form>
</DIV>
<%
Rs.MoveNext
Loop
end if%>

<DIV></DIV>
</DIV></DIV>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
