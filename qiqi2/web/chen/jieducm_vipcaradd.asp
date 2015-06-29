
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
option explicit
response.buffer=true	
Const PurviewLevel=1
'response.write "此功能被WEBBOY暂时禁止了！"
'response.end
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim rs, sql, strPurview,ok,id,fid,qq,num,iCount,rsm
dim Action,rsr,FoundErr,ErrMsg,tribename,pic,tribeinfo,startid,triberad,tribeid
action=request("action")	
if action="add" then	  
	 Sql= "select * from "&jieducm&"_vipcar where car_name='"&request("car_name")&"'"
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if rs.eof then
	 rs.addnew
		rs("car_name")=request.Form("car_name")
		rs("car_price")=request.Form("car_price")
	    rs("car_date")=request.Form("car_date")
    	rs("car_pic")=request.Form("pic")
		rs("car_content")=request.Form("car_content")
		rs("car_sort")=request.Form("car_sort")
		rs("now")=now()
		rs("car_fabudian")=request.Form("car_fabudian")
    	rs.update
		rs.close
	 else
	   Response.Write("<script>alert('对不起！此点卡名称已经存在!');history.back();</script>")
	   response.End()
	 end if
	 Response.Write("<script>alert('添加成功!');window.location.href='jieducm_vipcar.asp';</script>")
	 response.End()
end if
%>
<html>
<head>
<title>管理员管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function  save_onclick()
{

 var car_name=document.aspnetForm.car_name.value;
  if(car_name=="")
  {
    alert("点卡名称不能为空！");
	document.aspnetForm.car_name.focus();
	return false;
	}
	
 var car_price=document.aspnetForm.car_price.value;
  if(car_price=="")
  {
    alert("点卡价格不能为空！");
	document.aspnetForm.car_price.focus();
	return false;
	}
	 var car_fabudian=document.aspnetForm.car_date.value;
  if(car_fabudian=="")
  {
    alert("点卡有效期不能为空！");
	document.aspnetForm.car_date.focus();
	return false;
	}
	
var pic=document.aspnetForm.pic.value;
  if(pic=="")
  {
    alert("点卡图片不能为空！");
	document.aspnetForm.pic.focus();
	return false;
	}
	

  return true;  
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>网站点卡管理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="jieducm_vipcaradd.asp">添加网站VIP卡  </a> <a  href="jieducm_vipcar.asp">网站VIP卡管理</a></td>
  </tr>
</table>

<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
 <form action="?action=add" method="POST" name="aspnetForm" id="aspnetForm" onSubmit="return save_onclick()">

	<tr>

	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">排序：</DIV></td>
      <td class="tdbg"><label>
        <input name="car_sort" type="text" onKeyUp="if(isNaN(value))execCommand('undo')" size="5">
      （数字越小越排前）</label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">VIP卡名称：</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="car_name" >
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">VIP卡价格：</DIV></td>
      <td class="tdbg"><input type="text" name="car_price" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">送发布点数：</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="car_fabudian" onKeyUp="if(isNaN(value))execCommand('undo')">
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">有效期：</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="car_date" onKeyUp="if(isNaN(value))execCommand('undo')">
      天</label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">VIP卡图片：</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="pic" id="pic">
      </label> &nbsp; 
     <input type="button" name="Submit" value="上传图片"  onClick="window.open('../upload_flash.asp?formname=aspnetForm&editname=pic&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
     （图片规格322*184）	 </td>
    </tr>
    
    <tr class="tdbg">
      <td width="35%" class="tdbg"><DIV align="right">VIP卡简介：</DIV></td>
      <td width="65%" class="tdbg"><label>
        <textarea name="car_content" cols="40" rows="5"></textarea>
      </label></td>
    </tr>
    
  
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input  type="submit" name="Submit" value=" 添&nbsp;&nbsp;加 " style="cursor: hand;background-color: #cccccc;">&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='jieducm_vipcar.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
  </form>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>七七网络服务平台</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
<%
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>
