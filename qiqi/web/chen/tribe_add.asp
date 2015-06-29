
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

	  Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='"&request("username")&"' " ,Conn,3,3  
	  if  rs.eof then	
		Response.Write("<script>alert('对不起无此会员!');history.back();</script>")
		response.End()
	  end if
	  
	 Sql= "select * from "&jieducm&"_tribe where username='"&request("username")&"'"
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if rs.eof then
	 rs.addnew
		rs("username")=request("username")
		rs("tribename")=request("tribename")
    	rs("pic")=request("pic")
		rs("tribeinfo")=request("tribeinfo")
		rs("triberad")=request("triberad")
		rs("now")=now()
    	rs.update
		rs.close
	 else
	   Response.Write("<script>alert('对不起！此会员已经是酋长了!');history.back();</script>")
	   response.End()
	 end if
		
 Sql= "select * from "&jieducm&"_tribe where username='"&request("username")&"'"
 Set rs=server.createobject("ADODB.RECORDSET")
 Rs.Open Sql,conn,3,3
 if not rs.eof then
   tribeid=rs("id")
 end if
		
	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select tribeid From "&jieducm&"_register where username='"&request("username")&"' " ,Conn,3,3  
	  if  not rs.eof then	
		rs("tribeid")=tribeid
    	rs.update
		rs.close
	  end if
	  
	 Response.Write("<script>alert('添加成功!');window.location.href='tribemanage.asp';</script>")
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

 var Price=document.aspnetForm.tribename.value;
  if(Price=="")
  {
    alert("部落名不能为空！");
	document.aspnetForm.tribename.focus();
	return false;
	}
	
 var ProUrl=document.aspnetForm.pic.value;
  if(ProUrl=="")
  {
    alert("部落徽章不能为空！");
	document.aspnetForm.pic.focus();
	return false;
	}
	 var triberad=document.aspnetForm.triberad.value;
  if(triberad=="")
  {
    alert("部落激活码不能为空！");
	document.aspnetForm.triberad.focus();
	return false;
	}
	
		 var username=document.aspnetForm.username.value;
  if(username=="")
  {
    alert("酋长不能为空！");
	document.aspnetForm.username.focus();
	return false;
	}
	
 var ProUrl=document.aspnetForm.tribeinfo.value;
  if(ProUrl=="")
  {
    alert("部落简介不能为空！");
	document.aspnetForm.tribeinfo.focus();
	return false;
	}
  return true;  
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>部落管理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="tribe_add.asp">添加部落  </a> <a  href="tribemanage.asp">部落管理</a></td>
  </tr>
</table>

<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
 <form action="?action=add" method="POST" name="aspnetForm" id="aspnetForm" onSubmit="return save_onclick()">

	<tr>
	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">部落名称：</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="tribename" >
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">部落徽章：</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="pic" id="pic">
      </label> &nbsp; 
     <input type="button" name="Submit" value="上传图片"  onClick="window.open('../upload_flash.asp?formname=aspnetForm&editname=pic&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	 </td>
    </tr>
    <tr class="tdbg">
	
      <td class="tdbg"><DIV align="right">激活码：</DIV></td>
      <td class="tdbg"><label>
			<input type="text" name="triberad" >
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">酋长：</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="username" >
      会员名称</label></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><DIV align="right">部落简介：</DIV></td>
      <td width="65%" class="tdbg"><label>
        <textarea name="tribeinfo" cols="40" rows="5"></textarea>
      </label></td>
    </tr>
    
    
    <tr class="tdbg">
      <td colspan="2"><table id="PurviewDetail" width="100%" border="0" cellspacing="10" cellpadding="0" style="display:none">
        <tr>
          <td colspan="2" align="center" class="title"><strong>管 理 员 权 限 详 细 设 置</strong></td>
        </tr>

   
      </table></td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input  type="submit" name="Submit" value=" 添&nbsp;&nbsp;加 " style="cursor: hand;background-color: #cccccc;">&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='tribemanage.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
  </form>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>网络服务平台</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
