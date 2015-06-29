 
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim  sql, strPurview,ok,id,fid,qq,num,iCount,uid
username=request("username")
action=request.Form("action")
uid=request.Form("uid")
nums=request.Form("num")
if action="ok" then
if nums="" then
   Response.Write("<script>alert('天数不能为空！');history.back();</script>")
   response.End()
end if
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where   username='"&uid&"'" ,Conn,3,3  
	  if not rs.eof then  
    	rs("vip")=1
		rs("vipdel")=now()
		rs("vipdate")=nums
    	rs.update
    	rs.close
	  end if
	  
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")="后台设置vip"
   rs("content")=session("AdminName")&"后台设置VIP会员:"&uid&"vip有效期是:"&nums&"天"
   rs("jiegou")="设置成功"
   rs.update
   rs.close
   
	   num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	  Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=uid
    	rs("num")=num
		rs("class")="申请VIP用户"
		rs("content")=session("AdminName")&"后台设置VIP会员:"&uid&"vip有效期是:"&nums&"天"
		rs("price")=0
		rs("jiegou")="设置成功"
    	rs.update
    	rs.close
		
 Response.Write("<script>alert('设置成功!');window.location.href='usermanage.asp';</script>")	
response.End()	
		
end if

%>
<html>
<head>
<title>管理员管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>设置vip会员</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="" name="form1" onSubmit="javascript:return CheckAdd();">
	<tr>
	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right"><strong>会员名：</strong></div></td>
      <td width="65%" class="tdbg"><%=username%></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><div align="right"><strong>vip使用天数 ： </strong></div></td>
      <td class="tdbg"><label>
        <input type="text" name="num" onKeyUp="if(isNaN(value))execCommand('undo')">
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
      <td height="40" colspan="2" align="center" class="tdbg"><input type="hidden" name="action" value="ok">
	  <input type="hidden" name="uid" value="<%=username%>">
      <input  type="submit" name="Submit" value=" 添&nbsp;&nbsp;加 " style="cursor: hand;background-color: #cccccc;">&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='usermanage.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
  </form>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>七七网络服务平台</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
