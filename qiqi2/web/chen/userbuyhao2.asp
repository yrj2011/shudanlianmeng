
<%
'------------------------------------------------------------------

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
dim  sql, strPurview,ok,id,fid,qq,num,iCount,uid
dim Action,rs,FoundErr,ErrMsg
id=request("id")
uid=request("uid")
ok=request("ok")
if ok="add" then

if request("qq") ="" then
	   Response.Write("<script>alert('绑定买号不能为空!');window.history.back();</script>")
	   response.end()
end if

if request("num")="" then
	   Response.Write("<script>alert('初始买入信誉不能为空!');window.history.back();</script>")
	   response.end()
end if

	Sql = "select * from "&jieducm&"_register where id="&uid&""
	Set rs = Server.CreateObject("Adodb.RecordSet")
	rs.Open Sql,conn,1,1
	IF rs.Eof Then
	   Response.Write("<script>alert('对不起无此用户!');window.history.back();</script>")
	   response.end()
	 else
	 username=rs("username")
	end if
	
	
 	    Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select * From "&jieducm&"_xinyu where shopname='"&request("qq")&"'" ,Conn,3,3  
	   if rs.eof then
	    rs.addnew
		rs("username")=username
		rs("shopname")=request("qq")
		rs("num")=request("num")
		rs("class")=request("classid")	
		rs("start")=0
    	rs.update
    	rs.close
		 Response.Write("<script>alert('绑定成功!');window.location.href='usermanage.asp';</script>")
		 response.End()
	 else
	 	 Response.Write("<script>alert('此号已被其它用户绑定!');history.back();</script>")
         response.End()
	 end if
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
    <td height="22" colspan="2" align="center" class="title"><strong>绑定买号</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="?ok=add" name="form1" onSubmit="javascript:return CheckAdd();">
	<tr>
	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><strong>绑定买号：</strong></td>
      <td class="tdbg"><input name="qq" type="text"></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><strong>初始买入信誉 ：(输入数字)</strong></td>
      <td class="tdbg"><label>
        <input type="text" name="num" onKeyUp="if(isNaN(value))execCommand('undo')">
      </label></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><strong>分类：</strong></td>
      <td width="65%" class="tdbg"><label>
        <select name="classid">
          <option value="1">淘宝区</option>
          <option value="2">拍拍区</option>
          <option value="3">有啊区</option>
        </select>
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
      <td height="40" colspan="2" align="center" class="tdbg"><input type="hidden" name="uid" value="<%=id%>">
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
