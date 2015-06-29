
<%
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
dim rs, sql, strPurview,ok,id,fid,qq,num,iCount,class1
dim Action,rsr,FoundErr,ErrMsg
Action=Trim(request("Action"))
ok=request("ok")
fid=request("fid")
qq=request("qq")
num=request("num")
class1=request("class1")

if ok="add" then
 	 Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_qq " ,Conn,3,3  
	    rsr.addnew
		rsr("qq")=request.form("qq")
		rsr("num")=request.form("num")
		rsr("class")=request.Form("class")
    	rsr.update
    	rsr.close
		response.Redirect("qqadmin.asp")
elseif ok="mod" then
     id=request("id")
 	 Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_qq where id="&id&"" ,Conn,3,3  
		rsr("qq")=request.form("qq")
		rsr("num")=request.form("num")
		rsr("class")=request.Form("class")
    	rsr.update
    	rsr.close
		response.Redirect("qqadmin.asp")
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
    <td height="22" colspan="2" align="center" class="title"><strong>群 号 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<%if action="add" then%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="qqadminadd.asp?ok=add" name="form1" onSubmit="javascript:return CheckAdd();">
	<tr>
	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><strong>QQ 号：</strong></td>
      <td class="tdbg"><input name="qq" type="text"></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><strong>类型：</strong></td>
      <td class="tdbg"><input type="radio" name="class" value="1">
      指导  
        <input type="radio" name="class" value="2">
      财务</td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> QQ简介：</strong></td>
      <td width="65%" class="tdbg">
      
      <textarea name="num" id="num" cols="45" rows="5"></textarea></td>
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
		<input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='qqadmin.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
  </form>
</table>
<%elseif action="mod" then%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="qqadminadd.asp?ok=mod" name="form1" onSubmit="javascript:return CheckModifyPwd();">
	<tr>
	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    
    <tr>
      <td class="tdbg"><strong>QQ 号：</strong></td>
      <td class="tdbg"><input name="qq" type="text" value="<%=qq%>"></td>
    </tr>
    <tr>
      <td class="tdbg"><strong>类型：</strong></td>
      <td class="tdbg"><input type="radio" name="class" value="1" <%if class1="1" then%> checked="checked" <%end if%>>
指导
  <input type="radio" name="class" value="2"  <%if class1="2" then%> checked="checked" <%end if%>>
财务</td>
    </tr>
    <tr> 
      <td width="16%" class="tdbg"><strong>QQ简介：</strong></td>
      <td width="84%" class="tdbg"> <textarea name="num" id="num" cols="45" rows="5"><%=num%></textarea></td>
    </tr>
    
    <tr> 
      <td colspan="2" align="center" class="tdbg">
        <input type="hidden" name="id" id="id" value="<%=fid%>">
        <input  type="submit" name="Submit" value="&nbsp;保存修改结果&nbsp;" style="cursor: hand;background-color: #cccccc;">&nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='qqadmin.asp'" style="cursor: hand;background-color: #cccccc;">	  </td>
    </tr>
</form>
</table>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com/' target="_blank"><font color=#ff6600><strong>七七网络服务平台</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
