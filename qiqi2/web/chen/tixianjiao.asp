<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
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
id =request("id")
Sql3 = "select * from "&jieducm&"_tixian where id="&id&""
Set Rs3 = Server.CreateObject("Adodb.RecordSet")
Rs3.Open Sql3,conn,3,3
if not rs3.eof then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>完成提现--输入支付宝交易号</title>
   <SCRIPT language=javascript>
function  save_onclick22()
{

    var Price=document.form1.zfbnum.value;
  if(Price=="")
  {
    alert("请输入支付宝交易号！");
	document.form1.zfbnum.focus();
	return false;
	}
  

  return true;  
}
</script>  
</head>

<body>
<form id="form1" name="form1" method="post" action="tixianok.asp?id=<%=rs3("id")%>&action=ok" onSubmit="return save_onclick22();">
  <table width="336" border="0">
    <tr>
      <td width="125">提现流水号：</td>
      <td width="201"><%=rs3("pid")%></td>
    </tr>
    <tr>
      <td>提现用户名：</td>
      <td><%=rs3("username")%></td>
    </tr>
    <tr>
      <td>接收支付宝号：</td>
      <td><%=rs3("ReZfb")%></td>
    </tr>
    <tr>
      <td>真实姓名：</td>
      <td><%=rs3("ReRName")%></td>
    </tr>
    <tr>
      <td>交易号：</td>
      <td><input type="text" name="zfbnum" id="zfbnum" /></td>
    </tr>
    <tr>
      <td><input type="submit" name="button" id="button" value="提交" /></td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
<% end if%>