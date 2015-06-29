<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<SCRIPT src="../js/jieducm_zDialog.js" type=text/javascript></SCRIPT>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>

<%
pid=Replace(Trim(request("taobaoid")),"'","''")
Sql = "select * from "&jieducm&"_zhongxin where  pid='"&pid&"' and IfComeSet>0"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
IfComeSet=rs("IfComeSet")
ComeKey	=rs("ComeKey")
ComeNote=rs("ComeNote")
prourl=rs("prourl")
pidy=mid(prourl,instr(1,prourl,"id=")+3,len(prourl)-instr(1,prourl,"id=")-2)
Shopkeeper=rs("Shopkeeper")
else
 Response.Write("<script>alert('出错了请返回!');history.back();</script>")
 response.End()
end if
 
Sql= "select * from "&jieducm&"_jieshou where pid='"&pid&"'  and username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then			
 Response.Write("<script>alert('对不起,操作有误!');history.back();</script>")
 response.End()
end if

if request.Form("action")="jieducm_ok" then
prourl_jieducm=Replace(Trim(Request.Form("prourl_jieducm")),"'","''")
if prourl_jieducm="" then
 Response.Write("<script>alert('请输入要验证的商品链接id!');history.back();</script>")
 response.End()
end if
 
'if prourl_jieducm<>prourl then

if prourl_jieducm<>pidy then
 Response.Write("<script>alert('" & prourl_jieducm & "验证失败！');history.back();</script>")
 response.End()
else
 Response.Write("<script>alert('验证成功!此商品地址是发布方发布的商品地址，请放心购买！');history.back();</script>")
 response.End()
end if
end if
jieducm_taobao="http://www.taobao.com"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE>七七网</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
 <LINK href="../css/top_bottom.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.STYLE1 {font-size: 14px}
.STYLE2 {font-size: 14px}
-->
</style>

<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
	color: #333;
}
body {
	margin-left: 20px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
}
-->
</style>
</HEAD>
<BODY>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 
      borderColor=#dddddd cellPadding=3 width=100% align=center>
  <tr height="30">
    <td colspan="3" bgcolor="#CCCCCC" height="30"><div align="center" class="STYLE1">该任务属于商品来路任务，接收方需要完成下列的来路步骤后方能查看到商品的网址</div></td>
  </tr>
  <tr>
    <td colspan="3" height="20"><div align="left" class="STYLE2"><font color="#CC9900"><strong>来路商品详情</strong></font></div></td>
  </tr>
  <tr height="30">
    <td width="12%" height="30"><font color="#669900"><strong>第一步：</strong></font></td>
    <td height="30" colspan="2">
	  <div align="left">
	 <%if IfComeSet=1 then%>
	请在淘宝首页搜索店铺：
	<%elseif IfComeSet=2 then%>
	请在淘宝首页搜索商品：
	<%elseif IfComeSet=3 then%>
	打开该卖家信誉评价地址：
	<%end if%>
	<%=ComeKey%>	  </div></td>
  </tr>
  <tr height="30">
    <td height="30">&nbsp;</td>
    <td height="30" colspan="2"><a onClick="javascript:clipboardData.setData('text','<%=ComeKey%>');alert('复制成功');return false;" href="javascript:void(0)"><img src="../images/key_lai.jpg" width="84" height="18" border="0"></a>&nbsp;
	<a  onClick="jsWindon('<%=jieducm_taobao%>')" href="javascript:void(0)">
	 <img src="../images/taobao_in.jpg" width="96" height="19" border="0"></a></td>
  </tr>
  <tr height="30">
    <td height="30"><font color="#669900"><strong>第二步：</strong></font></td>
    <td height="30" colspan="2">对照提示信息：<%=ComeNote%></td>
  </tr>
  <tr height="30">
    <td height="30"><font color="#669900"><strong>第三步：</strong></font></td>
    <td height="30" colspan="2">复制商品页面地址栏的链接id，并粘贴到下面，然后点击【验证商品】按钮</td>
  </tr>
  <tr>
    <td height="30"><font color="#669900"><strong>掌柜名：</strong></font></td>
    <td colspan="2"><%=Shopkeeper%></td>
  </tr>
  <tr>
    <td height="30"><font color="#669900"><strong>链接id：</strong></font></td>
    <td colspan="2"><form name="form1" method="post" action="">
	<input name="action" type="hidden" value="jieducm_ok">
	<input name="taobaoid" type="hidden" value="<%=pid%>">
      <label>
        <input type="text" name="prourl_jieducm" style="WIDTH: 260px">
        </label>
      <label>
      <input type="submit" name="Submit" value="验证商品">
      </label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="/news.asp?/1482.html" target="_blank"><font color="#FF0000"><strong>什么是链接id？</strong></font></a>
    </form>    </td>
  </tr>
  
  <tr>
    <td colspan="3"><div align="center"> <a href="javascript:void(0)" onClick="Dialog.close()"><img src="../images/close.jpg" border="0"></a></div></td>
  </tr>
</table>
<p>
  <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</p>
</BODY></HTML>
