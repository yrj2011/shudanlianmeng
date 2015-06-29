<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>第1步 -注册教程_<%=stupname%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<link rel=stylesheet type=text/css href="css/global.css">
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>

<style type="text/css">
<!--
.STYLE2 {
	font-size: 16px;
	color: #FF0000;
}
-->
</style>
</head>

<body>  
<%
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 1 num from "&jieducm&"_jieshou where username='"&session("useridname")&"' and classid='1'  order by  id desc",conn,1,1
if not (rs.eof) then
  checknum=rs("num")
rs.close
end if
%>
<!--#include file="jieducm_top.asp"-->
<table width="965" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF" height="5"></td>
  </tr>
</table>

<table width="960" border="5" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666" class="index_tt">
  <tr>
    <td valign="top" align="left" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="28" bgcolor="#CCFFFF"><span class="STYLE2"><a class="sub-step-on" href="zcjc01.asp">第1步</a> >> <a href="zcjc02.asp">第2步</a> >> <a href="zcjc03.asp">第3步</a> >> <a href="zcjc04.asp">第4步</a> >> <a href="zcjc05.asp">第5步</a></span></td>
      </tr>
      <tr>
        <td height="51"><div align="center"><img src="images/teach/curse1/1.gif" alt="注册帮助" border="0" usemap="#Map" /><map name="Map" id="Map">
              <area shape="rect" coords="753,194,827,219" href="zcjc02.asp" />
            </map></div></td>
      </tr>
	   <tr>
        <td height="51"><div align="center"><a href="/zcjc02.asp"><img src="images/teach/xyb.gif" width="115" height="30" border="0" ></a></div></td>
      </tr>
    </table></td>
  </tr>
</table>
</DIV>
<!--#include file="jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>
