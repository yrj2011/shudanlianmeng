<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/Css.css" type=text/css rel=stylesheet>
<LINK href="../css/top_bottom.css" type=text/css rel=stylesheet>
<SCRIPT src="../js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.f12bt{
	font-size: 12px;
	color: #024E68;
	line-height: 20px;
	text-decoration: none;
	font-weight: bold;
}
-->
</style>

</head>

<body>  
<!--#include file="../jieducm_top.asp"-->
<table width="965" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF" height="5"></td>
  </tr>
</table>

<div align="center">
<TABLE cellSpacing=0 cellPadding=0 width=960 align=center border=0>
  

  <tr>
    <td valign="top" align="left" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="28"><a class="sub-step-on" href="01.htm">��1��</a> <a href="../02.htm">��2��</a> <a href="../03.htm">��3��</a> <a href="../04.htm">��4��</a> <a href="../05.htm">��5��</a> <a href="../06.htm">��6��</a> <a href="../07.htm">��7��</a> <a href="../08.htm">��8��</a> <a href="../09.htm">��9��</a> <a href="../10.htm">��10��</a> <a href="../11.htm">��11��</a></td>
      </tr>
      <tr>
        <td height="51"><div align="center"><img src="2.gif" alt="�����������" width="900" height="450" border="0" usemap="#Map" />
            <map name="Map" id="Map">
              <area shape="rect" coords="396,90,489,128" href="../02.htm" />
            </map>
        </div></td>
      </tr>
    </table></td>
  </tr>
  
 
</table>
</div>
<DIV class="page paddingt5">
  <div align="center">
    <SCRIPT src="../../images/foot_scroll.js"></SCRIPT>
  </div>
</DIV>
<!--#include file="../jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>