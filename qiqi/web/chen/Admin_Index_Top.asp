<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
Const PurviewLevel=0
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<title>��ˢƽ̨ƽ̨--��̨����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
a:link { color:#FFFFFF;text-decoration:none}
a:hover {color:#666666;}
a:visited {color:#000000;text-decoration:none}

td {FONT-SIZE: 9pt;  COLOR: #FFFFFF; FONT-FAMILY: "����"}
img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)} 
</style>
<base target="main">
<script>
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("images/admin_top_open.gif");

var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		parent.frame.cols="0,*";
		displayBar=false;
		obj.src="images/admin_top_open.gif";
		obj.title="����߹���˵�";
	}
	else{
		parent.frame.cols="180,*";
		displayBar=true;
		obj.src="images/admin_top_close.gif";
		obj.title="�ر���߹���˵�";
	}
}
</script>
</head>

<body background="images/admin_top_bg.gif" leftmargin="0" topmargin="0">
<table height="100%" width="100%" border=0 cellpadding=0 cellspacing=0>
<tr valign=middle>
	<td width=50>
	<img onClick="switchBar(this)" src="images/admin_top_close.gif" title="�ر���߹���˵�" style="cursor:hand">
	</td>

	<td width=40>
		<img src="images/admin_top_icon_1.gif" title="�ݶȴ�ý">
	</td>
	<td width=100>
		<a href="Admin_Admin.asp"><font color="#FF6600">�޸�����</font></a>
	</td>
	<td width=40>
		<img src="images/admin_top_icon_5.gif" title="�ݶȴ�ý">
	</td>
	<td width=100>
		<a href="send_zn_message.asp"><font color="#FF6600">�����ʼ�</font></a>
	</td>
	<td width=40>
		<img src="images/admin_top_icon_6.gif" title="�ݶȴ�ý">
	</td>
	<td width=100>
		<a href="inc/ReloadCache.asp"><font color="#FF6600">�޸���Ч</font></a>
	</td>
<!--	<td width=40>
		<img src="images/admin_top_icon_3.gif">
	</td>
	<td width=100>
		<a href="#">�ҵı���¼</a>
	</td>
	<td width=40>
		<img src="images/admin_top_icon_4.gif">
	</td>
	<td width=100>
		<a href="#">�ҵĶ���Ϣ</a>
	</td>
-->	
	<td align="center">&nbsp;&nbsp;</td>
</tr>
</table>
</body>
</html>
