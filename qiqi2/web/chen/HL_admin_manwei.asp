<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
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
if session("AdminName")="" then
   response.redirect("admin_login.asp")
end if%>
<%
action=request("action")
id=request("id")
if action<>"" then
Set rs = Server.CreateObject("ADODB.Recordset")
sql="delete from "&jieducm&"_ggwei where hl_id="&id&""
rs.open sql,conn


Set rsd = Server.CreateObject("ADODB.Recordset")
sqld="delete from "&jieducm&"_advertising where hl_ggwei="&id&""
rsd.open sqld,conn

%>
<script language="javascript">
alert("信息删除成功!");
location.href("Hl_admin_manwei.asp");
</script>
<%

set rs=nothing

set rsd=nothing
end if
%>


<style type="text/css">
<!--
body {
	background-color: #CED7F7;
}
body,td,th {
	font-size: 12px;
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
-->
</style>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>

<link href="inc/url.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE3 {color: #FF3300}
-->
</style>
</head>

<body>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" >
    <tr>
      <td colspan="3" bgcolor="#A6B5F0"><table cellpadding="2" cellspacing="1" border="0" width="100%" align="center" class="a2">
          <tr>
            <td height="25" colspan="3" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><b>广告位管理</b></td>
          </tr>
          <tr>
            <td height="25" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center">广告位ID</div></td>
            <td background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center">广告位名称</div></td>
            <td background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center">管理</div></td>
          </tr>
		  <%
		 set rs1=server.createobject("adodb.recordset")
		rs1.open "select * from "&jieducm&"_ggwei where hl_lei=0  order by hl_ID Desc",conn,1,1
		if rs1.recordcount=0 then 
		%>
          <tr>
            <td height="25" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center"></div></td>
            <td background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center"></div></td>
            <td background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center"></div></td>
          </tr>   <%else
		  do while not rs1.eof  
		 %>
          <tr>
            <td width="8%" height="25" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center"><%=rs1("hl_id")%></div></td>
            <td width="46%" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center"><%=rs1("hl_title")%></div></td>
            <td width="46%" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><div align="center"><a href="hl_admin_manwei.asp?action=yes&id=<%=rs1("hl_id")%>">删除</a></div></td>
          </tr>
           <%
		 rs1.movenext
		 loop
		 end if%>
      </table></td>
    </tr>
</table>

</body>
</html>
