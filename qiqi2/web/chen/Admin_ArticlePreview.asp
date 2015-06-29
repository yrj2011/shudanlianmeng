<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
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
dim FoundErr,ErrMsg
dim ClassID,tClass,ClassName,RootID,ParentID,ParentPath,Depth,Child,ClassMaster
call main()
'call CloseConn()

sub main()
%>
<html>
<head>
<script>
function zoom_img(e, o)
{
var zoom = parseInt(o.style.zoom, 10) || 100;
zoom += event.wheelDelta / 12;
if (zoom > 0) o.style.zoom = zoom + '%';
return false;
}
</script>
<title><%=request("Title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="760" border="0" align="center" cellpadding="5" cellspacing="1" class="border">
  <tr class="title"> 
    <td width="400" height="22">
	<%
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定所属栏目</li>"
		exit sub
	else
		ClassID=Clng(ClassID)
	end if
	set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassMaster From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
		exit sub
	else
		ClassName=tClass(0)
		RootID=tClass(1)
		ParentID=tClass(2)
		Depth=tClass(3)
		ParentPath=tClass(4)
		Child=tClass(5)
		ClassMaster=tClass(6)
	end if
	set tClass=nothing
if ParentID>0 then
	dim sqlPath,rsPath
	sqlPath="select ClassID,ClassName From "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ") order by Depth"
	set rsPath=server.createobject("adodb.recordset")
	rsPath.open sqlPath,conn,1,1
	do while not rsPath.eof
		response.Write rsPath(1) & "&nbsp;&gt;&gt;&nbsp;"
		rsPath.movenext
	loop
	rsPath.close
	set rsPath=nothing
end if
response.write ClassName & "&nbsp;&gt;&gt;&nbsp;"

	if request("IncludePic")=true then
		response.write "<font color=blue>[图文]</font>"
	end if
	response.write "<font color='ff6600'>"
	response.write request("Title")
	response.write "</font>"
	%>
	</td>
    <td width="50" height="22" align="center"> <%
			if lcase(request("OnTop"))="yes" then
				response.Write("<font color=blue>顶&nbsp;</font>")
			else
				response.write("&nbsp;&nbsp;&nbsp;")
			end if
			if lcase(request("Hot"))="yes" then
				response.write("<font color='#0066cc'>热&nbsp;</font>")
			else
				response.write("&nbsp;&nbsp;&nbsp;")
			end if
			if lcase(request("Elite"))="yes" then
				response.write("<font color=green>荐</font>")
			else
				response.write("&nbsp;&nbsp;")
			end if
			%> </td>
  </tr>
  
  <tr class="tdbg"> 
    <td colspan="3"><p align="center"><font size="6" face="黑体"><%=dvhtmlencode(request("Title"))%></font></p>
   <% if  request("Title1")<>"" then %>
      <div align="center"></div>
      <p align="center"><font size="4"><%=dvhtmlencode(request("Title1"))%></font>

  <% end if %>
<br>

        作者：<%=dvhtmlencode(request("Author"))%>&nbsp;&nbsp;&nbsp;&nbsp;转贴自：<%=dvhtmlencode(request("CopyFrom"))%>&nbsp;&nbsp;&nbsp;&nbsp;点击数：0&nbsp;&nbsp;&nbsp;&nbsp;文章录入：<%=session("AdminName")%></p>
      <p><%=ubbcode(request("Content"))%></p>
      </td>
  </tr>
<tr class="tdbg"><td align="center" colspan="2">
【<a href="javascript:window.close();">关闭窗口</a>】
</td></tr>
</table>
<!--#include file="Admin_fooder.asp"-->
</body>
</html>
<%
end sub
%>