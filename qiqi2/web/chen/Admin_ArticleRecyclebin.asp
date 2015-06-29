<%@language=vbscript codepage=936 %>
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=1
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim strFileName,FileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim i,j
dim keyword,strField
dim sql,rsArticleList
dim ClassID
dim strAdmin,arrAdmin
dim tClass,ClassName,RootID,ParentID,Depth,ParentPath,Child,ChildID,tID,tChild,ClassMaster
dim ManageType
ManageType=trim(request("ManageType"))
FileName="Admin_ArticleRecyclebin.asp"
ClassID=Trim(request("ClassID"))
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
end if
strField=trim(request("Field"))

if ClassID="" then
	ClassID=0
else
	ClassID=CLng(ClassID)
end if
if ClassID>0 then
	set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassMaster From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
		Call WriteErrMsg()
		response.end
	else
		ClassName=tClass(0)
		RootID=tClass(1)
		ParentID=tClass(2)
		Depth=tClass(3)
		ParentPath=tClass(4)
		Child=tClass(5)
	end if
end if
strFileName=FileName & "?ClassID=" & ClassID & "&strField=" & strField & "&keyword=" & keyword
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
%>
<html>
<head>
<title>捷-度-传-媒</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel()
{
	if(document.myform.Action.value=="ConfirmDel")
	{
		if(confirm("确定要彻底删除选中的文章吗？一旦删除将不能恢复！"))
			return true;
		else
			return false;
	}
	else if(document.myform.Action.value=="ClearRecyclebin")
	{
		if(confirm("确定要清空回收站？一旦清空将不能恢复！"))
			return true;
		else
			return false;
	}
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong>回 收 站 管 理</strong></td>
  </tr>
  <tr valign="middle" class="tdbg">
	<td width="80" align="right"><strong>管理导航：</strong></td>
    <td><a href="Admin_ArticleRecyclebin.asp">回收站管理首页</a> | <a href="Admin_ArticleDel.asp?Action=ClearRecyclebin" onClick="return confirm('确定要清空回收站吗？一旦清空将无法恢复！');">清空回收站</a> 
      | <a href="Admin_ArticleDel.asp?Action=RestoreAll" onClick="return confirm('确定还原回收站中的所有文章？');">还原回收站中的所有文章</a></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="left" class="title"><%call Admin_ShowRootClass()%></td>
  </tr>
  <tr valign="middle" class="tdbg">
    <td><%call Admin_ShowPath("回收站管理")%></td>
    <td width="200" height="22" align="right">
	<select name='JumpClass' id="JumpClass" onChange="if(this.options[this.selectedIndex].value!=''){location='<%=FileName & "?ClassID="%>'+this.options[this.selectedIndex].value;}">
      <option value='' selected>跳转栏目至…</option>
	  <%call Admin_ShowClass_Option(2,0)%>
	</select>
    </td>
  </tr>
<%
'sql="select Article.ArticleID,Article.ClassID,ArticleClass.ClassName,Article.Title,Article.Key1,Article.Author,Article.CopyFrom,Article.UpdateTime,Article.Editor,"
sql=sql & "select *  from   "&jieducm&"_article"
sql=sql & " inner join "&jieducm&"_ArticleClass on "&jieducm&"_Article.ClassID="&jieducm&"_ArticleClass.ClassID where "&jieducm&"_Article.Deleted=1 "
if ClassID>0 then
	if Child>0 then
		ChildID=""
		set tClass=nt2003.execute("select ClassID from "&jieducm&"_ArticleClass where ParentPath like '" & ParentPath & "," & ClassID & "%'")
		do while not tClass.eof
			if ChildID="" then
				ChildID=tClass(0)
			else
				ChildID=ChildID & "," & tClass(0)
			end if
			tClass.movenext
		loop
		sql=sql & " and "&jieducm&"_Article.ClassID in (" & ChildID & ")"
	else
		sql=sql & " and "&jieducm&"_Article.ClassID=" & ClassID
	end if
end if
if keyword<>"" then
	select case strField
		case "Title"
			sql=sql & " and Title like '%" & keyword & "%' "
		case "Content"
			sql=sql & " and Content like '%" & keyword & "%' "
		case "Author"
			sql=sql & " and Author like '%" & keyword & "%' "
		case "Editor"
			sql=sql & " and Editor like '%" & keyword & "%' "
		case else
			sql=sql & " and Title like '%" & keyword & "%' "
	end select
end if
sql=sql & " order by "&jieducm&"_Article.articleid desc"

Set rsArticleList= Server.CreateObject("ADODB.Recordset")
rsArticleList.open sql,conn,1,1
if rsArticleList.eof and rsArticleList.bof then
	totalPut=0
	if Child=0 then
		response.write "<tr class='tdbg'><td align='center' colspan='2'><br>没有任何被删除的文章！<br></td></tr></table>"
	else
		response.write "<tr class='tdbg'><td align='center' colspan='2'><br>此栏目的下一级子栏目中没有任何被删除的文章！<br></td></tr></table>"
	end if
else
   	totalPut=rsArticleList.recordcount
	if currentpage<1 then
   		currentpage=1
   	end if
   	if (currentpage-1)*MaxPerPage>totalput then
   		if (totalPut mod MaxPerPage)=0 then
     		currentpage= totalPut \ MaxPerPage
	  	else
	      	currentpage= totalPut \ MaxPerPage + 1
   		end if
   	end if
    if currentPage=1 then
       	showContent
       	response.write "<table class=border border=0 cellspacing=1 width=100% cellpadding=0><tr valign=middle class=tdbg><td>"
       	showpage strFileName,totalput,MaxPerPage,true,true,"篇文章"
       	response.write "</td></tr></table>"
 	else
     	if (currentPage-1)*MaxPerPage<totalPut then
       	   	rsArticleList.move  (currentPage-1)*MaxPerPage
       		dim bookmark
           	bookmark=rsArticleList.bookmark
            showContent
       	response.write "<table class=border border=0 cellspacing=1 width=100% cellpadding=0><tr valign=middle class=tdbg><td>"
       	showpage strFileName,totalput,MaxPerPage,true,true,"篇文章"
       	response.write "</td></tr></table>"
       	else
	        currentPage=1
           	showContent
       	response.write "<table class=border border=0 cellspacing=1 width=100% cellpadding=0><tr valign=middle class=tdbg><td>"
       	showpage strFileName,totalput,MaxPerPage,true,true,"篇文章"
       	response.write "</td></tr></table>"
	    end if
	end if
end if
rsArticleList.close
set rsArticleList=nothing  


sub showContent
   	dim ArticleNum
    ArticleNum=0
%>
<table width='100%' border="0" cellpadding="0" cellspacing="0"><tr>
    <form name="myform" method="Post" action="Admin_ArticleDel.asp" onSubmit="return ConfirmDel();">
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22"> 
            <td height="22" width="30" align="center"><strong>选中</strong></td>
            <td width="25" align="center"  height="22"><strong>ID</strong></td>
            <td align="center" ><strong>文章标题</strong></td>
            <td width="60" align="center" ><strong>录入</strong></td>
            <td width="40" align="center" ><strong>点击数</strong></td>
            <td width="60" align="center" ><strong>文章属性</strong></td>
            <td width="40" align="center" ><strong>已审核</strong></td>
            <td width="100" align="center" ><strong>操作</strong></td>
          </tr>
          <%do while not rsArticleList.eof%>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
            <td width="30" align="center"><input name='ArticleID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=cstr(rsArticleList("articleID"))%>' style="border: 0px;background-color: #eeeeee;"></td>
            <td width="25" align="center"><%=rsArticleList("articleid")%></td>
            <td> <%
			if rsArticleList("ClassID")<>ClassID then
				response.write "<a href='" & FileName & "?ClassID=" & rsArticleList("ClassID") & "'>[" & rsArticleList("ClassName") & "]</a>&nbsp;"
			end if
			if rsArticleList("IncludePic")=true then
				response.write "<font color=blue>[图文]</font>"
			end if
			response.write rsArticleList("title")
			%></td>
            <td width="60" align="center"><%= rsArticleList("Editor") %></td>
            <td width="40" align="center"><%= rsArticleList("Hits") %></td>
            <td width="60" align="center"> <%
			if rsArticleList("OnTop")=true then
				response.Write "<font color=blue>顶</font> "
			else
				response.write "&nbsp;&nbsp;&nbsp;"
			end if
			if rsArticleList("Hits")>=nt2003.site_setting(14) then
				response.write "<font color=red>热</a> "
			else
				response.write "&nbsp;&nbsp;&nbsp;"
			end if
			if rsArticleList("Elite")=true then
				response.write "<font color=green>荐</a>"
			else
				response.write "&nbsp;&nbsp;"
			end if
			%> </td>
            <td width="40" align="center"> <%
			if rsArticleList("Passed")=true then
				response.write "是"
			else
				response.write "否"
			end if%></td>
            <td width="100" align="center"><a href="Admin_ArticleDel.asp?Action=ConfirmDel&ArticleID=<%=rsArticleList("ArticleID")%>" onclick='return ConfirmDel();'>彻底删除</a> 
              <a href="Admin_ArticleDel.asp?Action=Restore&ArticleID=<%=rsArticleList("ArticleID")%>">还原</a></td>
          </tr>
          <%
		ArticleNum=ArticleNum+1
	   	if ArticleNum>=MaxPerPage then exit do
	   	rsArticleList.movenext
	loop
%>
        </table>
		<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
		  <tr valign="middle" class="tdbg">
			<td width="210">
			 <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">选中本页显示的所有文章 
			</td>
            <td width="700">
			  <input name="Action" type="hidden" id="Action2" value="ConfirmDel">
              <input name="submit1" type='submit' id="submit1" onClick="document.myform.Action.value='ConfirmDel'" value='&nbsp;彻底删除选定的文章&nbsp;' style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="submit" name="Submit2" value="&nbsp;清空回收站&nbsp;"  onClick="document.myform.Action.value='ClearRecyclebin'" style="cursor: hand;background-color: #cccccc;">
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input name="Submit3" type="submit" id="Submit3"  onClick="document.myform.Action.value='Restore'" value="&nbsp;还原选定的文章&nbsp;" style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;&nbsp;&nbsp;
              <input name="Submit4" type="submit" id="Submit4" onClick="document.myform.Action.value='RestoreAll'" value="&nbsp;还原所有文章&nbsp;" style="cursor: hand;background-color: #cccccc;"></td>
  </tr>
</table>
</td>
</form></tr></table>
<%
end sub

if ClassID>0 and Child>0 then
%>
<table width="100%" height="5" border="0" cellpadding="0" cellspacing="0"><tr><td></td></tr></table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class='border'>
  <tr height="20" class='tdbg'>
    <td width='150' align="right">【<%response.write "<a href='" & strFileName & "'>" & ClassName & "</a>"%>】子栏目导航：</td>
	<td><%call Admin_ShowChild()%></td></tr>
</table>
<%
end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td width="80" align="right"><strong>文章搜索：</strong></td>
    <td>
      <%call Admin_ShowSearchForm(FileName,2)%>
    </td>
  </tr>
</table>
<!--#include file="Admin_fooder.asp"-->
</body>
</html>
<%
''call CloseConn() 'shiyu
%>