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
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
nt2003.GetSite_Setting() 
dim ClassID,SpecialID
dim ClassMaster,BrowsePurview,AddPurview
ClassID=session("ClassID_Article")
SpecialID=session("SpecialID_Article")
if ClassID="" then
	ClassID=0
else
	ClassID=Clng(ClassID)
end if
if SpecialID="" then
	SpecialID=0
else
	SpecialID=Clng(SpecialID)
end if
%>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>七七传媒</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script language = "JavaScript">
function AddItem(strFileName){
  document.myform.IncludePic.checked=true;
  document.myform.DefaultPicUrl.value=strFileName;
  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);
  document.myform.DefaultPicList.selectedIndex+=1;
  if(document.myform.UploadFiles.value==''){
	document.myform.UploadFiles.value=strFileName;
  }
  else{
    document.myform.UploadFiles.value=document.myform.UploadFiles.value+"|"+strFileName;
  }
}
function selectPaginationType()
{
  document.myform.PaginationType.value=2;
}

function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 
  if (document.myform.ClassID.value=="")
  {
    alert("文章所属栏目不能指定为含有子栏目的栏目！");
	document.myform.ClassID.focus();
	return false;
  }
  if (document.myform.ClassID.value==0)
  {
    alert("文章所属栏目不能指定为外部栏目！");
	document.myform.ClassID.focus();
	return false;
  }

  if (document.myform.Title.value=="")
  {
    alert("文章标题不能为空！");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Key.value=="")
  {
    alert("关键字不能为空！");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("文章内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
 if (document.myform.Content.value.length>2048000)
  {
    alert("文章内容太长，超出了ACCESS数据库的限制（2048K）！建议将文章分成几部分录入。");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
function chsel()
	{
		if(document.myform.ClassID.value==29)
		{
			document.all["mrmq"].innerHTML="<table width=100%><tr class=tdbg><td width=112 align=right><strong>人物姓名：</strong></td><td colspan=2 >&nbsp;<input name=mUserName type=text id=mUserName size=20 maxlength=8></td></tr></table>";
		}
		else
		{
			document.all["mrmq"].innerHTML="";
		}
	}
</script>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_ArticleSave.asp" target="_self">
    <tr>
      <td height="22" align="center" class="title" colspan="3"><b>添 加 文 章（简洁模式）</b></td>
    </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td><select name='ClassID' onChange="chsel();"><%call Admin_ShowClass_Option(3,ClassID)%></select>              
            </td>
            <td><%response.write "<font color='#ff6600'><strong>注意：</strong></font><font color='#0066cc'>1、不能指定为含有子栏目的栏目，或者外部栏目"
			if AdminPurview=2 and AdminPurview_Article=3 then
              response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2、你只能在<font color='#ff6600'>红色栏目</font>及其子栏目中发表文章</font>"
			end if%></td>
          </tr>
          <!--<tr class="tdbg"> 
            <td width="100" align="right"><strong>所属专题：</strong></td>
            <td colspan="2"><%call Admin_ShowSpecial_Option(1,SpecialID)%></td>
          </tr>-->
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td colspan="2"><input name="Title" type="text" id="Title" value="" size="50" maxlength="255"> 
              <font color="#ff6600">*</font><select name="TitleFontColor" id="TitleFontColor">
           <option value="" selected>颜色</option>
           <OPTION value="">默认</OPTION>
           <OPTION value="#000000" style="background-color:#000000"></OPTION>
           <OPTION value="#FFFFFF" style="background-color:#FFFFFF"></OPTION>
           <OPTION value="#008000" style="background-color:#008000"></OPTION>
           <OPTION value="#800000" style="background-color:#800000"></OPTION>
           <OPTION value="#808000" style="background-color:#808000"></OPTION>
           <OPTION value="#000080" style="background-color:#000080"></OPTION>
           <OPTION value="#800080" style="background-color:#800080"></OPTION>
           <OPTION value="#808080" style="background-color:#808080"></OPTION>
           <OPTION value="#FFFF00" style="background-color:#FFFF00"></OPTION>
           <OPTION value="#00FF00" style="background-color:#00FF00"></OPTION>
           <OPTION value="#00FFFF" style="background-color:#00FFFF"></OPTION>
           <OPTION value="#FF00FF" style="background-color:#FF00FF"></OPTION>
           <OPTION value="#FF0000" style="background-color:#FF0000"></OPTION>
           <OPTION value="#0000FF" style="background-color:#0000FF"></OPTION>
           <OPTION value="#008080" style="background-color:#008080"></OPTION>
          </select><select name="TitleFontType" id="TitleFontType">
           <option value="0" selected>字形</option>
           <option value="1">粗体</option>
           <option value="2">斜体</option>
           <option value="3">粗+斜</option>
           <option value="0">规则</option>
          </select></td>
          </tr>
          <tr>
		  <td colspan="3" id="mrmq"></td>
		  </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>关键字：</strong></td>
            <td colspan="2"><input name="Key" type="text"
           id="Key" value="<%=session("Key")%>" size="50" maxlength="255"> <font color="#ff6600">*</font><br>
              <font color="#0066cc">用来查找相关文章，可输入多个关键字，中间用<font color="#ff6600">“|”</font>隔开。不能出现&quot;&quot;'*?,.()等字符。</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>作者：</strong></td>
            <td colspan="2"> <input name="Author" type="text"
           id="Author" value="<%=session("Author")%>" size="50" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>转贴自：</strong></td>
            <td colspan="2"> <input name="CopyFrom" type="text"
           id="CopyFrom" value="<%=session("CopyFrom")%>" size="50" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>文章内容：</strong></p>
              <p align="left"><font color="#0066cc"><% if nt2003.site_setting(8)=1 then%>&middot;　如果是从其它网站上复制内容，并且内容中包含有图片，本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度，请稍候（此功能需要服务器安装了IE5.5以上版本才有效）。<%end if%><br>
                <br>
                &middot;　换行请按</font><font color="#FF6600">Shift+Enter</font><br>
                <font color="#0066cc">&middot;　另起一段请按</font><font color="#FF6600">Enter</font><br>
              </p></td>
         <td colspan="2">
		  <textarea name="Content" style="display:none"></textarea> 
          <iframe ID="editor" src="editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe> 
        </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>包含图片：</strong></td>
            <td colspan="2"><input name="IncludePic" type="checkbox" id="IncludePic" value="yes" style="border: 0px;background-color: #eeeeee;">
              是<font color="#0066cc">（如果选中的话会在标题前面显示[图文]）</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>首页图片：</strong></td>
            <td colspan="2"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="56" maxlength="200">
              用于在首页的图片文章处显示 <br>
              直接从上传图片中选择： 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option selected>不指定首页图片</option>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles"> 
            </td>
          </tr>
<%if AdminPurview=1 or AdminPurview_Article<=2 then%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>立即发布：</strong></td>
            <td colspan="2"><input name="Passed" type="checkbox" id="Passed" value="yes" checked style="border: 0px;background-color: #eeeeee;">
              是<font color="#0066cc">（如果选中的话将直接发布，否则审核后才能发布。）</font></td>
          </tr>
<%end if%>
      </td>
    </tr>
    <tr class="tdbg">
     <td align="center" colspan="3">
      <input name="PaginationType" type="hidden" id="PaginationType" value="0"> 
      <input name="Action" type="hidden" id="Action" value="Add1">
      <input
  name="Add" type="submit"  id="Add" value=" 添&nbsp;&nbsp;加 " onClick="document.myform.action='Admin_ArticleSave.asp';document.myform.target='_self';" style="cursor: hand;background-color: #cccccc;">
      &nbsp;&nbsp; 
      <input
  name="Preview" type="submit"  id="Preview" value=" 预&nbsp;&nbsp;览 " onClick="document.myform.action='Admin_ArticlePreview.asp';document.myform.target='_blank';" style="cursor: hand;background-color: #cccccc;">
      &nbsp;&nbsp;
      <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor: hand;background-color: #cccccc;">
 </td></tr>
</form>
 </table>
 <!--#include file="Admin_fooder.asp"-->
</body>
</html>

