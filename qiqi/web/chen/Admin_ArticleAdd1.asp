<%@language=vbscript codepage=936 %>
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
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
<title>���ߴ�ý</title>
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
    alert("����������Ŀ����ָ��Ϊ��������Ŀ����Ŀ��");
	document.myform.ClassID.focus();
	return false;
  }
  if (document.myform.ClassID.value==0)
  {
    alert("����������Ŀ����ָ��Ϊ�ⲿ��Ŀ��");
	document.myform.ClassID.focus();
	return false;
  }

  if (document.myform.Title.value=="")
  {
    alert("���±��ⲻ��Ϊ�գ�");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Key.value=="")
  {
    alert("�ؼ��ֲ���Ϊ�գ�");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("�������ݲ���Ϊ�գ�");
	editor.HtmlEdit.focus();
	return false;
  }
 if (document.myform.Content.value.length>2048000)
  {
    alert("��������̫����������ACCESS���ݿ�����ƣ�2048K�������齫���·ֳɼ�����¼�롣");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
function chsel()
	{
		if(document.myform.ClassID.value==29)
		{
			document.all["mrmq"].innerHTML="<table width=100%><tr class=tdbg><td width=112 align=right><strong>����������</strong></td><td colspan=2 >&nbsp;<input name=mUserName type=text id=mUserName size=20 maxlength=8></td></tr></table>";
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
      <td height="22" align="center" class="title" colspan="3"><b>�� �� �� �£����ģʽ��</b></td>
    </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>������Ŀ��</strong></td>
            <td><select name='ClassID' onChange="chsel();"><%call Admin_ShowClass_Option(3,ClassID)%></select>              
            </td>
            <td><%response.write "<font color='#ff6600'><strong>ע�⣺</strong></font><font color='#0066cc'>1������ָ��Ϊ��������Ŀ����Ŀ�������ⲿ��Ŀ"
			if AdminPurview=2 and AdminPurview_Article=3 then
              response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2����ֻ����<font color='#ff6600'>��ɫ��Ŀ</font>��������Ŀ�з�������</font>"
			end if%></td>
          </tr>
          <!--<tr class="tdbg"> 
            <td width="100" align="right"><strong>����ר�⣺</strong></td>
            <td colspan="2"><%call Admin_ShowSpecial_Option(1,SpecialID)%></td>
          </tr>-->
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>���±��⣺</strong></td>
            <td colspan="2"><input name="Title" type="text" id="Title" value="" size="50" maxlength="255"> 
              <font color="#ff6600">*</font><select name="TitleFontColor" id="TitleFontColor">
           <option value="" selected>��ɫ</option>
           <OPTION value="">Ĭ��</OPTION>
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
           <option value="0" selected>����</option>
           <option value="1">����</option>
           <option value="2">б��</option>
           <option value="3">��+б</option>
           <option value="0">����</option>
          </select></td>
          </tr>
          <tr>
		  <td colspan="3" id="mrmq"></td>
		  </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>�ؼ��֣�</strong></td>
            <td colspan="2"><input name="Key" type="text"
           id="Key" value="<%=session("Key")%>" size="50" maxlength="255"> <font color="#ff6600">*</font><br>
              <font color="#0066cc">��������������£����������ؼ��֣��м���<font color="#ff6600">��|��</font>���������ܳ���&quot;&quot;'*?,.()���ַ���</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>���ߣ�</strong></td>
            <td colspan="2"> <input name="Author" type="text"
           id="Author" value="<%=session("Author")%>" size="50" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>ת���ԣ�</strong></td>
            <td colspan="2"> <input name="CopyFrom" type="text"
           id="CopyFrom" value="<%=session("CopyFrom")%>" size="50" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>�������ݣ�</strong></p>
              <p align="left"><font color="#0066cc"><% if nt2003.site_setting(8)=1 then%>&middot;������Ǵ�������վ�ϸ������ݣ����������а�����ͼƬ����ϵͳ�����ͼƬ���Ƶ���վ�������ϣ�ϵͳ������ͼƬ�Ĵ�С��Ӱ���ٶȣ����Ժ򣨴˹�����Ҫ��������װ��IE5.5���ϰ汾����Ч����<%end if%><br>
                <br>
                &middot;�������밴</font><font color="#FF6600">Shift+Enter</font><br>
                <font color="#0066cc">&middot;������һ���밴</font><font color="#FF6600">Enter</font><br>
              </p></td>
         <td colspan="2">
		  <textarea name="Content" style="display:none"></textarea> 
          <iframe ID="editor" src="editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe> 
        </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>����ͼƬ��</strong></td>
            <td colspan="2"><input name="IncludePic" type="checkbox" id="IncludePic" value="yes" style="border: 0px;background-color: #eeeeee;">
              ��<font color="#0066cc">�����ѡ�еĻ����ڱ���ǰ����ʾ[ͼ��]��</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>��ҳͼƬ��</strong></td>
            <td colspan="2"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="56" maxlength="200">
              ��������ҳ��ͼƬ���´���ʾ <br>
              ֱ�Ӵ��ϴ�ͼƬ��ѡ�� 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option selected>��ָ����ҳͼƬ</option>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles"> 
            </td>
          </tr>
<%if AdminPurview=1 or AdminPurview_Article<=2 then%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>����������</strong></td>
            <td colspan="2"><input name="Passed" type="checkbox" id="Passed" value="yes" checked style="border: 0px;background-color: #eeeeee;">
              ��<font color="#0066cc">�����ѡ�еĻ���ֱ�ӷ�����������˺���ܷ�������</font></td>
          </tr>
<%end if%>
      </td>
    </tr>
    <tr class="tdbg">
     <td align="center" colspan="3">
      <input name="PaginationType" type="hidden" id="PaginationType" value="0"> 
      <input name="Action" type="hidden" id="Action" value="Add1">
      <input
  name="Add" type="submit"  id="Add" value=" ��&nbsp;&nbsp;�� " onClick="document.myform.action='Admin_ArticleSave.asp';document.myform.target='_self';" style="cursor: hand;background-color: #cccccc;">
      &nbsp;&nbsp; 
      <input
  name="Preview" type="submit"  id="Preview" value=" Ԥ&nbsp;&nbsp;�� " onClick="document.myform.action='Admin_ArticlePreview.asp';document.myform.target='_blank';" style="cursor: hand;background-color: #cccccc;">
      &nbsp;&nbsp;
      <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor: hand;background-color: #cccccc;">
 </td></tr>
</form>
 </table>
 <!--#include file="Admin_fooder.asp"-->
</body>
</html>

