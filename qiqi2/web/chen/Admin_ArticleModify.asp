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
dim ArticleID,sql,rsArticle,FoundErr,ErrMsg,PurviewChecked,mUserName
dim Author,AuthorName,AuthorEmail,CopyFrom,CopyFromName,CopyFromUrl
dim ClassID,tClass,ClassName,RootID,ParentID,Depth,ParentPath,ClassMaster
dim SkinID,LayoutID,SkinCount,LayoutCount,BrowsePurview,AddPurview
ArticleID=trim(request("ArticleID"))
FoundErr=False
PurviewChecked=False
if ArticleID="" then 
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵ�����ID</li>"
	call WriteErrMsg()
	'call CloseConn()
	response.end
else
	ArticleID=Clng(ArticleID)
end if
sql="select * from "&jieducm&"_article where ArticleID=" & ArticleID & ""
Set rsArticle= Server.CreateObject("ADODB.Recordset")
rsArticle.open sql,conn,1,1
if rsArticle.bof and rsArticle.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�Ҳ�������</li>"
else
	ClassID=rsArticle("ClassID")
	mUserName = rsArticle("mUserName")
	set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,ClassMaster From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ</li>"
	else
		ClassName=tClass(0)
		RootID=tClass(1)
		ParentID=tClass(2)
		Depth=tClass(3)
		ParentPath=tClass(4)
		ClassMaster=tClass(5)
	end if
	if rsArticle("Editor")=AdminName and rsArticle("Passed")=False then
		PurviewChecked=True
	else
		if AdminPurview=1 or AdminPurview_Article<=2 then
			PurviewChecked=True
		else
			PurviewChecked=CheckClassMaster(ClassMaster,AdminName)
			if PurviewChecked=False and ParentID>0 then
				set tClass=nt2003.execute("select ClassMaster from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
				do while not tClass.eof
					PurviewChecked=CheckClassMaster(tClass(0),AdminName)
					if PurviewChecked=True then exit do
					tClass.movenext
				loop
			end if
		end if
		if PurviewChecked=False then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�Բ�������Ȩ�޲����������޸Ĵ��ģ�</li>"
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
	Author=rsArticle("Author")
	CopyFrom=rsarticle("CopyFrom")
	if instr(Author,"|")>0 then
		AuthorName=left(Author,instr(Author,"|")-1)
		AuthorEmail=right(Author,len(Author)-instr(Author,"|"))
	else
		AuthorName=Author
		AuthorEmail=""
	end if
	if instr(CopyFrom,"|")>0 then
		CopyFromName=left(CopyFrom,instr(CopyFrom,"|")-1)
		CopyFromUrl=right(CopyFrom,len(CopyFrom)-instr(CopyFrom,"|"))
	else
		CopyFromName=CopyFrom
		CopyFromUrl=""
	end if
%>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
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
  document.myform.PaginationType.selectedIndex=2;
}
function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

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
<style type="text/css">
<!--
.style1 {
	color: #ff6600;
	font-weight: bold;
}
.style2 {color: #FF6600}
.style3 {color: #0066cc}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_ArticleSave.asp?action=Modify">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� ��</b></td>
	</tr>
	<tr valign="middle" class="tdbg">
	  <td width="100" align="right"><strong>������Ŀ��</strong></td>
      <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <tr>
                 <td><%
			if AdminPurview=1 or AdminPurview_Article<=2 then
			 	response.write "<select name='ClassID' onChange='chsel();'>"	
				call Admin_ShowClass_Option(3,rsArticle("ClassID"))
				response.write "</select></td><td>"
				response.write "<font color='#FF6600'><strong>ע�⣺</strong></font><font color='#0066cc'>1������ָ��Ϊ��������Ŀ����Ŀ�������ⲿ��Ŀ</font>"
			 else
			 	call Admin_ShowPath2(ParentPath,ClassName,Depth)
				response.write "<input type='hidden' name='ClassID' value='" & rsArticle("ClassID") & "'>"
			 end if
			 %>
                 </td>
               </tr>
        </table>
</td>
  </tr>
          <!--<tr class="tdbg"> 
            <td width="100" align="right"><strong>����ר�⣺</strong></td>
            <td> 
              <%
			  if AdminPurview=1 or AdminPurview_Article<=2 then
			  	call Admin_ShowSpecial_Option(1,rsArticle("SpecialID"))
			  else
				if rsArticle("SpecialID")>0 then
					dim rsSpecial
					set rsSpecial=nt2003.execute("select * from Special where SpecialID=" & rsArticle("SpecialID"))
					if rsSpecial.bof and rsSpecial.eof then
						response.write "�Ҳ�������ר�⣡��������ר���Ѿ���ɾ����"
					else
						response.write rsSpecial("SpecialName")
					end if
					set rsSpecial=nothing
				end if
				response.write "<input type='hidden' name='SpecialID' value='" & rsArticle("SpecialID") & "'>"
			  end if%>
            </td>
          </tr>-->
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>���±��⣺</strong></td>
            <td><input name="Title" type="text"
           id="Title" value="<%=rsArticle("Title")%>" size="50" maxlength="255">
              <font color="#FF6600">*</font>
              <select name="TitleFontColor" id="TitleFontColor">
                <option value="" <%if rsArticle("TitleFontColor")="" then response.write " selected"%>>��ɫ</option>
                <OPTION value="">Ĭ��</OPTION>
                <OPTION value="#000000" style="background-color:#000000" <%if rsArticle("TitleFontColor")="#000000" then response.write " selected"%>></OPTION>
                <OPTION value="#FFFFFF" style="background-color:#FFFFFF" <%if rsArticle("TitleFontColor")="#FFFFFF" then response.write " selected"%>></OPTION>
                <OPTION value="#008000" style="background-color:#008000" <%if rsArticle("TitleFontColor")="#008000" then response.write " selected"%>></OPTION>
                <OPTION value="#800000" style="background-color:#800000" <%if rsArticle("TitleFontColor")="#800000" then response.write " selected"%>></OPTION>
                <OPTION value="#808000" style="background-color:#808000" <%if rsArticle("TitleFontColor")="#808000" then response.write " selected"%>></OPTION>
                <OPTION value="#000080" style="background-color:#000080" <%if rsArticle("TitleFontColor")="#000080" then response.write " selected"%>></OPTION>
                <OPTION value="#800080" style="background-color:#800080" <%if rsArticle("TitleFontColor")="#800080" then response.write " selected"%>></OPTION>
                <OPTION value="#808080" style="background-color:#808080" <%if rsArticle("TitleFontColor")="#808080" then response.write " selected"%>></OPTION>
                <OPTION value="#FFFF00" style="background-color:#FFFF00" <%if rsArticle("TitleFontColor")="#FFFF00" then response.write " selected"%>></OPTION>
                <OPTION value="#00FF00" style="background-color:#00FF00" <%if rsArticle("TitleFontColor")="#00FF00" then response.write " selected"%>></OPTION>
                <OPTION value="#00FFFF" style="background-color:#00FFFF" <%if rsArticle("TitleFontColor")="#00FFFF" then response.write " selected"%>></OPTION>
                <OPTION value="#FF00FF" style="background-color:#FF00FF" <%if rsArticle("TitleFontColor")="#FF00FF" then response.write " selected"%>></OPTION>
                <OPTION value="#FF0000" style="background-color:#FF0000" <%if rsArticle("TitleFontColor")="#FF0000" then response.write " selected"%>></OPTION>
                <OPTION value="#0000FF" style="background-color:#0000FF" <%if rsArticle("TitleFontColor")="#0000FF" then response.write " selected"%>></OPTION>
                <OPTION value="#008080" style="background-color:#008080" <%if rsArticle("TitleFontColor")="#008080" then response.write " selected"%>></OPTION>
              </select>
              <select name="TitleFontType" id="TitleFontType">
                <option value="0" <%if rsArticle("TitleFontType")="0" then response.write " selected"%>>����</option>
                <option value="1" <%if rsArticle("TitleFontType")="1" then response.write " selected"%>>����</option>
                <option value="2" <%if rsArticle("TitleFontType")="2" then response.write " selected"%>>б��</option>
                <option value="3" <%if rsArticle("TitleFontType")="3" then response.write " selected"%>>��+б</option>
                <option value="0" <%if rsArticle("TitleFontType")="4" then response.write " selected"%>>����</option>
              </select>
            </td>
          </tr>
		  <tr>
		  <td colspan="3" id="mrmq">
		  <%IF ClassID=29 Then%><table width=100%><tr class=tdbg><td width=112 align=right><strong>����������</strong></td><td colspan=2 >&nbsp;<input name=mUserName type=text value="<%=mUserName%>" id=mUserName size=20 maxlength=8></td></tr></table><%End IF%>
		  </td>
		  </tr>
          <% if  rsArticle("Title1")<>"" then %>
         <tr class="tdbg"> 
         <td width="100" align="right"><strong>���¸����⣺</strong></td>
          <td><input name="Title1" type="text"
           id="Title1" value="<%=rsArticle("Title1")%>" size="50" maxlength="255"></td>
         </tr>
          <% end if %>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>�ؼ��֣�</strong></td>
            <td><input name="Key" type="text"
           id="Key" value="<%=rsArticle("Key1")%>" size="50" maxlength="255"> 
              <font color="#FF6600">*</font><br>
              <font color="#0066cc">��������������£����������ؼ��֣��м���</font><font color="#FF6600">��|��</font><font color="#0066cc">�ֿ������ܳ���&quot;'*?()���ַ���</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>���ߣ�</strong></td>
            <td>������ 
              <input name="AuthorName" type="text"
           id="AuthorName" value="<%=AuthorName%>" size="20" maxlength="30"><!-- 
              &nbsp;&nbsp;&nbsp;&nbsp;Email�� 
              <input name="AuthorEmail" type="text" id="AuthorEmail" value="<%=AuthorEmail%>" size="40" maxlength="100"> -->
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>ԭ������</strong></td>
            <td>���ƣ� 
              <input name="CopyFromName" type="text"
           id="CopyFromName" value="<%=CopyFromName%>" size="20" maxlength="50"> 
              <!--&nbsp;&nbsp;&nbsp;&nbsp;�� ַ�� 
              <input name="CopyFromUrl" type="text" id="CopyFromUrl" value="<%=CopyFromUrl%>" size="40" maxlength="200">--></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>�������ݣ�</strong></p>
              <p align="left"><font color="#0066cc">&middot;������Ǵ�������վ�ϸ������ݣ����������а�����ͼƬ����ϵͳ�����޸Ľ��ʱ���ѷǱ�վͼƬ���Ƶ���վ�������ϣ�ϵͳ������ͼƬ�Ĵ�С��Ӱ���ٶȣ����Ժ򣨴˹�����Ҫ��������װ��IE5.5���ϰ汾����Ч����<br>
                <br>
                &middot;�������밴</font><font color="#ff6600">Shift+Enter</font><br>
                <font color="#0066cc">&middot;������һ���밴</font><font color="#ff6600">Enter</font></p></td>
            <td> <textarea name="Content" style="display:none"></textarea>
               <iframe ID="editor" src="editor.asp?Action=Modify&ArticleID=<%=ArticleID%>" frameborder=1 scrolling=no width="600" height="405"></iframe>  
            </td>
          </tr>
        
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>����ͼƬ��</strong></td>
            <td><input name="IncludePic" type="checkbox" id="IncludePic" value="yes" <% if rsArticle("IncludePic")=true then response.Write("checked") end if%> style="border: 0px;background-color: #eeeeee;">
              ��<font color="#0066cc">�����ѡ�еĻ����ڱ���ǰ����ʾ[ͼ��]��</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>��ҳͼƬ��</strong></td>
            <td><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="<%=rsArticle("DefaultPicUrl")%>" size="56" maxlength="200">
              ��������ҳ��ͼƬ���´���ʾ <br>
              ֱ�Ӵ��ϴ�ͼƬ��ѡ�� 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option value=""<% if rsArticle("DefaultPicUrl")="" then response.write "selected" %>>��ָ����ҳͼƬ</option>
                <%
				if rsArticle("UploadFiles")<>"" then
					dim IsOtherUrl
					IsOtherUrl=True
					if instr(rsArticle("UploadFiles"),"|")>1 then
						dim arrUploadFiles,intTemp
						arrUploadFiles=split(rsArticle("UploadFiles"),"|")						
						for intTemp=0 to ubound(arrUploadFiles)
							if rsArticle("DefaultPicUrl")=arrUploadFiles(intTemp) then
								response.write "<option value='" & arrUploadFiles(intTemp) & "' selected>" & arrUploadFiles(intTemp) & "</option>"
								IsOtherUrl=False
							else
								response.write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
							end if
						next
					else
						if rsArticle("UploadFiles")=rsArticle("DefaultPicUrl") then
							response.write "<option value='" & rsArticle("UploadFiles") & "' selected>" & rsArticle("UploadFiles") & "</option>"
							IsOtherUrl=False
						else
							response.write "<option value='" & rsArticle("UploadFiles") & "'>" & rsArticle("UploadFiles") & "</option>"		
						end if
					end If
					if IsOtherUrl=True then
						response.write "<option value='" & rsArticle("DefaultPicUrl") & "' selected>" & rsArticle("DefaultPicUrl") & "</option>"
					end if
				end if
				 %>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles" value="<%=rsArticle("UploadFiles")%>"> 
            </td>
          </tr>
<%if AdminPurview=1 or AdminPurview_Article<=2 then%>
          <br>��ע�⣬�˹���ֻ���ǰ̨�û����Թ���Ա�޽�����</td>
          </tr>-->
<%'�޸����¶��û����в�����������%>
<%end if%>
<tr class="tdbg"><td colspan="2" align="center">
      <input name="ArticleID" type="hidden" id="ArticleID" value="<%=rsArticle("ArticleID")%>">
      <input
  name="Save" type="submit"  id="Save" value="&nbsp;�����޸Ľ��&nbsp;" style="cursor: hand;background-color: #cccccc;">
      &nbsp; 
      <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor: hand;background-color: #cccccc;">
</td></tr>
</form>  
</table>
<!--#include file="Admin_fooder.asp"-->
</body>
</html>
<%
end if
rsArticle.close
set rsArticle=nothing

%>