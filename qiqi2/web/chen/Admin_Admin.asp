
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
Const PurviewLevel=1
'response.write "�˹��ܱ�WEBBOY��ʱ��ֹ�ˣ�"
'response.end
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim rs, sql, strPurview,iCount
dim Action,FoundErr,ErrMsg
Action=Trim(request("Action"))
%>
<html>
<head>
<title>���ߴ�ý</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
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
    if (e.Name != "chkAll"&&e.disabled!=true)
       e.checked = form.chkAll.checked;
    }
}

function CheckAdd()
{
  if(document.form1.username.value=="")
    {
      alert("�û�������Ϊ�գ�");
	  document.form1.username.focus();
      return false;
    }
    
  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
  if (document.form1.Purview[1].checked==true){
	GetClassPurview();
  }
}
function CheckModifyPwd()
{
  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}

function CheckModifyPurview()
{
  if (document.form1.Purview[1].checked==true){
	GetClassPurview();
  }
}

function GetClassPurview()
{
    document.form1.ClassInputer_Article.value="";
	document.form1.ClassChecker_Article.value="";
	document.form1.ClassMaster_Article.value="";
	if(document.form1.AdminPurview_Article[2].checked==true){
		for(var i=0;i<frmArticle.document.myform.Purview_Add.length;i++){
			if (frmArticle.document.myform.Purview_Add[i].checked==true){
				if (document.form1.ClassInputer_Article.value=="")
					document.form1.ClassInputer_Article.value=frmArticle.document.myform.Purview_Add[i].value;
				else
					document.form1.ClassInputer_Article.value+=","+frmArticle.document.myform.Purview_Add[i].value;
			}
		}
		for(var i=0;i<frmArticle.document.myform.Purview_Check.length;i++){
			if (frmArticle.document.myform.Purview_Check[i].checked==true){
				if (document.form1.ClassChecker_Article.value=="")
					document.form1.ClassChecker_Article.value=frmArticle.document.myform.Purview_Check[i].value;
				else
					document.form1.ClassChecker_Article.value+=","+frmArticle.document.myform.Purview_Check[i].value;
			}
		}
		for(var i=0;i<frmArticle.document.myform.Purview_Manage.length;i++){
			if (frmArticle.document.myform.Purview_Manage[i].checked==true){
				if (document.form1.ClassMaster_Article.value=="")
					document.form1.ClassMaster_Article.value=frmArticle.document.myform.Purview_Manage[i].value;
				else
					document.form1.ClassMaster_Article.value+=","+frmArticle.document.myform.Purview_Manage[i].value;
			}
		}
	}
    document.form1.ClassInputer_Soft.value="";
	document.form1.ClassChecker_Soft.value="";
	document.form1.ClassMaster_Soft.value="";
	if(document.form1.AdminPurview_Soft[2].checked==true){
		for(var i=0;i<frmSoft.document.myform.Purview_Add.length;i++){
			if (frmSoft.document.myform.Purview_Add[i].checked==true){
				if (document.form1.ClassInputer_Soft.value=="")
					document.form1.ClassInputer_Soft.value=frmSoft.document.myform.Purview_Add[i].value;
				else
					document.form1.ClassInputer_Soft.value+=","+frmSoft.document.myform.Purview_Add[i].value;
			}
		}
		for(var i=0;i<frmSoft.document.myform.Purview_Check.length;i++){
			if (frmSoft.document.myform.Purview_Check[i].checked==true){
				if (document.form1.ClassChecker_Soft.value=="")
					document.form1.ClassChecker_Soft.value=frmSoft.document.myform.Purview_Check[i].value;
				else
					document.form1.ClassChecker_Soft.value+=","+frmSoft.document.myform.Purview_Check[i].value;
			}
		}
		for(var i=0;i<frmSoft.document.myform.Purview_Manage.length;i++){
			if (frmSoft.document.myform.Purview_Manage[i].checked==true){
				if (document.form1.ClassMaster_Soft.value=="")
					document.form1.ClassMaster_Soft.value=frmSoft.document.myform.Purview_Manage[i].value;
				else
					document.form1.ClassMaster_Soft.value+=","+frmSoft.document.myform.Purview_Manage[i].value;
			}
		}
	}
	document.form1.ClassInputer_Photo.value="";
	document.form1.ClassChecker_Photo.value="";
	document.form1.ClassMaster_Photo.value="";
    if(document.form1.AdminPurview_Photo[2].checked==true){
		for(var i=0;i<frmPhoto.document.myform.Purview_Add.length;i++){
			if (frmPhoto.document.myform.Purview_Add[i].checked==true){
				if (document.form1.ClassInputer_Photo.value=="")
					document.form1.ClassInputer_Photo.value=frmPhoto.document.myform.Purview_Add[i].value;
				else
					document.form1.ClassInputer_Photo.value+=","+frmPhoto.document.myform.Purview_Add[i].value;
			}
		}
		for(var i=0;i<frmPhoto.document.myform.Purview_Check.length;i++){
			if (frmPhoto.document.myform.Purview_Check[i].checked==true){
				if (document.form1.ClassChecker_Photo.value=="")
					document.form1.ClassChecker_Photo.value=frmPhoto.document.myform.Purview_Check[i].value;
				else
					document.form1.ClassChecker_Photo.value+=","+frmPhoto.document.myform.Purview_Check[i].value;
			}
		}
		for(var i=0;i<frmPhoto.document.myform.Purview_Manage.length;i++){
			if (frmPhoto.document.myform.Purview_Manage[i].checked==true){
				if (document.form1.ClassMaster_Photo.value=="")
					document.form1.ClassMaster_Photo.value=frmPhoto.document.myform.Purview_Manage[i].value;
				else
					document.form1.ClassMaster_Photo.value+=","+frmPhoto.document.myform.Purview_Manage[i].value;
			}
		}
	}
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>�� �� Ա �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="Admin_Admin.asp">����Ա������ҳ</a>&nbsp;|&nbsp;<a href="Admin_Admin.asp?Action=Add">��������Ա</a></td>
  </tr>
</table>
<%
if Action="Add" then
	call AddAdmin()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="ModifyPwd" then
	call ModifyPwd()
elseif Action="ModifyPurview" then
	call ModifyPurview()
elseif Action="SaveModifyPwd" then
	call SaveModifyPwd()
elseif Action="SaveModifyPurview" then
	call SaveModifyPurview()
elseif Action="Del" then
	call DelAdmin()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
''call CloseConn() 'shiyu

sub main()
	If Not IsObject(Conn) Then ConnectionDatabase 'shiyu
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from "&jieducm&"_admin order by id"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
%>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="Admin_Admin.asp" onSubmit="return confirm('ȷ��Ҫɾ��ѡ�еĹ���Ա��');">
   <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr align="center" class="title">
    <td  width="30"><strong>ѡ��</strong></td>
    <td  width="30" height="22"><strong> ���</strong></td>
    <td height="22"><strong> �� �� ��</strong></td>
    <td  width="100" height="22"><strong> Ȩ ��</strong></td>
    <td width="100"><strong>����¼IP</strong></td>
    <td width="120"><strong>����¼ʱ��</strong></td>
    <td  width="150" height="22"><strong> �� ��</strong></td>
  </tr>
  <%do while not rs.EOF %>
  <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td width="30">
	  <input name="ID" type="checkbox" id="ID" value="<%=rs("ID")%>" onClick="unselectall()" style="border: 0px;background-color: #eeeeee;">	</td>
    <td width="30"><%=rs("ID")%></td>
    <td>
	  <%
	    if rs("username")=AdminName then
		  response.write "<font color=#FF6600><b>" & rs("UserName") & "</b></font>"
	    else
		response.write rs("UserName")
	    end if
	   %>	</td>
    <td width="100"> 
	  <%
		  select case rs("purview")
		    case 1
              strPurview="<font color=#0066CC>��������Ա</font>"
            case 2
              strpurview="��ͨ����Ա"
		  end select
		  response.write(strPurview)
       %>    </td>
    <td width="100">
	  <%
	    if rs("LastLoginIP")<>"" then
		  response.write rs("LastLoginIP")
	    else
		response.write "&nbsp;"
  	    end if
	  %>	</td>
    <td width="120">
	  <%
	    if rs("LastLoginTime")<>"" then
		  response.write rs("LastLoginTime")
   	    else
		response.write "&nbsp;"
	    end if
	  %>	</td>
    <td width="150">
	  <%
  	    response.write "<a href='Admin_Admin.asp?Action=ModifyPwd&ID=" & rs("ID") & "'>�޸�����</a>&nbsp;&nbsp;"
	    response.write "<a href='Admin_Admin.asp?Action=ModifyPurview&ID=" & rs("ID") & "'>�޸�Ȩ��</a>&nbsp;&nbsp;"
	    if iCount>1 and rs("UserName")<>AdminName then
		  response.write "<a href='Admin_Admin.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('ȷ��Ҫɾ���˹���Ա��');"">ɾ��</a>"
	    else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
	    end if
	  %>    </td>
  </tr>
  <%
	rs.MoveNext
loop
  %>
</table>  
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
	<tr valign="middle" class="tdbg">
      <td width="200">
	    <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
              ѡ�б�ҳ��ʾ�����й���Ա</td>
      <td>
	    <input name="Action" type="hidden" id="Action" value="Del">
        <input name="Submit" type="submit" id="Submit" value="&nbsp;ɾ��ѡ�еĹ���Ա&nbsp;" style="cursor: hand;background-color: #cccccc;">
	  </td>
  </tr>
</table>
</td>
</form>
</tr>
</table>
<%
	rs.Close
	set rs=Nothing
end sub

sub AddAdmin()
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:return CheckAdd();">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� �� Ա</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> �� �� ����</strong></td>
      <td width="65%" class="tdbg"><input name="username" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> ��ʼ���룺 </strong></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="Password">
      </font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> ȷ�����룺</strong></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="PwdConfirm">
      </font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong>Ȩ�����ã� </strong></td>
      <td width="65%" class="tdbg"><table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr>
          <td width="100"><input name="Purview" type="radio" value="1" checked onClick="PurviewDetail.style.display='none'" style="border: 0px;background-color: #eeeeee;">
��������Ա��</td>
          <td>ӵ������Ȩ�ޡ�</td>
        </tr>
        <tr>
          <td width="100"><input type="radio" name="Purview" value="2"  style="border: 0px;background-color: #eeeeee;">
��ͨ����Ա��</td>
          <td>
<input type="checkbox" name="class1" id="class1" value="1" >
            ��վ��Ϣ����<br>
<input type="checkbox" name="class2" id="class2" value="1">
����վ����\�ʼ�<br>
<input type="checkbox" name="class3" id="class3" value="1">
��������<br>
<input type="checkbox" name="class4" id="class4" value="1">
����������<br>
<input type="checkbox" name="class5" id="class5" value="1">
���Թ���<br>
<input type="checkbox" name="class6" id="class6" value="1">
�ϴ��ļ�����<br>
<input type="checkbox" name="class7" id="class7" value="1">
�Զ���ִ��SQ<br>
<input type="checkbox" name="class8" id="class8" value="1">
ϵͳ��ʶ��<br>
<input type="checkbox" name="class9" id="class9" value="1">
���¹���<br>
<input type="checkbox" name="class10" id="class10" value="1">
��ֵ������<br>
<input type="checkbox" name="class11" id="class11" value="1">
�������� <br>
<input type="checkbox" name="class12" id="class12" value="1">
�Ա����������  <br>
<input type="checkbox" name="class13" id="class13" value="1">
������������� <br>
<input type="checkbox" name="class14" id="class14" value="1">
�а����������<br> 
<input type="checkbox" name="class15" id="class15" value="1">
VIP���������<br>
<input type="checkbox" name="class16" id="class16" value="1">
�û����� <br>
<input type="checkbox" name="class17" id="class17" value="1">
��������б�<br>
<input type="checkbox" name="class18" id="class18" value="1">
������¼ <br>
<input type="checkbox" name="class19" id="class19" value="1">
������� <br>
<input type="checkbox" name="class20" id="class20" value="1">
������ <br>
<input type="checkbox" name="class21" id="class21" value="1">
���߹��� <br>
<input type="checkbox" name="class22" id="class22" value="1">
SQLע�����  <br>

 </td>
        </tr>
      </table>
</td>
    </tr>
    <tr class="tdbg">
      <td colspan="2"><table id="PurviewDetail" width="100%" border="0" cellspacing="10" cellpadding="0" style="display:none">
        <tr>
          <td colspan="2" align="center" class="title"><strong>�� �� Ա Ȩ �� �� ϸ �� ��</strong></td>
        </tr>
        <tr valign="top">
          <td width="49%"><fieldset><legend>����Ƶ��</legend>
              <input type="radio" name="AdminPurview_Article" value="1" style="border: 0px;background-color: #eeeeee;">
Ƶ������Ա��ӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��<br>
<input type="radio" name="AdminPurview_Article" value="2" style="border: 0px;background-color: #eeeeee;">
��Ŀ�ܱࣺӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��<br>
<input name="AdminPurview_Article" type="radio" value="3" checked style="border: 0px;background-color: #eeeeee;">
��Ŀ����Ա��ֻӵ�в�����Ŀ����Ȩ�ޣ����������Ŀ��ר��<br>
<input type="radio" name="AdminPurview_Article" value="4" style="border: 0px;background-color: #eeeeee;"> 
�ڴ�Ƶ�������κι���Ȩ��<br>
<iframe id="frmArticle" height="200" width="100%" src="Admin_SetClassPurview.asp?ChannelID=2"></iframe>  <br>
<strong><font color="#FF6600">ע��</font></strong>��ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ�����Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ���������Ȩ�ޣ�����������Ŀ��ָ������Ĺ���Ȩ�ޡ�
          <input name="ClassMaster_Article" type="hidden" id="ClassMaster_Article">
          <input name="ClassChecker_Article" type="hidden" id="ClassChecker_Article">
          <input name="ClassInputer_Article" type="hidden" id="ClassInputer_Article">
          </fieldset></td>
        </tr>
        <tr valign="top">
          <td width="49%"><fieldset><legend>���԰�</legend>
              <table width="100%" border="0" cellspacing="1" cellpadding="2">
                <tr>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Reply" style="border: 0px;background-color: #eeeeee;">�ظ����� </td>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Modify" style="border: 0px;background-color: #eeeeee;">�޸�����</td>
                </tr>
                <tr>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Del" style="border: 0px;background-color: #eeeeee;">ɾ������</td>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Check" style="border: 0px;background-color: #eeeeee;">�������</td>
                </tr>
              </table>
          </fieldset><br>
            <fieldset>
            <legend>��վ����Ȩ��</legend>
            <table width="100%" border="0" cellspacing="1" cellpadding="2">
              <tr>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Channel" style="border: 0px;background-color: #eeeeee;">Ƶ������</td>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="AD" style="border: 0px;background-color: #eeeeee;">������</td>
              </tr>
              <tr>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Announce" style="border: 0px;background-color: #eeeeee;">�������</td>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="FriendSite" style="border: 0px;background-color: #eeeeee;">�������ӹ���</td>
              </tr>
              <tr>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Vote" style="border: 0px;background-color: #eeeeee;">�������&nbsp;</td>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Counter" style="border: 0px;background-color: #eeeeee;">ͳ�ƹ���</td>
              </tr>
              <tr>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="User" style="border: 0px;background-color: #eeeeee;">ע���û�����</td>
                <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="MailList" style="border: 0px;background-color: #eeeeee;">�ʼ��б����</td>
              </tr>
              <tr>
                <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Skin" style="border: 0px;background-color: #eeeeee;">��ɫģ�����</td>
                <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Layout" style="border: 0px;background-color: #eeeeee;">�������ģ�����</td>
              </tr>
              <tr>
                <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="JS" style="border: 0px;background-color: #eeeeee;">JS�������</td>
                <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Database" disabled style="border: 0px;background-color: #eeeeee;">���ݿ����</td>
              </tr>
              <tr>
                <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="UpFile" style="border: 0px;background-color: #eeeeee;">�ϴ��ļ�����</td>
                <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="ModifyPwd" checked style="border: 0px;background-color: #eeeeee;">�޸��Լ��ĵ�¼����</td>
              </tr>
            </table>
            </fieldset></td>
        </tr>
      </table></td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg">
	    <input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input  type="submit" name="Submit" value=" ��&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;">&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_Admin.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
  </form>
</table>
<%
end sub

sub ModifyPwd()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	sql="Select * from "&jieducm&"_Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����ڴ��û���</li>"
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:return CheckModifyPwd();">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� �� Ա �� ��</strong></font></div></td>
    </tr>
    <tr> 
      <td width="16%" class="tdbg"><strong>�� �� ����</strong></td>
      <td width="84%" class="tdbg"><%=rs("UserName")%> <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr> 
      <td width="16%" class="tdbg"><strong>�� �� �룺</strong></td>
      <td width="84%" class="tdbg"><input type="password" name="Password"></td>
    </tr>
    <tr> 
      <td width="16%" class="tdbg"><strong>ȷ�����룺</strong></td>
      <td width="84%" class="tdbg"><input type="password" name="PwdConfirm"></td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg">
	    <input name="Action" type="hidden" id="Action" value="SaveModifyPwd"> 
        <input  type="submit" name="Submit" value="&nbsp;�����޸Ľ��&nbsp;" style="cursor: hand;background-color: #cccccc;">&nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_Admin.asp'" style="cursor: hand;background-color: #cccccc;">
	  </td>
    </tr>
</form>
</table>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub ModifyPurview()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	sql="Select * from "&jieducm&"_Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����ڴ��û���</li>"
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:CheckModifyPurview();">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� �� Ա Ȩ ��</strong></font></div></td>
    </tr>
    <tr> 
      <td width="16%" class="tdbg"><strong>�� �� ����</strong></td>
      <td class="tdbg">&nbsp;&nbsp;<%=rs("UserName")%> <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="215" class="tdbg"><strong>Ȩ�����ã� </strong></td>
      <td width="767" class="tdbg"><table width="100%" border="0" cellspacing="1" cellpadding="2">
     
        <tr>
          <td width="100">
		    <input name="Purview" type="radio" value="1" onClick="PurviewDetail.style.display='none'" <%if rs("Purview")=1 then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">��������Ա��		  </td>
          <td>ӵ������Ȩ�ޡ�</td>
        </tr>
       
           <tr>
          <td> <input type="radio" name="Purview" value="2" <%if rs("Purview")=2 then response.write "checked"%>  style="border: 0px;background-color: #eeeeee;">��ͨ����Ա��</td>
          <td>
          
<input type="checkbox" name="class1" id="class1" <% if rs("class1")=1 then%> checked <% end if%> value="1" >
            ��վ��Ϣ����<br>
<input type="checkbox" name="class2" id="class2" value="1" <% if rs("class2")=1 then%> checked <% end if%>>
����վ����\�ʼ�<br>
<input type="checkbox" name="class3" id="class3" value="1" <% if rs("class3")=1 then%> checked <% end if%>>
��������<br>
<input type="checkbox" name="class4" id="class4" value="1" <% if rs("class4")=1 then%> checked <% end if%>>
����������<br>
<input type="checkbox" name="class5" id="class5" value="1" <% if rs("class5")=1 then%> checked <% end if%>>
���Թ���<br>
<input type="checkbox" name="class6" id="class6" value="1" <% if rs("class6")=1 then%> checked <% end if%>>
�ϴ��ļ�����<br>
<input type="checkbox" name="class7" id="class7" value="1" <% if rs("class7")=1 then%> checked <% end if%>>
�Զ���ִ��SQ<br>
<input type="checkbox" name="class8" id="class8" value="1" <% if rs("class8")=1 then%> checked <% end if%>>
ϵͳ��ʶ��<br>
<input type="checkbox" name="class9" id="class9" value="1" <% if rs("class9")=1 then%> checked <% end if%>>
���¹���<br>
<input type="checkbox" name="class10" id="class10" value="1" <% if rs("class10")=1 then%> checked <% end if%>>
��ֵ������<br>
<input type="checkbox" name="class11" id="class11" value="1" <% if rs("class11")=1 then%> checked <% end if%>>
�������� <br>
<input type="checkbox" name="class12" id="class12" value="1" <% if rs("class12")=1 then%> checked <% end if%>>
�Ա����������  <br>
<input type="checkbox" name="class13" id="class13" value="1" <% if rs("class13")=1 then%> checked <% end if%>>
������������� <br>
<input type="checkbox" name="class14" id="class14" value="1" <% if rs("class14")=1 then%> checked <% end if%>>
�а����������<br> 
<input type="checkbox" name="class15" id="class15" value="1" <% if rs("class15")=1 then%> checked <% end if%>>
VIP���������<br>
<input type="checkbox" name="class16" id="class16" value="1" <% if rs("class16")=1 then%> checked <% end if%>>
�û����� <br>
<input type="checkbox" name="class17" id="class17" value="1" <% if rs("class17")=1 then%> checked <% end if%>>
��������б�<br>
<input type="checkbox" name="class18" id="class18" value="1" <% if rs("class18")=1 then%> checked <% end if%>>
������¼ <br>
<input type="checkbox" name="class19" id="class19" value="1" <% if rs("class19")=1 then%> checked <% end if%>>
������� <br>
<input type="checkbox" name="class20" id="class20" value="1" <% if rs("class20")=1 then%> checked <% end if%>>
������ <br>
<input type="checkbox" name="class21" id="class21" value="1" <% if rs("class21")=1 then%> checked <% end if%>>
���߹��� <br>
<input type="checkbox" name="class22" id="class22" value="1" <% if rs("class22")=1 then%> checked <% end if%>>
SQLע�����   </td>
        </tr>
      </table>
     </td>
    </tr>
	<tr valign="middle" class="tdbg">
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg">
	    <input name="Action" type="hidden" id="Action" value="SaveModifyPurview"> 
        <input  type="submit" name="Submit" value="&nbsp;�����޸Ľ��&nbsp;" style="cursor: hand;background-color: #cccccc;">&nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_Admin.asp'" style="cursor: hand;background-color: #cccccc;">
	  </td>
    </tr>
</form>
</table>
<%
	end if
	rs.close
	set rs=nothing
end sub
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com/' target="_blank"><font color=#ff6600><strong>���߻�ˢƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
<%
sub SaveAdd()
	dim username, password,PwdConfirm, purview
	dim AdminPurview_Article,AdminPurview_Guest,AdminPurview_Others
	dim ClassInputer_Article,ClassChecker_Article,ClassMaster_Article

	username=trim(Request("username"))
	password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	purview=trim(Request("purview"))
	class1=request("class1")
	class2=request("class2")
	class3=request("class3")
	class4=request("class4")
	class5=request("class5")
	class6=request("class6")
	class7=request("class7")
	class8=request("class8")
	class9=request("class9")
	class10=request("class10")
	class11=request("class11")
	class12=request("class12")
	class13=request("class13")
	class14=request("class14")
	class15=request("class15")
	class16=request("class16")
	class17=request("class17")
	class18=request("class18")
	class19=request("class19")
	class20=request("class20")
	class21=request("class21")
		class22=request("class22")
	AdminPurview_Article=Clng(trim(Request("AdminPurview_Article")))
	AdminPurview_Guest=replace(replace(trim(request("AdminPurview_Guest"))," ",""),"'","")
	AdminPurview_Others=replace(replace(trim(request("AdminPurview_Others"))," ",""),"'","")
	ClassInputer_Article=replace(replace(trim(request("ClassInputer_Article"))," ",""),"'","")
	ClassChecker_Article=replace(replace(trim(request("ClassChecker_Article"))," ",""),"'","")
	ClassMaster_Article=replace(replace(trim(request("ClassMaster_Article"))," ",""),"'","")
	if AdminPurview_Guest<>"" then
		AdminPurview_Guest=AdminPurview_Guest & "," & "Manage"
	end if
	if username="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ʼ���벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ������������ʼ������ͬ��</li>"
	end if
	if purview="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�û�Ȩ�޲���Ϊ�գ�</li>"
	else
		purview=CInt(purview)
	end if

	if FoundErr=True then
		exit sub
	end if
	sql="Select * from "&jieducm&"_Admin where username='"&username&"'"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if not rs.eof then
	  Response.Write("<script>alert('�˹���Ա�Ѵ���!');history.back();</script>")
      response.End()
	else	
   	rs.addnew
 	rs("username")=username
   	rs("password")=md5(password)
    rs("purview")=purview
	if purview=2 then
	rs("class1")=class1
	rs("class2")=class2
	rs("class3")=class3
	rs("class4")=class4
	rs("class5")=class5
	rs("class6")=class6
	rs("class7")=class7
	rs("class8")=class8
	rs("class9")=class9
	rs("class10")=class10
	rs("class11")=class11
	rs("class12")=class12
	rs("class13")=class13
	rs("class14")=class14
	rs("class15")=class15
	rs("class16")=class16
	rs("class17")=class17
	rs("class18")=class18
	rs("class19")=class19
	rs("class20")=class20
	rs("class21")=class21
	rs("class22")=class22
	end if
	rs.update
	end if
    rs.Close
	set rs=Nothing
	if AdminPurview_Article=3 then
		call AddClassMaster("ArticleClass","ClassInputer",UserName,ClassInputer_Article)
		call AddClassMaster("ArticleClass","ClassChecker",UserName,ClassChecker_Article)
		call AddClassMaster("ArticleClass","ClassMaster",UserName,ClassMaster_Article)
	end if
	Call main()
end sub

sub SaveModifyPwd()
	dim UserID, UserName,password,PwdConfirm
	UserID=trim(Request("ID"))
	password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
	else
		UserID=Clng(UserID)
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ�������������������ͬ��</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * from "&jieducm&"_Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����ڴ˹���Ա��</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs("password")=md5(password)
 	rs.update
	rs.Close
   	set rs=Nothing
    call main()
end sub

sub SaveModifyPurview()
	dim UserID,UserName,Purview
	dim AdminPurview_Article,AdminPurview_Guest,AdminPurview_Others
	dim ClassInputer_Article,ClassChecker_Article,ClassMaster_Article
	dim OldAdminPurview_Article
	UserID=trim(Request("ID"))
	purview=trim(Request("purview"))
	
	class1=request("class1")
	class2=request("class2")
	class3=request("class3")
	class4=request("class4")
	class5=request("class5")
	class6=request("class6")
	class7=request("class7")
	class8=request("class8")
	class9=request("class9")
	class10=request("class10")
	class11=request("class11")
	class12=request("class12")
	class13=request("class13")
	class14=request("class14")
	class15=request("class15")
	class16=request("class16")
	class17=request("class17")
	class18=request("class18")
	class19=request("class19")
	class20=request("class20")
	class21=request("class21")
	class22=request("class22")
	if AdminPurview_Guest<>"" then
		AdminPurview_Guest=AdminPurview_Guest & "," & "Manage"
	end if
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
	else
		UserID=Clng(UserID)
	end if
	if purview="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�û�Ȩ�޲���Ϊ�գ�</li>"
	else
		purview=CInt(purview)
	end if
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * from "&jieducm&"_Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����ڴ˹���Ա��</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	UserName=rs("UserName")
	OldAdminPurview_Article=rs("AdminPurview_Article")
    rs("purview")=purview
	if purview=1 then
		rs("AdminPurview_Article")=3
		rs("AdminPurview_Guest")=""
		rs("AdminPurview_Others")=""
	else
		rs("AdminPurview_Article")=AdminPurview_Article
		rs("AdminPurview_Guest")=AdminPurview_Guest
		rs("AdminPurview_Others")=AdminPurview_Others
	end if
	if purview=2 then
	rs("class1")=class1
	rs("class2")=class2
	rs("class3")=class3
	rs("class4")=class4
	rs("class5")=class5
	rs("class6")=class6
	rs("class7")=class7
	rs("class8")=class8
	rs("class9")=class9
	rs("class10")=class10
	rs("class11")=class11
	rs("class12")=class12
	rs("class13")=class13
	rs("class14")=class14
	rs("class15")=class15
	rs("class16")=class16
	rs("class17")=class17
	rs("class18")=class18
	rs("class19")=class19
	rs("class20")=class20
	rs("class21")=class21
	rs("class22")=class22
	end if
 	rs.update
	rs.Close
   	set rs=Nothing
	if OldAdminPurview_Article=3 then
		call RemoveClassMaster(""&jieducm&"_ArticleClass",UserName)
	end if
	if AdminPurview_Article=3 then
		call AddClassMaster("ArticleClass","ClassInputer",UserName,ClassInputer_Article)
		call AddClassMaster("ArticleClass","ClassChecker",UserName,ClassChecker_Article)
		call AddClassMaster("ArticleClass","ClassMaster",UserName,ClassMaster_Article)
	end if
    call main()
end sub

sub DelAdmin()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ���Ĺ���ԱID</li>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		sql="Select * from "&jieducm&"_Admin where ID in (" & UserID & ")"
	else
		UserID=clng(UserID)
		sql="select * from "&jieducm&"_Admin where ID=" & UserID
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	do while not rs.eof
		if rs("Purview")=2 then
			if rs("AdminPurview_Article")=3 then
				call RemoveClassMaster("ArticleClass",rs("UserName"))
			end if
		end if
		rs.delete
		rs.update
		rs.movenext
	loop
	rs.close
	set rs=nothing
	call main()
end sub

sub AddClassMaster(SheetName,FieldName,strUserName,arrClassID)
	if isNull(arrClassID) or arrClassID="" then
		exit sub
	end if
	dim sqlMaster,rsMaster,ClassMaster
	sqlMaster="select ClassID," & FieldName & " from " & SheetName & " where ClassID in (" & arrClassID & ") order by RootID,OrderID"
	Set rsMaster=Server.CreateObject("Adodb.RecordSet")
	rsMaster.open sqlMaster,conn,1,3
	do while not rsMaster.eof
		if isNull(rsMaster(1)) or rsMaster(1)="" then
			rsMaster(1)=strUserName
		else
			rsMaster(1)=rsMaster(1) & "|" & strUserName
		end if
		rsMaster.update
		rsMaster.movenext
	loop
	rsMaster.close
	set rsMaster=nothing
end sub

sub RemoveClassMaster(SheetName,strUserName)
	dim sqlMaster,rsMaster,ClassMaster,arrClassMaster,i
	sqlMaster="select ClassInputer,ClassChecker,ClassMaster from " & SheetName
	Set rsMaster=Server.CreateObject("Adodb.RecordSet")
	rsMaster.open sqlMaster,conn,1,3
	do while not rsMaster.eof
		rsMaster(0)=RemoveStr(rsMaster(0),strUserName)
		rsMaster(1)=RemoveStr(rsMaster(1),strUserName)
		rsMaster(2)=RemoveStr(rsMaster(2),strUserName)
		rsMaster.update
		rsMaster.movenext
	loop
	rsMaster.close
	set rsMaster=nothing
end sub

function RemoveStr(str1,str2)
	if isnull(str1) or str1="" then
		RemoveStr=""
		exit function
	end if
	if str2="" then
		RemoveStr=str1
		exit function
	end if
	if instr(str1,"|")>0 then
		dim arrStr,tempStr,i
		arrStr=split(str1,"|")
		for i=0 to ubound(arrStr)
			if arrStr(i)<>str2 then
				if tempStr="" then
					tempStr=arrStr(i)
				else
					tempStr=tempStr & "|" & arrStr(i)
				end if
			end if
		next
		RemoveStr=tempStr
	else
		if str1=str2 then
			RemoveStr=""
		else
			RemoveStr=str1
		end if
	end if	
end function
%>