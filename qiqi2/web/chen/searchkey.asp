
<%

response.buffer=true	
Const PurviewLevel=1
'response.write "�˹��ܱ�WEBBOY��ʱ��ֹ�ˣ�"
'response.end
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<%
TT = Request("TT")
ID = Request("ID")
Action=Trim(request("Action"))
IF Action = "del" Then
	conn.execute("delete * from searchkey where ID="&ID)
	Response.Write("<script>alert('ɾ���ɹ�');location.href='searchkey.asp?TT=admin';</script>")
End IF
key = Request("key")
IF Action = "add" Then
	conn.execute("insert into searchkey(key) values('"&key&"')")
	Response.Write("<script>alert('��ӳɹ�');location.href='searchkey.asp?TT=add';</script>")
End IF
%>
<html>
<head>
<title>����Ա����</title>
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
<style type="text/css">
<!--
.style16 {color: #0066FF;
	font-weight: bold;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>�� �� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="?TT=add">��ӹؼ���</a> | <a href="?TT=admin">����ؼ���</a></td>
  </tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
   <td>
   <%
   IF TT = "admin" Then
   	Sql = "select * from searchkey Order by ID desc"
	Set Rs = Server.CreateObject("adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Not Rs.Eof Then
		
   %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr align="center" class="title">
    <td height="22"><strong> �ؼ���</strong></td>
    <td  width="150" height="22"><strong> �� ��</strong></td>
  </tr>
  <%Do While Not Rs.Eof  
			n = n - 1 %>
  <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td><%=Rs("key")%></td>
    <td width="150"><a href="?Action=del&ID=<%=RS("ID")%>">ɾ��</a></td>
  </tr>
  <%
	rs.MoveNext
loop
  %>
</table> 
<%End IF%> 
<%Else%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr align="center" class="title">
    <td height="22" colspan="3"><strong> ��ӹؼ���</strong><strong> </strong></td>
    </tr>
	<form action="?Action=add" method="post" name="form1">
  <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
    <td width="119" align="center">�ؼ���</td>
    <td width="396" align="center"><input name="key" type="text" id="key" size="60"></td>
    <td width="474" align="left"><input type="submit" name="Submit" value="�ύ"></td>
  </tr>
</form>
</table>
<%End IF%>
   </td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>���߻�ˢƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>