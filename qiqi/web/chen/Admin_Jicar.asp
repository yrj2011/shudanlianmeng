<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../md5.asp"-->
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
if session("AdminName")="" then
   response.redirect("admin_login.asp")
end if
if request("act")="delet" then
if request("id")="" then
 Response.Write("<script>alert('����ѡ��Ҫɾ���ļ�¼!');history.back();</script>")
 response.End()
end if
conn.execute"delete from "&jieducm&"_dj586_Jicar where id in ("&request("id")&")" 

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
rs.addnew
rs("username")=session("AdminName")  
rs("class")="ɾ����ֵ��"
rs("content")="ɾ������Ϊ"&request("kid")&"�ĳ�ֵ��"		
rs("jiegou")="�ɹ�"
rs.update
rs.close
conn.close
set conn=nothing
response.redirect Request.ServerVariables("HTTP_REFERER")

elseif request.querystring("act")="save" then

dayer=request("dayer")
no=request("not")
paymoney=request("paymoneyj")
if no="" then
 Response.Write("<script>alert('������������Ϊ��!');history.back();</script>")
 response.End()
elseif paymoney="" then
 Response.Write("<script>alert('��ֵ����Ϊ��!');history.back();</script>")
 response.End()
end if

set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_dj586_Jicar",conn,1,3
for i=1 to no
randomize
p=Chr(Int(Rnd * 26 + 65))
p1=Chr(Int(Rnd * 26 + 65))
p2=int(90000000*rnd)+10000000
p3=Chr(Int(Rnd * 26 + 65))
p4=Chr(Int(Rnd * 26 + 65))
p5=Chr(Int(Rnd * 26 + 65))
p6=Chr(Int(Rnd * 26 + 65))
p7=int(90*rnd)+10

k=int(90000000*rnd)+10000000
k2=int(90000000*rnd)+10000000
k3=k&k2

Sqlcar = "select top 1 id from "&jieducm&"_dj586_Jicar order by id desc"
Set Rscar = Server.CreateObject("Adodb.RecordSet")
Rscar.Open Sqlcar,conn,1,1
IF not Rscar.Eof Then
jieducm_kid=rscar("id")
End IF


carid=p&p1&p2&jieducm_kid&p3&p4&p5&p6&p7

Set rsr=server.createobject("ADODB.RECORDSET")
rsr.open "Select * From "&jieducm&"_beifei where ka='"&carid&"'" ,Conn,3,3  
if rsr.eof then
rsr.addnew
rsr("ka")=carid
rsr("pwd")=k3
rsr.update
else
 Response.Write("<script>alert('�п�����ͬ�ĳ�ֵ��!����������');history.back();</script>")
 response.End()
end if


rs.addnew
rs("carid")=carid
rs("carpws")=md5(k3)
rs("oklook")=1
rs("day")=dayer
rs("paymoney")=paymoney
rs("adduser")="����Ա"
rs("date")=Now()
rs("ok")=0
rs.update


next
rsr.close
rs.close
set rs=nothing


Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
rs.addnew
rs("username")=session("AdminName")  
rs("class")="�Զ����ɳ�ֵ��"
rs("content")="������"&no&"��"&paymoney&"Ԫ�ĳ�ֵ��"		
rs("jiegou")="�ɹ�"
rs.update
rs.close
conn.close
set conn=nothing
response.redirect Request.ServerVariables("HTTP_REFERER")

elseif request("act")="ji" then
id=request("id")
set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_dj586_Jicar where id="&id&"",conn,1,3
if not rs.eof then
 rs("ok")=2
 rs.update
 rs.close
 set rs=nothing
 
  Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")  
		rsr("class")="�����ֵ��"
		rsr("content")="�����Ϊ"&request("kid")&"�ĳ�ֵ��"		
		rsr("jiegou")="�ɹ�"
    	rsr.update
    	rsr.close
 
 conn.close
 set conn=nothing
 response.redirect Request.ServerVariables("HTTP_REFERER")
 end if
end if
%>
<html>
<head>
<title>���ߴ�ý-��ֵ������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Inc/Admin.css" type="text/css">
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
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }

function ConfirmDel()
{	
	if(document.myform.Action.value=="delet")
	{
		document.myform.action="admin_jicar.asp";
		if(confirm("ȷ��Ҫɾ��ѡ�еļ�¼�𣿱����������޷��ָ���"))
		    return true;
		else
			return false;
	}
	
}

</SCRIPT>

<script language="JavaScript">
<!--
    function CheckForm()
    {
        if (document.jieducm_myform.not.value==""){
            alert("������������Ϊ�գ�");
            document.jieducm_myform.not.focus();
            return false;
        }
 
        if (document.jieducm_myform.paymoneyj.value==""){
            alert("�㿨����Ϊ�գ�");
            document.jieducm_myform.paymoneyj.focus();
            return false;
        }
    }
-->
</script>
</head>
<body>
<table align="center" width="100%" border="1" cellspacing="0" cellpadding="4" class=dj586_Com_bk style="border-collapse: collapse">
<tr class=dj586_Com_ss> 
<td colspan="6"><div class='bodytitle'><div class='bodytitleleft'></div><div class='bodytitletxt'>��ֵ������ </div></div></td>
</tr>
<tr align="left" class=dj586_Com_ds>
<td colspan="6">  ��������<a href=Admin_Okjicar.asp>�㿨��ֵ��¼</a>| <a href="Admin_Jicar.asp" >��ֵ������</a> | </td>      
</tr>
</table><br>
  
    <table align="center" width="100%"   border="1" cellspacing="0" cellpadding="4" class="dj586_Com_bk" style="border-collapse: collapse">
      <tr class=dj586_Com_ss> 
        <td height=25 colspan=2> <div align="left"><div class='bodytitle'><div class='bodytitleleft'></div><div class='bodytitletxt'><b>�Զ����ɳ�ֵ����</b> </div></div></div></td>
      </tr>
      <tr> 
        <td height=23  class="dj586_Com_ds">
		<table width="353" border="0" align="center"> 
		  <form name="jieducm_myform" method="POST" action="?act=save" onSubmit="return CheckForm();">
          <tr>
            <td width="139">������
              <input name="not"  id="not" type="text" size="9" maxlength="2" onKeyUp="if(isNaN(value))execCommand('undo')">
            ��</td>
            <td width="130">��ֵ��
              <input name="paymoneyj" id="paymoneyj" type="text" size="9" maxlength="4" onKeyUp="if(isNaN(value))execCommand('undo')">
            Ԫ</td>
            <td width="70"><input class="button" name="B1" type="submit" value=" ���� " ></td>
          </tr> </form>
        </table></td>
        <td width="60%" class="dj586_Com_ds"><table width="315" border="0">
          <form name="form1" method="post" action="?action=jieducm_search">
          <tr>
            <td width="116">����Ҫ��ѯ�Ŀ��ţ�</td>
            <td width="144">
              <label>
                <input type="text" name="search">
              </label>
           
            </td>
            <td width="41"> 
              <label>
                <input type="submit" name="Submit" value="��ѯ">
              </label>
         
            </td>
          </tr>   </form>
        </table></td>
      </tr>
    </table>
<form name="myform" method="Post" action="admin_jicar.asp" onSubmit="return ConfirmDel();">
  
<table   width="100%" align="center" border="1" cellspacing="0" cellpadding="4" class="dj586_Com_bk" style="border-collapse: collapse">
  <tr class=dj586_Com_ss align="center">
    <td width="20%" colspan="2">����</td>

   
    <td width="10%">�� ֵ</td>
    <td width="10%">���ɻ�Ա</td>
    <td width="10%">����ʱ��</td>
       <td width="10%">״̬</td>
    <td width="10%">����</td>
  </tr>
 <%
l=request("l")
myPagesize=adflperpage
action=request("action")
search=request.Form("search")
set rs=server.createobject("adodb.recordset")
if action="jieducm_search" and search<>"" then
rs.open "select * from "&jieducm&"_dj586_Jicar where carid='"&search&"' and ok=0 or ok=2 order by id desc",conn,1,1
else
rs.open "select * from "&jieducm&"_dj586_Jicar where ok=0 or ok=2 order by id desc",conn,1,1
end if
call myPages(rs,myPagesize)
if rs.recordcount=0 then
response.write "<tr class=dj586_Com_ds align=center><td colspan=9>û������</td></tr>"
else
'�ָ�
line=myPagesize
dim i
i=1
do while not rs.eof and line>0
'ѭ����ʼ
%>
  <tr class="dj586_Com_ds" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td align="center" colspan="2"><input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>><a href="admin_jiagesear.asp?jieducm_ka=<%=rs("carid")%>" title="��ѯ����"><%=rs("carid")%></a></td>

   
    <td align="center"><%=rs("paymoney")%> Ԫ</td>
    <td align="center"><%=rs("adduser")%></td>
    <td> <%=rs("date")%></td>
     <td align="center"> <%if rs("ok")=0 then%>δ����<%elseif rs("ok")=2 then%><font color="red">�Ѽ���</font><% end if%></td>
    <td align="center"><a href="?act=ji&id=<%=rs("id")%>&kid=<%=rs("carid")%>"><font color="#000000">����</font></a>  <a href="?act=delet&id=<%=rs("id")%>"><font color="#000000">ɾ��</font></a></td>
  </tr>
<%
'ѭ������
rs.movenext
line=line-1
i=i+1
loop
end if
response.write "<tr class=dj586_Com_qs><td colspan=6 align='center'>&nbsp;&nbsp;"
call listpages(mycondition,request("l"))
response.write "</td><td><center>�� <font color=#FF0000>"&rs.recordcount&"</font> �ŵ㿨</center></td></tr>"
rs.close
set rs=nothing
%>

</table>
<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">ѡ�б�ҳ��ʾ�����м�¼
<input name="submit" type='submit' value='&nbsp;ɾ��ѡ���ļ�¼&nbsp;' onClick="document.myform.Act.value='delet'"  style="cursor: hand;background-color: #cccccc;">
              <input name="Act" type="hidden" id="Action" value="delet">
</form>
</body>
</html>
<%
sub listpages(mycondition,listtype)
	if pages<1 then
		exit sub
	end if
	
	
	if p>0 then
		response.write	"<a href="&request.ServerVariables("script_name")&"?currentpage=10&p="&p-1&mycondition&"&l="&listtype&"><img src=Images/prepre.gif alt=ǰʮҳ border=0></a>&nbsp;&nbsp;&nbsp;"
    else
	response.write	"<img src=Images/prepre.gif alt=ǰʮҳ border=0>&nbsp;&nbsp;&nbsp;"
    end if
	'-------------------����д����ǰʮ��ҳ��
	response.write "<img src=Images/lined.gif border=0><img src=Images/k.gif border=0 width=5>"
	for i=1 to 10
		if (p*10+i)>pages then exit for
		if currentpage=i then
		
			response.write "<b><a style='color:ff6600'"
		else
			response.write "<a"
		end if 
		response.write " href="&request.ServerVariables("script_name")&"?currentpage="&i&"&p="&p&mycondition&"&l="&listtype&"><font class=no >"&(p*10+i)&"</font></a></b>&nbsp;&nbsp;<img src=Images/lined.gif border=0><img src=Images/k.gif width=5 border=0>"	
     next
	 response.write ""
	 '--------------------�쿴��ʮҳ������
    if (p*10+10)<pages then
		response.write "<img src=Images/k.gif border=0><a href="&request.ServerVariables("script_name")&"?currentpage=1&p="&p+1&mycondition&"&l="&listtype&"><img src=Images/nextnext.gif alt=��ʮҳ border=0></a>&nbsp;&nbsp;" 
		else
		response.write "<img src=Images/k.gif border=0><img src=Images/nextnext.gif alt=��ʮҳ border=0>&nbsp;&nbsp;"
	end if
	'-----------------------------------------------����ҳ����
end sub

sub myPages(myRS,mysize)  '------mysizeΪ�ڲ���������ҳ��û�ж��壩��myRSΪ��ҳ�洫�ݹ�����RS���󣨵�ַ���ݣ�
	if myRS.eof and myRS.bof then str="û�м�¼"
	if str="" then
		if mysize="" or NOT IsNumeric(mysize) then
			mysize=15
		end if
		myRS.PageSize=mysize
		pages=myRS.pagecount
		records=myRS.recordcount
		On Error Resume Next 'ȡ������
		currentPage=request("currentPage")
		if currentPage="" then
			currentPage=1
		end if
		currentPage=CInt(currentPage)
		if Err.Number <> 0 Then
			currentPage=1
			Err.Clear
		end if
		if currentPage<1 then
			currentPage=1
		elseIf currentPage>10 then
			currentPage=10
		end if
		'----------------����p
		p=request("p")
		if p="" then
			p=0
		end if
		p=CLng(p)
		if Err.Number <> 0 Then
			p=0
			Err.Clear
		end if
		if p<0 then
			p=0
		end if
		'--�ж��Ƿ����ҳ����Χ
		nowPage=p*10+currentPage
		if nowPage>pages then
			p=(pages-1)\10
			currentPage=((pages-1) mod 10)+1
		end if
		myRS.absolutepage=p*10+currentPage
	else
	   currentPage=1
	   records=0
	   pages=1
	end if
end sub

dim pages,records,currentPage,p '--------�����������Щ��������ҳ���У����е�ַ����


%>