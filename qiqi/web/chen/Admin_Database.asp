<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1
Const CheckChannelID=0
Const PurviewLevel_Others="Database"
'response.write "�˹��ܱ�WEBBOY��ʱ��ֹ�ˣ�"
'response.end
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->

<!--#include file="inc/function.asp"-->
<%
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
dim dbpath,db
dim ObjInstalled
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
%>
<html>
<head>
<title>���ݿⱸ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong>�� �� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="Admin_Database.asp?Action=Backup">�������ݿ�</a></td>
  </tr>
</table>
<%
if Action="Backup" or Action="BackupData" then
	set conn=nothing 
	call ShowBackup()
elseif Action="Compact" or Action="CompactData" then
	set conn=nothing 
	call ShowCompact()
elseif Action="Restore" or Action="RestoreData" then
	set conn=nothing 
	call ShowRestore()
elseif Action="Init" or Action="Clear" then
	call ShowInit()
	set conn=nothing 
elseif Action="SpaceSize" then
	call SpaceSize()
	set conn=nothing 
else
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���������</li>"
	set conn=nothing 
end if
if FoundErr=True then
	call WriteErrMsg()
end if

sub ShowBackup()
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="Admin_Database.asp?action=BackupData">
  <tr>
	<td colspan="3" align="center" class="title"><b>�� �� �� �� ��</b></td>
  </tr>
  <tr class="tdbg"> 
    <td height="150" align="center" valign="middle">
	<%    
if request("action")="BackupData" then
	call backupdata()
else
%>
	<table cellpadding="3" cellspacing="1" border="0" width="100%">
		  <tr class="tdbg"> 
            <td width="200" height="33" align="right">ѡ�������</td>
            <td><INPUT TYPE="radio" NAME="act" id="act_backup"  value="backup"><label for=act_backup>�������ݿ�</label></td>
            <td>Ҫ�������ݿ���ѡ�����</td>
          </tr>
          <tr class="tdbg"> 
            <td width="200" height="33" align="right">���ݿ�����</td>
            <td><input type=text size=20 NAME="databasename" value="<%=request("databasename")%>"></td>
            <td>������Ҫ���ݵ����ݿ���</td>
          </tr>
		  <tr class="tdbg"> 
            <td width="200" height="34" align="right">����Ŀ¼��</td>
            <td height="34"><input type=text size=20 NAME="bak_path" value="Database"></td>
            <td height="34">���·��Ŀ¼����Ŀ¼�����ڣ����Զ�����</td>
          </tr>
          <tr class="tdbg"> 
            <td width="200" height="34" align="right">�������ƣ�</td>
            <td height="34"><input type=text size=20 NAME="bak_file" value="<%=date()%>.bak"></td>
            <td height="34">��׺��Ĭ��Ϊ��.bak��������ͬ���ļ���������</td>
          </tr>
         <tr class="tdbg" align="center">  
            <td height="40" colspan="3"><input name="submit" type=submit value=" ��ʼ���� "></td>
          </tr>
	 </table>

<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
    </td>
 </tr>
 </form>
</table>
<%
end sub

sub ShowCompact()
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="Admin_Database.asp?action=CompactData">
  <tr>
	<td colspan="2" align="center" class="title"><b>���ݿ�����ѹ��</b></td>
	</tr>
	<tr class="tdbg">
		<td align="center" valign="middle"> 
<%    
if request("action")="CompactData" then
	call CompactData()
else
%>
<br>
      <font color="#FF6600"><b>ע��</b></font>ѹ��ǰ�������ȱ������ݿ⣬���ⷢ��������� <br>
	  ���ݿ�λ�ã�
	<input name="db" type="text" id="db" size="23" value="Database/vippop@#2004.asp">
	</td></tr>
	<tr class="tdbg"><td align="center">
	<input name="submit" type=submit value=" &nbsp;ѹ�����ݿ�&nbsp; " <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;">
      </td> </tr>
     <%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
      </td>
    </tr>
</form>
</table>
<%
end sub

sub ShowRestore()
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="Admin_Database.asp?action=RestoreData">
  <tr>
	<td colspan="2" align="center" class="title"><b>���ݿ�ָ�</b></td>
  </tr>
        <%
if request("action")="RestoreData" then
	call RestoreData()
else
%>
        <tr class="tdbg">
            <td width="200" height="30" align="right">�������ݿ�·������ԣ���</td>
            <td height="30"><input name=backpath type=text id="backpath" value="DataBackup\vippop@#2004.asp" size=50 maxlength="200"></td>
          </tr>
          <tr class="tdbg">
            <td width="200" height="30" align="right">��ǰ���ݿ�·������ԣ���</td>
            <td><input name="db" type="text" size="50" value="Database/vippop@#2004.asp"></td>
          </tr>
          <tr align="center" class="tdbg"> 
            <td colspan="2"><input name="submit" type=submit value=" &nbsp;�ָ�����&nbsp; " <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;"></td>
          </tr>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
            
</form>
</table>
<%
end sub


%>
<!--#include file="Admin_fooder.asp"-->
</body>
</html>
