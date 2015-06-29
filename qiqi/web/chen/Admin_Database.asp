<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1
Const CheckChannelID=0
Const PurviewLevel_Others="Database"
'response.write "此功能被WEBBOY暂时禁止了！"
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
<title>数据库备份</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong>数 据 库 管 理</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_Database.asp?Action=Backup">备份数据库</a></td>
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
	ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
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
	<td colspan="3" align="center" class="title"><b>备 份 数 据 库</b></td>
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
            <td width="200" height="33" align="right">选择操作：</td>
            <td><INPUT TYPE="radio" NAME="act" id="act_backup"  value="backup"><label for=act_backup>备份数据库</label></td>
            <td>要备份数据库请选择此项</td>
          </tr>
          <tr class="tdbg"> 
            <td width="200" height="33" align="right">数据库名：</td>
            <td><input type=text size=20 NAME="databasename" value="<%=request("databasename")%>"></td>
            <td>请输入要备份的数据库名</td>
          </tr>
		  <tr class="tdbg"> 
            <td width="200" height="34" align="right">备份目录：</td>
            <td height="34"><input type=text size=20 NAME="bak_path" value="Database"></td>
            <td height="34">相对路径目录，如目录不存在，将自动创建</td>
          </tr>
          <tr class="tdbg"> 
            <td width="200" height="34" align="right">备份名称：</td>
            <td height="34"><input type=text size=20 NAME="bak_file" value="<%=date()%>.bak"></td>
            <td height="34">后缀名默认为“.bak”。如有同名文件，将覆盖</td>
          </tr>
         <tr class="tdbg" align="center">  
            <td height="40" colspan="3"><input name="submit" type=submit value=" 开始备份 "></td>
          </tr>
	 </table>

<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
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
	<td colspan="2" align="center" class="title"><b>数据库在线压缩</b></td>
	</tr>
	<tr class="tdbg">
		<td align="center" valign="middle"> 
<%    
if request("action")="CompactData" then
	call CompactData()
else
%>
<br>
      <font color="#FF6600"><b>注：</b></font>压缩前，建议先备份数据库，以免发生意外错误。 <br>
	  数据库位置：
	<input name="db" type="text" id="db" size="23" value="Database/vippop@#2004.asp">
	</td></tr>
	<tr class="tdbg"><td align="center">
	<input name="submit" type=submit value=" &nbsp;压缩数据库&nbsp; " <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;">
      </td> </tr>
     <%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
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
	<td colspan="2" align="center" class="title"><b>数据库恢复</b></td>
  </tr>
        <%
if request("action")="RestoreData" then
	call RestoreData()
else
%>
        <tr class="tdbg">
            <td width="200" height="30" align="right">备份数据库路径（相对）：</td>
            <td height="30"><input name=backpath type=text id="backpath" value="DataBackup\vippop@#2004.asp" size=50 maxlength="200"></td>
          </tr>
          <tr class="tdbg">
            <td width="200" height="30" align="right">当前数据库路径（相对）：</td>
            <td><input name="db" type="text" size="50" value="Database/vippop@#2004.asp"></td>
          </tr>
          <tr align="center" class="tdbg"> 
            <td colspan="2"><input name="submit" type=submit value=" &nbsp;恢复数据&nbsp; " <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;"></td>
          </tr>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
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
