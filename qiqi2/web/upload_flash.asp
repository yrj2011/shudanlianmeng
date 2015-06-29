<%			
if  session("AdminName")<>"" then
else
response.Redirect("../login.asp")
response.End()
end if
uppath=request("uppath")&"/"		'文件上传路径
filelx=request("filelx")				'文件上传类型
formName=request("formName")			'回传到上页面编辑框所在Form的Name
EditName=request("EditName")			'回传到上页面编辑框的Name
%>
<html>
<head>
<title>图片上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<script language="javascript">
<!--
function mysub()
{
		esave.style.visibility="visible";
}
-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="upfile_flash.asp" enctype="multipart/form-data" >
  <div id="esave" style="position:absolute; top:18px; left:40px; z-index:10; visibility:hidden"> 
    <TABLE WIDTH=340 BORDER=0 CELLSPACING=0 CELLPADDING=0>
      <TR><td width=20%></td>
	<TD bgcolor=#104A7B width="60%"> 
	<TABLE WIDTH=100% height=120 BORDER=0 CELLSPACING=1 CELLPADDING=0>
	<TR> 
	          <td bgcolor=#eeeeee align=center><font color=red>正在上传文件，请稍候...</font></td>
	</tr>
	</table>
	</td><td width=20%></td>
	</tr></table></div>
<br>
<table width="400" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
    <tr> 
      <td height="22" align="left" valign="middle" bgcolor="#f1f1f1" width="400">&nbsp;图片上传 
        <input type="hidden" name="filepath" value="<%=uppath%>">
        <input type="hidden" name="filelx" value="<%=filelx%>">
        <input type="hidden" name="EditName" value="<%=EditName%>">
        <input type="hidden" name="FormName" value="<%=formName%>">
        <input type="hidden" name="act" value="uploadfile">
 (允许上传的图片类型：　jpg　gif)
      </td>
    </tr>
    <tr align="center" valign="middle"> 
      <td align="left" id="upid" height="80" width="400" bgcolor="#FFFFFF"> 选择文件: 
        <input type="file" name="file1" style="width:300'"  class="wenbenkuang" value="">
      </td>
    </tr>
    <tr align="center" valign="middle"> 
      <td bgcolor="#f1f1f1" height="24" width="400"> 
        <input type="submit" name="Submit" value="开始上传" class="go-wenbenkuang" onClick="javascript:mysub()">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
