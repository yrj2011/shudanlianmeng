<!--#include file="../inc/in_conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#nclude file="inc/function.asp"-->
<!--#nclude file="inc/admin_code_article.asp"-->
<%
content = Request.Form("content")
IF content <> "" Then
conn.execute("update mb set moban='"&content&"'")
Response.Write("<script language='javascript'>alert('更新成功');location.href='admin_moban.asp';</script>")
Response.End()
End IF

Sql = "select * from mb"
Set Rs = Server.CreateObject("adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Not Rs.Eof Then
content = Rs("moban")
End IF
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>行业协会</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script language="javascript">
	function checkinfo()
	{
		if(document.hy.hyClass.value=="")
		{
			alert("摸板内容不能为空");
			document.hy.hyClass.focus();
			return false;
		}
		
	}
</script>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
    <tr>
      <td height="22" colspan="3" align="center" class="title"><strong>摸 板</strong></td>
    </tr>
</table>
 <table width="100%"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#505050" style="margin:2px; ">
   <form name="hy" method="post" action="" onSubmit="return checkinfo();">
   <tr bgcolor="#ECECEC">
     <td width="22%"><div align="right"><strong>内容：</strong></div></td>
     <td width="78%">	 <input name="content" type="hidden" value="<%=server.HTMLEncode(content)%>"><IFRAME ID="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=content&style=standard" frameborder="0" scrolling="no" width="574" height="350">
     </IFRAME>
	 </td>
   </tr>
   <tr bgcolor="#ECECEC">
     <td>&nbsp;</td>
     <td><div align="center">
       <input type="submit" name="Submit" value="提交">
        　 
        
     </div></td>
   </tr>
   </form>
</table>

 <!--#include file="Admin_fooder.asp"-->
</body>
</html>

