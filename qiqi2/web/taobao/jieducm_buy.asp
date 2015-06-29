<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
taobaoid=request.QueryString("taobaoid")

%>
<SCRIPT src="../js/jieducm_zDialog.js" type=text/javascript></SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE>七七网</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
 <LINK href="../css/top_bottom.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.STYLE1 {font-size: 14px}
.STYLE2 {font-size: 14px}

body,td,th {
	font-size: 12px;
	color: #333;
}
body {
	margin-left: 20px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
}
-->
</style>
</HEAD>
<BODY>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0  borderColor=#dddddd cellPadding=3 width=100% align=center>
<FORM id=form method=post  name=form action="jieducm_jieshou.asp?id=<%=taobaoid%>">
  <tr height="30">
    <td colspan="3" bgcolor="#CCCCCC" height="30"><div align="center" class="STYLE1">请选择你购买此商品的淘宝用户名</div></td>
  </tr>
  <tr>
    <td colspan="3" height="20"><div align="left" class="STYLE2"><font color="#CC9900"><strong>已绑定的淘宝号</strong></font></div></td>
  </tr>
  <tr height="30">
    <td height="30" colspan="3"><font color="#669900"><strong>第一步：</strong></font>
      <%	
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=1 and start=0 order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
				    wei=0
					Response.Write("<a href=buyhao.asp>对不起，你还没有绑定买号！绑定买号后，才可以接手任务!</a>")
				Else
				    wei=1
					Do While Not Rs.Eof
			  %>
      <input type=radio name=xiaohaoname <%if checknum=rs("shopname") then%> checked <%end if%> value=<%=rs("shopname")%> />
      <%=rs("shopname")%>&nbsp;&nbsp;
      <%
			  	Rs.MoveNext
				Loop
				End IF
			  %></td>
    </tr>
  <tr height="30">
    <td width="12%" height="30">&nbsp;</td>
    <td height="30" colspan="2"><label>
      <input type="submit" name="Submit" value="提交">
    </label></td>
  </tr>

 
  
  <tr>
    <td colspan="3"><div align="center"> <a href="javascript:void(0)" onClick="Dialog.close()"><img src="../images/close.jpg" border="0"></a></div></td>
  </tr></FORM>
</table>

</BODY></HTML>
