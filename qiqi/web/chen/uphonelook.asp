<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">

<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->

<%
'------------------------------------------------------------------

'------------------------------------------------------------------
id=request("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select codenum From "&jieducm&"_register where id="&id&""  ,Conn,3,3  
if  not rs.eof then
code=rs("codenum")
rs.close
set rs=nothing
end if
%>

<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {font-size: 16px}
-->
</style>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">


<DIV style="WIDTH: 950px">
<DIV class=regheight 
style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 15px; PADDING-TOP: 15px"><BR>
  <BR>
  <DIV style="COLOR: red"><SPAN id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="COLOR: red"></SPAN></DIV>
  <DIV class=regsubmit>
  <div align="center">此用户手户验证码为：<%=code%><br>
    <br>
    &nbsp;  
  <INPUT name="按钮"  type=button style="CURSOR: pointer; HEIGHT: 30px" value="  返回  " onClick="history.back();">
  </div>
</DIV><BR><BR></DIV></DIV>


</DIV>
</BODY></HTML>
