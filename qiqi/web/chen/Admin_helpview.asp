<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
response.buffer=true	
Const PurviewLevel=2    '操作权限
Const CheckChannelID=0    '所属频道，0为不检测所属频道
Const PurviewLevel_Others="Skin"
%>
<br>
<p><p>
<table border="0" cellspacing="1" cellpadding="5" class=tableborder1 align=center style="width:90%">
<tr><td width="100%" class=tablebody1>
<script>
document.write(opener.txtRun.value)
</script>
</td></tr>
<tr>
<td height="22" align=center class=tablebody2>
<input type="button" name="close" value="[关  闭]" onclick="window.close()">
</td>
</tr>
</table>
<br>
</body>
</html>