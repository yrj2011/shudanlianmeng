
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
if request.Form("action")="ok" then
tdate=request.Form("startdate")
mdate=request.Form("enddate")
classid=request.Form("classid")
 if tdate="" or mdate="" then
 	 Response.Write("<script>alert('开始和结束日期都不能为空!');history.back();</script>")
     response.End()
 end if
 
 if tdate="请选择日期" or mdate="请选择日期" then
  	 Response.Write("<script>alert('开始和结束日期都不能为空!');history.back();</script>")
     response.End()
 end if
 
sql="select count(*) as yu  from "&jieducm&"_zhongxin where classid='"&classid&"' and start='5' and  (now between '"&tdate&"' and '"&mdate&"')"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
    countnum=rs("yu")
else
  	 Response.Write("<script>alert('数据库中无此日期范围内的数据!');history.back();</script>")
     response.End()
end if


 Sql = "select pid from "&jieducm&"_zhongxin where classid='"&classid&"' and start='5' and  (now between '"&tdate&"' and '"&mdate&"')"
 Set rs = Server.CreateObject("Adodb.RecordSet")
 rs.Open Sql,conn,1,1
 IF not rs.Eof Then
 Do While Not Rs.Eof
     Set rsj=server.createobject("ADODB.RECORDSET")
     rsj.open "delete  from "&jieducm&"_jieshou where pid = '"&rs("pid")&"' ",conn,3,3
  Rs.MoveNext
  Loop
 End IF
 
 

sqlfy="delete  from "&jieducm&"_zhongxin where classid='"&classid&"' and start='5' and  (now between '"&tdate&"' and '"&mdate&"') "
Conn.execute(sqlfy)

Response.Write("<script>alert('本次共删除"&countnum&"条记录!');window.location.href='admin_deltaobao.asp';</script>")
response.End()
	 
end if
%>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<script src="inc/jsdate.js" type="text/javascript"></script>

<form name="form1" method="post" action="admin_deltaobao.asp">
<input name="action" type="hidden" value="ok">
  <label></label>
  <label></label>
  <table width="100%" border="0">
    <tr>
      <td>&nbsp;</td>
      <td><span class="STYLE1">本功能删除已完成的任务&nbsp; 删除后无法恢复，请小心使用!如果记录非常多请耐心等待！<br>
        <br>
      </span></td>
    </tr>
    <tr>
      <td><div align="right">选择类别:</div></td>
      <td><label>
        <select name="classid">
          <option value="1">淘宝区</option>
          <option value="2">拍拍区</option>
		  <option value="5">充值提现区</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td width="149"><div align="right">你想删除:</div></td>
      <td width="853">从
        <input type="text" name="startdate"   value="请选择日期" onclick="SelectDate(this,'yyyy-MM-dd')" >
        <label>到</label>  <input type="text" name="enddate"   value="请选择日期" onclick="SelectDate(this,'yyyy-MM-dd')" ></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="批量删除已完成的任务" onClick="return confirm('确定要删除操作吗？');" ></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;&nbsp;</td>
    </tr>
  </table>
  <label><br>
  <br>
  </label>
</form>