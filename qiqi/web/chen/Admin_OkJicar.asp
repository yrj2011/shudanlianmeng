<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#INCLUDE FILE="../background.asp"-->

<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
if session("AdminName")="" then
   response.redirect("admin_login.asp")
end if

%>
<html>
<head>
<title>点卡充值记录</title>
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
		if(confirm("确定要删除选中的记录吗？本操作将把无法恢复！"))
		    return true;
		else
			return false;
	}
	
}

</SCRIPT>
<body>
<table align="center" width="100%" border="1" cellspacing="0" cellpadding="4" class=dj586_Com_bk style="border-collapse: collapse">
<tr class=dj586_Com_ss> 
<td colspan="6"><div class='bodytitle'><div class='bodytitleleft'></div><div class='bodytitletxt'>点卡充值记录 </div></div></td>
</tr>
<tr align="left" class=dj586_Com_ds>
<td colspan="6">  管理导航：<a href=Admin_Okjicar.asp>点卡充值记录</a>| <a href="Admin_Jicar.asp">充值卡管理</a> | 
  <form name="form1" method="post" action="">
    <label>
      输入要查询的卡号或会员名：<input type="text" name="username">
      </label>
    <label>
    <input type="submit" name="Submit" value="查询">
	<input name="action" type="hidden" value="search">
    </label>
  </form>
  </td>      
</tr>
</table><br>
<form name="myform" method="Post" action="admin_jicar.asp" onSubmit="return ConfirmDel();">

<table align="center" width="100%" border="1" cellspacing="0" cellpadding="4" class=dj586_Com_bk style="border-collapse: collapse">
  <tr class=dj586_Com_ss align="center">
    <td width="20%">卡号</td>
     <td width="10%">冲值时间</td>
    <td width="10%">面 值</td>
    <td width="10%">充值会员</td>
    <td width="10%">操作</td>
  </tr>
<%
action=request.Form("action")
uname=trim(request.Form("username"))
if action="" then
sql="select * from "&jieducm&"_dj586_Jicar where ok=1 order by id desc" 
elseif action="search" then
sql="select * from "&jieducm&"_dj586_Jicar where ok=1 and  (userid='"&uname&"' or carid='"&uname&"') order by id desc" 
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	dim maxperpage,url,PageNo
	 url="admin_okjicar.asp"
	 rs.pagesize=20
	 PageNo=REQUEST("PageNo")
	 if PageNo="" or PageNo=0 then PageNo=1
	 RS.AbsolutePage=PageNo
	 TSum=rs.pagecount
	 maxperpage=rs.pagesize
	 RowCount=rs.PageSize
	   PageNo=PageNo+1
	   PageNo=PageNo-1
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  response.Write("对不起没有您想要的页数")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("没有这一页!")
		Response.End
	 End if
	i=0
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	 DO WHILE NOT rs.EOF AND RowCount>0%>
  
  <tr class="dj586_Com_ds" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td align="center"><input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>><%=rs("carid")%></td>

   <td align="center"><%=rs("adddate")%></td>
    <td align="center"><%=rs("paymoney")%> 元</td>
    <td align="center"><%=rs("userid")%></td>
    <td align="center"><font color="#000000">正在使用</font></td>
  </tr>
 <%
	i=i+1
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
	  <tr class="dj586_Com_ds" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
    <td align="center" colspan="5">    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %>
</td>
  </tr>
</table>
<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">选中本页显示的所有记录
<input name="submit" type='submit' value='&nbsp;删除选定的记录&nbsp;' onClick="document.myform.Act.value='delet'"  style="cursor: hand;background-color: #cccccc;">
              <input name="Act" type="hidden" id="Action" value="delet">
</form>
</body>
</html>
