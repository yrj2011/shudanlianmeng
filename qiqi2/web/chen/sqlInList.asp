<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->
<%
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../jieducm/#jieducm.asp")
conn.Open connstr

if request("action")="del" then
	Id_ary=split(request("idx"),",")
	for i=0 to ubound(Id_ary)
		conn.execute("delete from Rc_SqlIn where id="&Id_ary(i))	
	next
	response.write "<script language='javascript'>alert('批量删除操作成功!');location.href='"&Request.ServerVariables("HTTP_REFERER")&"';</script>"
	response.end
elseif request("action")="unlock" then
	conn.execute("update Rc_SqlIn set Kill_ip=False where id="&request("id"))
	response.write "<script language='javascript'>location.href='?moduleid="&request("moduleid")&"';</script>"
	response.end
elseif request("action")="lock" then
	conn.execute("update Rc_SqlIn set Kill_ip=true where id="&request("id"))
	response.write "<script language='javascript'>location.href='?moduleid="&request("moduleid")&"';</script>"
	response.end
end if

%>
<%
moduleid=request("moduleid")
'从模块号,人员号获取当前模块的权限
'算法:分列四种权限 1:添加cAdd  2:修改cEdit 3:删除cDel 4:审核cSh
'首先查找人员对此模块的权限,有则将各权限值为真,无则假;再查找该人员所在的角色,从角色找出该模块的权限,有则真,无则假,复合累加


	cAdd=true
	cEdit=true
	cDel=true
	cSh=true



%>
<html>
<head>
<title>top</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/style.css" type="text/css">
<script language="javascript">
function check(){
	var l=form1.elements.length;
	var s=0;
	for(var i=0;i<l;i++){
		if(form1.elements[i].checked){
			if(form1.elements[i].name!='chkall')
				s++;
		}
	}
	if(s==0){
		alert("请至少选中一条记录才能操作!");
		return false;
	}
	if(confirm("这将对你选择的记录,确定要做此操作吗?"))
		form1.submit();
}

function CheckAll(form) {
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name != 'chkall' && e.type=='checkbox')
			e.checked = form.chkall.checked;
	}
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="492" height="22" background="images/right_bg1.jpg">&nbsp;当前位置：<%=lmwz%>
    </td>
    <td width="507"><div align="left">
      <form name="form2" method="post" action="">
	  <input name="search" type="hidden" value="ok">
        <label>
        输入要查询的IP :
        <input type="text" name="ipsear">
        </label>
            <label>
            <input type="submit" name="Submit" value="查询">
            </label>
      </form>
      </div></td>
  </tr>
</table>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  
  <tr> 
    <td height="10"></td>
  </tr>
</table>
<form action="?action=del" method="post" name="form1" style="margin:0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#82D1E7" >
          <tr class="bb02">
            <td width="5%" align="center">编号</td>
            <td width="9%" align="center"><font color=red>操作ＩＰ</font></td>
            <td width="7%" align="center">是否锁定</td>
            <td width="6%" align="center">操作</td>
            <td width="19%" align="center">操作页面</td>
            <td width="16%" align="center">操作时间</td>
            <td width="8%" align="center">提交方式</td>
            <td width="9%" align="center">提交参数</td>
            <td width="16%" align="center">提交数据</td>
            <td width="5%" align="center">操作</td>
          </tr>
		  <%
		  i=0
		  MaxPerPage=20
		  if request("maxpars")<>"" then  MaxPerPage=request("maxpars")
		  set rs=server.CreateObject("adodb.recordset")
		  if request.Form("search")="" then
		  sql="select * from Rc_SqlIn order by id desc"	
		  elseif request.Form("search")="ok" then
		  sql="select * from Rc_SqlIn where SqlIn_IP='"&request.Form("ipsear")&"'  order by id desc"	
		  end if
		  rs.open sql,conn,1,1
		  
			rs.pagesize=MaxPerPage			 
			If trim(Request("Page"))<>"" then
			   CurrentPage= CLng(request("Page")) 
			   If CurrentPage> rs.PageCount then 
				  CurrentPage = rs.PageCount 
			   End If 
			Else 
			   CurrentPage= 1 
			End If 
		
			totalPut=rs.recordcount
			if CurrentPage<>1 then 
				if (currentPage-1)*MaxPerPage<totalPut then 
					rs.move(currentPage-1)*MaxPerPage 
					bookmark=rs.bookmark
				end if 
			end if
		
			if (totalPut mod MaxPerPage)=0 then  
				n= totalPut \ MaxPerPage
			else  
				n= totalPut \ MaxPerPage + 1  
			end if
		  do while not rs.eof

		  %>
		  <tr bgcolor="<%if i mod 2 then response.write "#ECF9FC" else response.write "#FFFFFF"%>" onMouseOver="this.style.backgroundColor='#C5EBF5'"  onmouseout="this.style.backgroundColor=''" >
            <td height="25" align="center" class="bb01"><%=i+1%></td>
            <td>&nbsp;<%=rs("SqlIn_IP")%></td>
            <td align="center"><%	if rs("Kill_ip")=true then 
			response.write "<font color='red'>已锁定</font>"
		else
			response.write "<font color='green'>已解锁</font>"
		end if
	%></td>
            <td align="center">
 <%	if rs("Kill_ip")=true then 
			response.write "<a href="&URL&"?action=unlock&moduleid="&moduleid&"&id="&rs("id")&">解锁IP</a>"
		else
			response.write "<a href="&URL&"?action=lock&moduleid="&moduleid&"&id="&rs("id")&">锁定IP</a>"
		end if
	%></td>
            <td align="center"><%=rs("SqlIn_WEB")%></td>
            <td align="center"><%=rs("SqlIn_TIME")%></td>
            <td align="center"><%=rs("SqlIn_FS")%></td>
            <td align="center"><%=rs("SqlIn_CS")%></td>
            <td align="center"><%=rs("SqlIn_SJ")%></td>
            <td align="center"><INPUT NAME='idx' TYPE='checkbox' CLASS='radio_all' id="idx" VALUE='<%=rs("id")%>'></td>
          </tr>
		<%
		  i=i+1
		  if i>=MaxPerPage then exit do
		  rs.movenext
		  loop
		  set rs=nothing
		  %>
        </table>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="65%" height="22">
			<%if i>0 then%>分页：
				  第<%=currentpage%>页 共<%=n%>页 共<%=totalPut%>条纪录 
				  <%	k=currentPage
				if k<>1 then
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page=1'>首页</a></b>]&nbsp&nbsp"
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page="+cstr(k-1)+"'>上一页</a></b>]&nbsp&nbsp"
				else
					Response.Write "[首页]&nbsp&nbsp[上一页]&nbsp&nbsp"
				end if
				if k<>n then
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page="+cstr(k+1)+"'>下一页</a></b>]&nbsp&nbsp"
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page="+cstr(n)+"'>尾页</a></b>] "
				else
					Response.Write "[下一页]&nbsp&nbsp[尾页]"
				end if
					 
			set rs=nothing 
			%>
          	&nbsp;选择：<select name="maxpars" onChange="location.href='?moduleid=<%= moduleid%>&maxpars='+this.value">
            <option value="20">显示条数</option>
            <option value="20">前20条</option>
            <option value="40">前40条</option>
            <option value="80">前80条</option>
            <option value="160">前160条</option>
            <option value="100000">全部</option>
          </select>
        <%end if%>
		
	</td><td width="35%" height="22" align="right">
		<%if cDel=true then%>
        <input type="checkbox" name="chkall" onClick="CheckAll(this.form)" value="on"><font color="#FF0000">全选</font><%end if%>
        <%if cDel=true then%>
		 <input type="button" name="Submit22" value="删除" class="submit" onClick="check()">
		 <%end if%></td>
  </tr>
  <input type="hidden" name="moduleid" value="<%= moduleid%>">
  <input type="hidden" name="opt" value="1">
</table>
</form>
<table width="98%" height="23" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
  </tr>
</table>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="7"></td>
  </tr>
</table>  
</body>
</html>




