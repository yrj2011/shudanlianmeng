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
	response.write "<script language='javascript'>alert('����ɾ�������ɹ�!');location.href='"&Request.ServerVariables("HTTP_REFERER")&"';</script>"
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
'��ģ���,��Ա�Ż�ȡ��ǰģ���Ȩ��
'�㷨:��������Ȩ�� 1:���cAdd  2:�޸�cEdit 3:ɾ��cDel 4:���cSh
'���Ȳ�����Ա�Դ�ģ���Ȩ��,���򽫸�Ȩ��ֵΪ��,�����;�ٲ��Ҹ���Ա���ڵĽ�ɫ,�ӽ�ɫ�ҳ���ģ���Ȩ��,������,�����,�����ۼ�


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
		alert("������ѡ��һ����¼���ܲ���!");
		return false;
	}
	if(confirm("�⽫����ѡ��ļ�¼,ȷ��Ҫ���˲�����?"))
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
    <td width="492" height="22" background="images/right_bg1.jpg">&nbsp;��ǰλ�ã�<%=lmwz%>
    </td>
    <td width="507"><div align="left">
      <form name="form2" method="post" action="">
	  <input name="search" type="hidden" value="ok">
        <label>
        ����Ҫ��ѯ��IP :
        <input type="text" name="ipsear">
        </label>
            <label>
            <input type="submit" name="Submit" value="��ѯ">
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
            <td width="5%" align="center">���</td>
            <td width="9%" align="center"><font color=red>�����ɣ�</font></td>
            <td width="7%" align="center">�Ƿ�����</td>
            <td width="6%" align="center">����</td>
            <td width="19%" align="center">����ҳ��</td>
            <td width="16%" align="center">����ʱ��</td>
            <td width="8%" align="center">�ύ��ʽ</td>
            <td width="9%" align="center">�ύ����</td>
            <td width="16%" align="center">�ύ����</td>
            <td width="5%" align="center">����</td>
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
			response.write "<font color='red'>������</font>"
		else
			response.write "<font color='green'>�ѽ���</font>"
		end if
	%></td>
            <td align="center">
 <%	if rs("Kill_ip")=true then 
			response.write "<a href="&URL&"?action=unlock&moduleid="&moduleid&"&id="&rs("id")&">����IP</a>"
		else
			response.write "<a href="&URL&"?action=lock&moduleid="&moduleid&"&id="&rs("id")&">����IP</a>"
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
			<%if i>0 then%>��ҳ��
				  ��<%=currentpage%>ҳ ��<%=n%>ҳ ��<%=totalPut%>����¼ 
				  <%	k=currentPage
				if k<>1 then
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page=1'>��ҳ</a></b>]&nbsp&nbsp"
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page="+cstr(k-1)+"'>��һҳ</a></b>]&nbsp&nbsp"
				else
					Response.Write "[��ҳ]&nbsp&nbsp[��һҳ]&nbsp&nbsp"
				end if
				if k<>n then
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page="+cstr(k+1)+"'>��һҳ</a></b>]&nbsp&nbsp"
					response.write "[<b>"+"<a href='?moduleid="&moduleid&"&page="+cstr(n)+"'>βҳ</a></b>] "
				else
					Response.Write "[��һҳ]&nbsp&nbsp[βҳ]"
				end if
					 
			set rs=nothing 
			%>
          	&nbsp;ѡ��<select name="maxpars" onChange="location.href='?moduleid=<%= moduleid%>&maxpars='+this.value">
            <option value="20">��ʾ����</option>
            <option value="20">ǰ20��</option>
            <option value="40">ǰ40��</option>
            <option value="80">ǰ80��</option>
            <option value="160">ǰ160��</option>
            <option value="100000">ȫ��</option>
          </select>
        <%end if%>
		
	</td><td width="35%" height="22" align="right">
		<%if cDel=true then%>
        <input type="checkbox" name="chkall" onClick="CheckAll(this.form)" value="on"><font color="#FF0000">ȫѡ</font><%end if%>
        <%if cDel=true then%>
		 <input type="button" name="Submit22" value="ɾ��" class="submit" onClick="check()">
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




