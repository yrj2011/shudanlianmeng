<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>七七传媒</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>

<body>
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

	if(document.myform.Action.value=="Del")
	{
		document.myform.action="shendel.asp";
		if(confirm("确定要删除选中的用户吗？本操作将把无法恢复！"))
		    return true;
		else
			return false;
	}
	else
	{
		document.myform.action="shendel.asp";
		if(document.myform.TargetClassID.value=="")
		{
			alert("不能将文章移动到含有子栏目的栏目或外部栏目中！");
			return false;
		}
		if(confirm("确定要将选中的文章移动到指定的栏目吗？"))
		    return true;
		else
			return false;
	}
}

</SCRIPT>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong>申 诉 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>管理导航：</strong></td>
    <td>&nbsp;</td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
    <form name="myform" method="Post" action="shendel.asp" onSubmit="return ConfirmDel();">
      <table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
        <tr>
          <td colspan="2" align="left" class="title"></td>
        </tr>
        <tr valign="middle" class="tdbg">
          <td height="22"></td>
          <td width="200" height="22" align="right"></td>
        </tr>
      </table>
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22"> 
            <td height="22" width="59" align="center"><strong>选中</strong></td>
            <td width="79" align="center" ><strong>申诉人</strong></td>
            <td width="99" align="center" ><strong>标题</strong></td>
            <td width="52" align="center" ><strong>被申诉人</strong></td>
            <td width="84" align="center" ><strong>任务ID </strong></td>
            <td width="105" align="center" ><strong>申诉时间</strong></td>
            <td width="116" align="center" ><strong>被申诉人是否辩解</strong></td>
            <td width="116" align="center" ><strong>申诉人是否辩解</strong></td>
            <td width="125" align="center" ><strong>官方意见</strong></td>
            <td width="147" align="center" ><strong>操作</strong></td>
            <!--<td width="40" align="center" ><strong>已审核</strong></td>-->
          </tr>
       
			<%	
			action=request("action")
		   if action="sear" then
		        sql="select * from "&jieducm&"_toushu where pid='"&request("text")&"' "
			elseif action="" then
			   	sql="SELECT * FROM "&jieducm&"_toushu order by id desc"
          end if

			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	 url="shenshu.asp"
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
	
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	  %>
       <% DO WHILE NOT rs.EOF AND RowCount>0%>
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
            <td width="59" align="center"><input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>></td>
            <td width="79" align="center"><%=rs("username")%></td>
            <td width="99" align="center"><%=rs("title")%></td>
            <td width="52" align="center"><%=rs("name")%></td>        
            <td width="84" align="center"><%=rs("pid")%></td>
            <td width="105" align="center"><%=rs("now")%></td>
            <td width="116" align="center"><%if rs("bianjie2") ="否" then%>否<%else%>是<% end if%></td>
            <td width="116" align="center"><%if rs("bianjie") ="否" then%>否<%else%>是<% end if%></td>
            <td width="125" align="center"><%if rs("guang") ="否" then%>否<%else%>是<% end if%></td>
            <!--<td width="40" align="center"> 
			'if rsArticleList("Passed")=true then
				'response.write "是"
			'else
				'response.write "否"
			'end if%></td>-->
            <td width="147" align="center"><a href="shenshug.asp?id=<%=rs("id")%>">官方意见</a>&nbsp; <a href="shendel.asp?id=<%=rs("id")%>">删除</a><br></td>   
          </tr>
            <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          </table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td height="30" colspan="2"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
              选中本页显示的所有用户 
              <input name="submit" type='submit' value='&nbsp;删除选定的用户&nbsp;' onClick="document.myform.Action.value='Del'"  style="cursor: hand;background-color: #cccccc;">
              <input name="Action" type="hidden" id="Action" value="Del">
&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td  ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %>
  </div></td></tr>
</table>
</td>
</form></tr>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="?action=sear">
                  <%  
	  set guest=nothing 
      set rs = nothing%>
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle title="捷 度 传媒">搜索任务ID：
                  <input 
      class=input1 size=25 name=text>
　
<input name="submit" type=submit class=input2 value=" 搜 索 ">
                </form> 
            </div></td>
    </tr>
          <tr>
            <td height=12>&nbsp;</td>
          </tr>
          </td>
</table>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com/' target="_blank"><font color=#ff6600><strong>七七传媒-七七互刷平台</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
<%

set rs=nothing
conn.close
set conn=nothing%>    

</body>
</html>
