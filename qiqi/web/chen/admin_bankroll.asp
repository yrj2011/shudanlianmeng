<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>发言排名</title>
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
		document.myform.action="tidel.asp";
		if(confirm("确定要删除选中的记录吗？本操作将把无法恢复！"))
		    return true;
		else
			return false;
	}
	else
	{
		document.myform.action="chongdel.asp";
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
	<td colspan="2" align="center" class="title"><strong>服 务 卡 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>管理导航：</strong></td>
    <td><a href="?action=0">等待处理</a> <a href="?action=2">已撤销</a> <a href="?action=3">锁定撤销</a>&nbsp;<a href="?action=1">提现完成</a></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
    <form name="myform" method="Post" action="tidel.asp" onSubmit="return ConfirmDel();">
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
            <td width="151" align="center"  height="22"><strong>流水号</strong></td>
            <td width="163" align="center" ><strong>提现类型</strong></td>
            <td width="163" align="center" ><strong>提现用户名</strong></td>
            <td width="123" align="center" ><strong>提现金额</strong></td>
            <td width="113" align="center" ><strong>真实姓名</strong></td>
            <td width="139" align="center" ><strong>账号</strong></td>
            <td width="161" align="center" ><strong>状态</strong></td>
            <td width="161" align="center" ><strong>提现时间</strong></td>
            <td width="161" align="center" ><strong>操作</strong></td>
            <!--<td width="40" align="center" ><strong>已审核</strong></td>-->
          </tr>
       
			<%	
action=request("action")
if action="" then
   	sql="SELECT * FROM "&jieducm&"_tixian  order by id desc"
elseif action="ok" then
  sql="SELECT * FROM "&jieducm&"_tixian where username ='"&request("text")&"'  order by id desc"
elseif action="0" then
   	sql="SELECT * FROM "&jieducm&"_tixian  where start='0' order by id desc"
elseif action="1" then
   	sql="SELECT * FROM "&jieducm&"_tixian  where start='1' order by id desc"
elseif action="2" then
    	sql="SELECT * FROM "&jieducm&"_tixian  where start='2' order by id desc"
elseif action="3" then
    	sql="SELECT * FROM "&jieducm&"_tixian  where start='3' order by id desc"

end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	dim maxperpage,url,PageNo
	if action="" then
	 url="admin_bankroll.asp"
	elseif action="ok" then
		 url="admin_bankroll.asp?action=ok&text="&request("text")&""
	else
		 url="admin_bankroll.asp?action="&action&""
	end if
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
	   DO WHILE NOT rs.EOF AND RowCount>0%>
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
            <td width="151" align="center"><%=rs("pid")%></td>
            <td width="163" align="center"><%if rs("class")=1 then%>淘宝商品地址提现
			<%elseif rs("class")=2 then%>支付宝提现
			<%elseif rs("class")=3 then%>财富通提现
			<%end if%></td>
            <td width="163" align="center"> <%=rs("username")%></td>
            <td width="123" align="center"><%=rs("ReNum")%>元</td>
            <td width="113" align="center"><%=rs("ReRName")%></td>        
            <td width="139" align="center"><%=rs("ReZfb")%></td>
            <td width="161" align="center">
			<%
			if(rs("start")=0) then
			   response.write("<font color=red>处理中</font>")
			elseif (rs("start")=1 )then
			   response.write("提现成功")
			   response.Write("<br>交易号：")
			   response.Write(rs("zfbnum"))
			 elseif (rs("start")=2 )then
			   response.write("撤消提现")
			 elseif (rs("start")=3 ) then
			    response.write("已锁定")
			 end if
			%>            </td>
            <td width="161" align="center"><%=rs("now")%></td>
            <td width="161" align="center">
           <% if rs("start")<>2  and rs("start")<>1 then%><a href="tixianok.asp?id=<%=rs("id")%>&action=rem"> 撤消提现</a> <% else%><font color="#999999">撤消提现</font><% end if%>
           | <% if rs("start")<>3  and rs("start")<> 1 and rs("start")<>2 then%><a href="tixianok.asp?id=<%=rs("id")%>&action=shou">锁定用户撤消</a><% else%><font color="#999999">锁定用户撤消</font><% end if%> <br> 
           
           <% if rs("start")<>1 and rs("start")<>2 then%>
            <a href="tixianjiao.asp?id=<%=rs("id")%>">提现完成</a>
            <%else%>
            <font color="#999999">提现完成</font>
            <% end if%>          </td>
            <!--<td width="40" align="center"> 
			'if rsArticleList("Passed")=true then
				'response.write "是"
			'else
				'response.write "否"
			'end if%></td>-->
          </tr>
            <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          </table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td height="30" colspan="2"><input name="Action" type="hidden" id="Action" value="Del">
&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %>
  </div></td></tr>
</table>
</td>
</form></tr>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="?action=ok">
                  <%  
	  set guest=nothing 
      set rs = nothing%>
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>搜索用户名：
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
<%
set rs=nothing
conn.close
set conn=nothing%>

</body>
</html>
