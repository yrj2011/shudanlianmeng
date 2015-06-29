<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->

<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统**************************************
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
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<style type="text/css" media="screen">
	
	@import url("../css/interiormail.css");
</style>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
 
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>



<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 私人站内信 &gt;&gt; </div>
                    <div class=pp8><strong>私人站内信</strong></div>
               
					  <FORM name=form method=post>
           
           
            <%

dim sql,rs
	
keyword=replace(trim(request.form("keyword")),"'","")
if keyword <>"" then
sql="select * from "&jieducm&"_china_message where (uid='"&session("useridname")&"' ) and zn='yes' and ( messagename like '%" & keyword & "%' or  messagelxfs like '%" & keyword & "%' or messagetext like '%" & keyword & "%' ) order by id desc"
else
sql="select * from "&jieducm&"_china_message where (uid='"&session("useridname")&"' ) and zn='yes' order by id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3%>
            <SCRIPT language=javascript>
function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
function Checked()
{
	var j = 0
	for(i=0;i < document.form.elements.length;i++){
		if(document.form.elements[i].name == "adid"){
			if(document.form.elements[i].checked){
				j++;
			}
		}
	}
	return j;
}

function DelAll()
{
	if(Checked()  <= 0){
		alert("您至少选择1条消息!");
	}	
	else{
		if(confirm("确定要删除选择的消息吗？\n此操作不可以恢复！")){
			form.action="delm.asp?del=message";
			form.submit();
		}
	}
}

</SCRIPT>
          <br>
				<a href="send_message.asp"><img src="../img/jieducm_xin_img.jpg" width="95" height="24" border="0"></a><br>
				<br>

              			  <div id="TopBox">您现在查看的是： </div>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
              		<tr>
		<td colspan="8">
			<ul class="private-msg-tab">
				<li ><a href="user.asp">站内信</a></li>
                <li class="selected"><a href="users.asp ">私人站内信</a></li>
				</ul>		</td>
               
	</tr>

                <tr>
                  <th scope="col" width="9%"align="center"><div align="left">
                    <input id=chkAll 
                  onClick=CheckAll(this.form) type=checkbox 
                  value=checkbox name=chkAll>
                  </div></th>
                  <th scope="col" width="10%"align="center"><div align="left">编号</div></th>
                  <th scope="col" width="15%" align="center"><div align="left">发送人</div></th>
                  <th scope="col" width="31%" align="center"><div align="left">标题</div></th>
                  <th scope="col" width="13%" ><div align="left">接收人</div></th>
                  <th scope="col" width="22%" align="center"><div align="left">发送时间</div></th>
                </tr>
                 <%if rs.eof or rs.bof then
response.write"<font color='red'>暂无相关消息！</font>"
'response.end
else
const maxperpage=20
dim currentpage
rs.pagesize=maxperpage
currentpage=request.querystring("pageid")
if isnumeric(currentpage)=false then
response.write "<script>alert('参数错误，关闭窗口！');window.close();</script>"
response.end
end if
if currentpage="" then
currentpage=1
elseif currentpage<1 then
currentpage=1
else
currentpage=clng(currentpage)
	if currentpage > rs.pagecount then
	currentpage=rs.pagecount
	end if
end if

dim totalput,n
totalput=rs.recordcount
if totalput mod maxperpage=0 then
n=totalput\maxperpage
else
n=totalput\maxperpage+1
end if
if n=0 then
n=1
end if
rs.move(currentpage-1)*maxperpage
i=0
do while i< maxperpage and not rs.eof
messagetext=rs("messagetext")

%>
                	    				    			<tr>
                	    				    			  <td><%if rs("uid")="all" then%>
                <b><font color="#FF0000">!</font></b><%else%><input type="checkbox" name="adid" value="<%=rs("id")%>" onClick=Checked(form)><%end if%></td>
	    				<td><%=i+currentpage*maxperpage-maxperpage+1%></td>					
	    				<td><%=rs("messagename")%></td>
	    				<td>
                           <a href="zn_messagelook.asp?id=<%=rs("id")%>" ><% if rs("hits")=0 then%><font color="#000000" size="2 px;"><%else%> <font color="#0066CC"><%end if%><strong><%=rs("messagelxfs")%></strong></font></a>
                        </td>
	    				<td> <%if rs("uid")="all" then%>
                全部会员 
                <%else%>
                <%=rs("uid")%> 
                <%end if%></td>
	    				<td><%=rs("messagedate")%></td>
	    			</tr>
                    
						 <%i=i+1                     
    rs.movenext                                                               
    loop    
                                                                                       
    end if%>							               
                         <tr>
                           <td colspan="8" class="w">&nbsp; <div class="PageButton">
				  	<INPUT title=删除 onclick=DelAll() type=button value=删除 name=Submit>
&nbsp;&nbsp;				  </div>	</td>
                         </tr>
                <tr>
                  <td colspan="8" class="w">				 			  </td>
                </tr>
                
                              <tr>
                  <td colspan="8">
		  <!--AdForward Begin:-->
		  <!--AdForward End--></td> 
		                  </tr>
              </table>
      </form>
               <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="CAE2F8">  
   <tr>                                                     
    <td height="28" bgcolor="#FFFFFF">
      <p align="center">页数：<%=currentpage%>/<%=n%> 
        <%k=currentpage                                                                                         
   	if k<>1 then%>
        <a href="?pageid=1">首页</a> <a href="?pageid=<%=k-1%>">上一</a> 
        <%else%>
        首页&nbsp;上一页 
        <%end if%>
        <%if k<>n then%>
        <a href="?pageid=<%=k+1%>">下一</a> <a href="?pageid=<%=n%>">尾页</a> 
        <%else%>
        下一页&nbsp;尾页 
        <%end if%>
        共有 <%=totalput%> 条消息 
    </td>                                                                                  
    <form action="" method="post" name="search"><td width="240" align="center" bgcolor="#FFFFFF">关键字
      <input  maxLength="20" name="keyword" onFocus="this.value=''" size="18"  value="<%=keyword%>"> 
    <input  type="submit" value="搜索" style="font-size: 12px" name="search"></td>
    </form>
   </tr>                                                                
</table>   </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
