<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK 
href="../images/mission.css" type=text/css rel=stylesheet>
<LINK 
href="../images/Default.css" 
type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<title>七七传媒</title>
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
		document.myform.action="admin_recordmdel.asp";
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
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   

<DIV 
style="CLEAR: both; OVERFLOW: hidden; WIDTH: 950px; BACKGROUND-COLOR: #f3f8fe">
  <DIV 
style="CLEAR: both; MARGIN-BOTTOM: 10px; MARGIN-LEFT: 1px;MARGIN-TOP: 10px; COLOR: #ff0000">&nbsp;&nbsp;选择下列操作记录方式： 
* <a href="admin_Recordm.asp">所有操作记录列表</a> * <a href="admin_Recordm.asp?action=manage">任务操作列表</a> * <a href="admin_Recordm.asp?action=chong">充值操作列表</a> * <a href="admin_Recordm.asp?action=login">登录后台列表</a></DIV>
</DIV>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="admin_recordm.asp?action=ok">
                  
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>搜索管理员或管理操作：
                  <input class=input1 size=25 name=text>
　
<input name="submit" type=submit class=input2 value=" 搜 索 ">
                </form> 
            </div></td>
    </tr>
          <tr>
            <td height=12>&nbsp;</td>
          </tr>
          </td>
</table>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<UL class=mission>
  <LI style="WIDTH: 20%">用户名 </LI>
  <LI style="WIDTH: 10%">类型 </LI>
  <LI style="WIDTH: 40%">详细 </LI>

  <LI style="WIDTH: 10%">结果 </LI>
  <LI style="WIDTH: 18%">时间 </LI></UL></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<form name="myform" method="Post" action="admin_recordmdel.asp" onSubmit="return ConfirmDel();">
<%	
action=request("action")

if action="" then
sql="select * from "&jieducm&"_recordm  order by id desc"
elseif action="chong" then
sql="select * from "&jieducm&"_recordm where   class='任务' or class='充值' or class='增值' or class='提现' or class='其它' order by id desc"
elseif action="manage" then
sql="select * from "&jieducm&"_record where class='清理买家' or class='清理删除' or class='返回上级' or class='确认发货' or class='买方好评' or class='卖方完成' order by id desc"
elseif action="login" then
sql="select * from "&jieducm&"_recordm where   class='登录后台' order by id desc"

elseif action="ok" then
sql="select * from "&jieducm&"_recordm where  username = '"&request("text")&"' or class like '%"&request("text")&"%' order by id desc"
end if

			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	if action="" then
	   url="admin_recordm.asp"
	elseif action="ren" then
	   url="admin_recordm.asp?action=ren"
    elseif action="chong" then
	  	 url="admin_recordm.asp?action=chong"
    elseif action="zeng" then
	 	 url="admin_recordm.asp?action=zeng"
    elseif action="login" then
		 url="admin_recordm.asp?action=login"
   elseif action="manage" then
    	 url="admin_recordm.asp?action=manage"
   elseif action="ok" then
   	 url="admin_recordm.asp?action="&request("text")&""
  end if
	 if session("webid")<>"" then url=url&"&username="&session("webid")
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
<DIV 
style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 50px">
<UL class=missionitem>
 
  <LI style="WIDTH: 20%"><input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>><%=rs("username")%> </LI>
  <LI style="WIDTH: 10%"><%=rs("class")%> </LI>
  <LI style="WIDTH: 40%"><%=rs("content")%></LI>
  <LI style="WIDTH: 10%"><%=rs("jiegou")%> </LI>
  <LI style="WIDTH: 18%"><%=rs("now")%> </LI></UL>
</DIV>
   <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></DIV>
<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
              选中本页显示的所有记录
              <input name="submit" type='submit' value='&nbsp;删除选定的记录&nbsp;' onClick="document.myform.Action.value='Del'"  style="cursor: hand;background-color: #cccccc;">
              <input name="Action" type="hidden" id="Action" value="Del"></form>
&nbsp;&nbsp;&nbsp;

</DIV>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
