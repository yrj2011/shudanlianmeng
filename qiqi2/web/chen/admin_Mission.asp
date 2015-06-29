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
<LINK href="../images/mission.css" type=text/css rel=stylesheet>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
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
		document.myform.action="admin_myssiondel.asp";
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
style="CLEAR: both; MARGIN-BOTTOM: 10px; MARGIN-LEFT: 1px;MARGIN-TOP: 10px; COLOR: #ff0000">&nbsp;&nbsp;本功能只能删除已完成的任务&nbsp;&nbsp; </DIV>
</DIV>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">选择任务区：<a href="?action=1">淘宝普通区</a> |<a href="?action=2">拍拍区</a> | <a href="?action=4">充值提现区</a></div></td>
    </tr>
          <tr>
            <td height=12>&nbsp;</td>
          </tr>
          </td>
</table>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<UL class=mission>
  <LI style="WIDTH: 20%">任务编号 </LI>
  <LI style="WIDTH: 10%">发布人 </LI>
  <LI style="WIDTH: 30%">发布时间</LI>
  <LI style="WIDTH: 10%"> 任务价格</LI>
  <LI style="WIDTH: 20%">任务区</LI>
  </UL></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<form name="myform" method="Post" action="admin_myssiondel.asp" onSubmit="return ConfirmDel();">
<%	
action=request("action")
if action="" then
sql="select * from "&jieducm&"_zhongxin where classid<>'6' and start='5' order by id asc"
elseif action="1" then
sql="select * from "&jieducm&"_zhongxin where classid=1 and start='5' order by id "
elseif action="2" then
sql="select * from "&jieducm&"_zhongxin where classid=2 and start='5' order by id "
elseif action="3" then
sql="select * from "&jieducm&"_zhongxin where classid=3 and start='5' order by id "
elseif action="4" then
sql="select * from "&jieducm&"_zhongxin where classid=5 and start='5' order by id "
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then					
	dim maxperpage,url,PageNo
	if action="" then
	 url="admin_mission.asp"
	else
	 url="admin_mission.asp?action="&action&""
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
 DO WHILE NOT rs.EOF AND RowCount>0

 %>
<DIV style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 50px">
<UL class=missionitem>
  <LI style="WIDTH: 20%">
  
  <input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>><%=rs("pid")%> </LI>
  <LI style="WIDTH: 10%"><%=rs("username")%> </LI>
  <LI style="WIDTH: 30%"><%=rs("now")%></LI>
  <LI style="WIDTH: 10%"><%=rs("price")%></LI>
  <LI style="WIDTH: 20%"> 
  <%if rs("classid")=1 then
    response.Write("淘宝普通区任务")
  elseif rs("classid")=2 then
   response.Write("拍拍区任务")
  elseif rs("classid")=3 then 
  response.Write("有啊区任务")
  elseif rs("classid")=5 then
  response.Write("充值提现区任务")
  end if
  %> </LI>
  </UL>
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

</BODY></HTML>
