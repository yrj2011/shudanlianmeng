<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
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
set Rs=server.createobject("adodb.recordset")
sql="select cang from "&jieducm&"_stup "
Rs.open sql,conn,1,1
if not(Rs.eof)  then
 stupcang=rs("cang")
end if
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
<SCRIPT src="../js/jieducm_Dialog.js" type=text/javascript></SCRIPT>
<title>无标题文档</title>

<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="admin_MyMissionjie5.asp">
                  
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>搜索任务ID或用户名：
                  <input class=input1 size=25 name=text value="<%=request.Form("text")%>">
　
<input name="sort" type="hidden" value="ok" />
<input name="submit" type=submit class=input2 value=" 搜 索 ">
                </form> 
            </div></td>
  </tr>
          <tr>
            <td height=12>&nbsp;</td>
          </tr>
          </td>
</table>
<DIV style="CLEAR: both; OVERFLOW: hidden; WIDTH: 950px; BACKGROUND-COLOR: #f3f8fe"><IMG 
src="images/arrow2.gif">&nbsp;<SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 18px; COLOR: red">已接收藏区任务</SPAN>&nbsp; 
* <A href="admin_MyMissionjie5.asp">所有任务</A>
*<A href="admin_MyMissionjie5.asp?sort=0">等待收藏的任务</A>
*<A href="admin_MyMissionjie5.asp?sort=1">已经完成</A></DIV>

<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<UL class=mission>
  <LI style="WIDTH: 15%">发布方任务ID </LI>
  <LI style="WIDTH: 10%">价格 </LI>
  <LI style="WIDTH: 15%">商品信息 </LI>
  <LI style="WIDTH: 15%">接手方 </LI>
  <LI style="WIDTH: 15%">接手时间 </LI>
  <LI style="WIDTH: 15%">状&nbsp;&nbsp;态</LI>
   <LI style="WIDTH: 15%">操&nbsp;&nbsp;作</LI>
  </UL></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<%
action=request("sort")
if action="" then
	Sql = "select * from "&jieducm&"_jieshou where  classid='6'  order by id desc"
elseif action="0" then 
	 Sql = "select * from "&jieducm&"_jieshou where  start='1'   classid='6' order by id desc"
elseif action="1" then
	 Sql = "select * from "&jieducm&"_jieshou where start='2'   and  classid='6' order by id desc"
elseif action="ok" then
 sql="select * from "&jieducm&"_jieshou where  classid='6' and (pid like '%"&request("text")&"%' or  username like '%"&request("text")&"%')   order by id desc"				
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	dim maxperpage,url,PageNo
	if action="" then
	 url="admin_MyMissionjie5.asp"
	elseif action="ok" then
	url="admin_MyMissionjie5.asp?sort=ok&text="&request("text")&""
	else
	 url="admin_MyMissionjie5.asp?sort="&action&""
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
 usernamef=rs("username2")
 pid=rs("pid")
 bei=rs("bei")
 usernamej=rs("username")
 start=rs("start")
 id=rs("id")
nowj=rs("now")
num=rs("num")
Sql2 = "select * from "&jieducm&"_zhongxin where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then
nowf=rs2("now")
usernamef=rs2("username")
tips=rs2("tips")
bei=rs2("bei")
end if
 %>
 <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%">
<DIV style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 80px">
<UL class=missionitem>
 
  <LI style="WIDTH: 15%"><a href="admin_recordu.asp?username=<%=usernamef%>"><%=usernamef%></a></font><br />
<%=pid%> <br /><%=bei%></LI>
  <LI style="WIDTH: 10%"><font color=#ff6600>平台担保<br>
<img src="../images/jieducm_zf.gif" width="13" height="16"><%=stupcang%>个发布点</font></LI>
  <LI style="WIDTH: 15%"><input name="<%=pid%>"  type="hidden" id="<%=pid%>" size="9" value="<%=rs("prourl")%>">
<INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=复制商品地址> <BR>
掌柜：<%=rs("Shopkeeper")%>
<br />
 <%
			  	Sqlu = "select qq from "&jieducm&"_register where username='"&usernamef&"'"
				Set Rsu = Server.CreateObject("Adodb.RecordSet")
				Rsu.Open Sqlu,conn,1,1
				if not rsu.eof then
				qq=rsu("qq")
				rsu.close
				end if
			  %>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes"><img height=16 alt=七七互刷平台欢迎您 src="../images/qq1.gif" width=23 border=0>
</a>
</span></LI>
  <LI style="WIDTH: 15%">
  <%
			  	Sqlu = "select username,qq from "&jieducm&"_register where username='"&usernamej&"'"
				Set Rsu = Server.CreateObject("Adodb.RecordSet")
				Rsu.Open Sqlu,conn,1,1
				if not rsu.eof then
				qqj=rsu("qq")
				rsu.close
				end if
			  %>
<a href="admin_recordu.asp?username=<%=usernamej%>"> <%=usernamej%></a><br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qqj%>&site=qq&menu=yes"><img height=16 alt=七七互刷平台欢迎您 src="../images/qq1.gif" width=23 border=0> </a>

</LI>
  <LI style="WIDTH: 15%">
  <%=nowj%>	<br />接手淘宝号：<font color="#FF0000"><%=num%></font>
   </LI>
  <LI style="WIDTH: 15%">
  <%if start=1 then%>
  等待收藏
  <%elseif start=2 then%>
  完成收藏
  <%end if%>
  </LI>
   <LI style="WIDTH: 15%">
   <%if start=1 then%>
      <a href="delmy5.asp?action=qingli&id=<%=id%>" onClick="return confirm('确定要删除清理买家吗？');">清理买家</a>
<br /><br />
   <a href="delmy5.asp?action=ok&id=<%=id%>" onClick="return confirm('确定要后台强制完成收藏吗？');">后台强制完成收藏</a>
   <%end if%>
   </LI>
  </UL></DIV></td>
  </tr>
</table>
   <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> 
<% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></DIV>


</DIV>
<%
set rsu=nothing
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
