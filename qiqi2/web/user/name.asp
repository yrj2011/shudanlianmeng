<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../background.asp"-->
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
heiname=request("heiname")
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
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
 <DIV class=pp9>
      <DIV style="PADDING-BOTTOM: 15px; WIDTH: 97%">
      <DIV class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;官方黑名单 &gt;&gt; </DIV>
      <DIV class=pp8><STRONG>官方黑名单</STRONG></DIV>
      <DIV 
style=" MARGIN-LEFT: 15px; CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 10px; MARGIN-bottom: 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 700px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #f3f8fe">
<DIV 
style="CLEAR: both; MARGIN-TOP: 10px; MARGIN-LEFT: 5px; LINE-HEIGHT: 200%; TEXT-ALIGN: left">
<UL>
    <%
	 	Sql = "select  name ,count(*) as yu from  "&jieducm&"_hei   group by name"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>
  <LI 
  style="LIST-STYLE-POSITION: inside; LIST-STYLE-IMAGE: url(../images/list_bg.gif); BORDER-BOTTOM: #999999 1px dashed"><%=rs("name")%>(已有<%=rs("yu")%>人列其为黑名单) 
   </LI>
    <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
</UL>
</DIV></DIV>
		
	<DIV 
style=" MARGIN-left: 7px; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 20px; FLOAT: left; BORDER-LEFT: #abbec8 1px solid; WIDTH: 320px; LINE-HEIGHT: 150%; BORDER-BOTTOM: #abbec8 1px solid; HEIGHT: 260px; TEXT-ALIGN: center">
<DIV style="CLEAR: both; MARGIN-TOP: 10px; MARGIN-LEFT: 5px; LINE-HEIGHT: 200%; TEXT-ALIGN: left">
<UL>


  <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
    <TBODY>
        <%
			  	Sql = "select * from "&jieducm&"_hei where username='"&session("useridname")&"' order by ID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
				i=1
					Do While Not Rs.Eof
					i=i+1
			  %>  <FORM id="aForm<%=i%>" name="aForm<%=i%>" action="" method=post>

    <TR>
      <TD width=120><IMG  src="../images/index1_42.jpg"  ><%=rs("name")%></TD>
      <TD><INPUT id=ModBtn1 type=submit value=撤销黑名单 name=ModBtn1><input name="action" type="hidden" value="del">
        <input type="hidden" name="id" id="id"  value="<%=rs("id")%>">      </TD>
      </TR> </FORM>
      <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
      </TBODY></TABLE>
     
      </UL></DIV></DIV>	
	  
	  
	  
	  <DIV 
style="BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 20px; FLOAT: right; BORDER-LEFT: #abbec8 1px solid; WIDTH: 350px; LINE-HEIGHT: 150%; BORDER-BOTTOM: #abbec8 1px solid; HEIGHT: 260px">
<DIV style="WIDTH: 320px; LINE-HEIGHT: 50px; HEIGHT: 50px">你要将谁拉入黑名单(对方用户名)： 
</DIV>
<FORM id=aspnetForm name=aspnetForm action="" method=post>
<DIV style="WIDTH: 320px; LINE-HEIGHT: 80px; HEIGHT: 80px">
<INPUT id=name name=name value="<%=heiname%>"><input name="action" type="hidden" value="ok">
<INPUT id=button1 type=submit value=加入我的黑名单 name=button1></DIV></FORM>
<DIV 
style="WIDTH: 320px; COLOR: red; LINE-HEIGHT: 50px; HEIGHT: 60px"></DIV></DIV>
      </DIV>
 </DIV>
 
 
 </td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
</BODY></HTML>
<%
action=HtmlEncode(trim(request.form("action")))
name1=HtmlEncode(trim(request.form("name")))
if action="ok" then
	Sql = "select * from "&jieducm&"_register where username='"&name1&"'"
	Set Rsm = Server.CreateObject("Adodb.RecordSet")
	Rsm.Open Sql,conn,1,1
	IF Rsm.Eof Then
	   Response.Write("<script>alert('对不起无此用户!');window.location.href='name.asp';</script>")
	   response.end()
	else
	    if name1=session("useridname") then	
		  response.write("<script>alert('对不起,不能添加自己。');window.location.href='name.asp';</script>")
		  response.End()
		end if
 	    Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select * From "&jieducm&"_hei where name='"&name1&"' and username='"&session("useridname")&"'" ,Conn,3,3 
		if not rs.eof then
		  response.write("<script>alert('对不起,此用户已在你的黑名单中,不能重复加入');window.location.href='name.asp';</script>")
		  response.End()
		else 
	    rs.addnew
		rs("name")=name1
    	rs("username")=session("useridname")	
		rs("class")=1
    	rs.update
    	rs.close
		response.Write "<script language='javascript'>alert(' 操作成功！'); window.location.href='name.asp';  </script>"
		response.End()
		end if
	end if		
	
elseif action="del" then
id=request("id")
sql="delete from "&jieducm&"_hei where id in ("&id&") and username='"&session("useridname")&"'"
conn.execute(sql)  
response.Write "<script language='javascript'>alert(' 撤销成功！'); window.location.href='name.asp';  </script>"
response.End()
end if
call  CloseConn()
%>