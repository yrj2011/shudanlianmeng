<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
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
id=request("id")
if request.Form<>"" then
	 code=HtmlEncode(trim(request.form("code")))
	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='" &session("useridname")& "' and czm='"&code&"'" ,Conn,3,3  
				IF Rs.Eof Then
					Response.Write("<script>alert('操作码不正确!');</script>")
				Else
					 session("czm")=rs("czm")
					 if id="ti" then
					   response.Redirect("remoney.asp")
					 else
    	              response.redirect("PushMission.asp")
					 end if
				end if
set rs=nothing

end if
%>	
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<SCRIPT language=javascript>
function  save_onclick()
{

    var qq1=document.form.code.value;
  if(qq1=="")
  {
    alert("请输入您的操作码！");
	document.form.code.focus();
	return false;
	}
	return true;
	}
	</script>
<!-- 左隐藏菜单结束 --><!-- 左隐藏菜单BIG -->
<div align="center">
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 15px; BACKGROUND-COLOR: #ffffff">
<FORM action="" method="POST" name="form" id="form" onSubmit="return save_onclick()">
<DIV></DIV>

<DIV style="WIDTH: 910px; TEXT-ALIGN: center">
<DIV 
style="PADDING-RIGHT: 20px; PADDING-LEFT: 20px; PADDING-BOTTOM: 20px; WIDTH: 700px; PADDING-TOP: 20px; TEXT-ALIGN: left">
<DIV style="FONT-WEIGHT: bolder; FONT-SIZE: 18px; COLOR: red">输入操作码</DIV><BR>
<DIV style="COLOR: #999900"><FONT 
color=#ff0000>你必须输入操作码才能继续操作，每次登陆都要输入一次才能发布任务，提现，赠送发布点、修改个人资料</FONT></DIV><BR>
<DIV style="COLOR: red"><SPAN id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="COLOR: red"></SPAN></DIV><BR>
<DIV>
<P>操作码：</P><INPUT id=code type=password 
maxLength=16 name=code><SPAN 
id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator1 
style="DISPLAY: none; COLOR: red">* 操作码不能为空！</SPAN></DIV><BR>
<DIV><INPUT id=ctl00_AllBaseContentPlaceHolder_button1  type=submit value="  提交  " name=ctl00$AllBaseContentPlaceHolder$button1></DIV><BR><BR></DIV></DIV>


<DIV> </DIV>



</FORM></DIV></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
