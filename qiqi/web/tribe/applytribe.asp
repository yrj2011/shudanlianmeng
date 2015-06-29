<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="../user/checklogin.asp"-->
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
action=request("action")
if action="ok" then    
	 tribename2=HtmlEncode(request.Form("tribename"))	 
	 pic=request.Form("pic")
	 tribeinfo=Request.Form("tribeinfo")
     triberad=request.Form("triberad")
if jifei<5000 or vip<>1 then
	  Response.Write("<script>alert('积分达到5000以上或尊贵vip会员才有权限建部落 !');window.location.href='index.asp';</script>")
	  response.End()
end if	 
	Sql= "select * from "&jieducm&"_tribe where username='"&session("useridname")&"'"
	Set rs=server.createobject("ADODB.RECORDSET")
	Rs.Open Sql,conn,3,3
	if not rs.eof then
	  Response.Write("<script>alert('你已经创建了一个部落!');window.location.href='index.asp';</script>")
	  response.End()
	else			
		rs.addnew
		rs("username")=session("useridname")
		rs("tribename")=tribename2
    	rs("pic")=pic
		rs("tribeinfo")=tribeinfo
		rs("triberad")=triberad
		rs("now")=now()
    	rs.update
		rs.close
	end if	

 Sql= "select * from "&jieducm&"_tribe where username='"&session("useridname")&"'"
 Set rs=server.createobject("ADODB.RECORDSET")
 Rs.Open Sql,conn,3,3
 if not rs.eof then
   tribeid=rs("id")
 end if
				 	
	  Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
	  if not rs.eof then
    	rs("tribeid")=tribeid
    	rs.update
		rs.close
		Response.Write("<script>alert('恭喜您创建部落成功!');window.location.href='index.asp';</script>")
		response.End()
	  end if
	  
elseif action="edit" then
	 tribename2=request.Form("tribename")	 
	 pic=request.Form("pic")
	 tribeinfo=Request.Form("tribeinfo")
	 triberad=request.Form("triberad")
	 Sql= "select * from "&jieducm&"_tribe where username='"&session("useridname")&"'"
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if not rs.eof then
		rs("tribename")=tribename2
    	rs("pic")=pic
		rs("tribeinfo")=tribeinfo
		rs("triberad")=triberad
    	rs.update
		rs.close
	 end if
	 Response.Write("<script>alert('修改成功!');window.location.href='index.asp';</script>")
	 response.End()
end if
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
<SCRIPT language=javascript>
function  save_onclick()
{

 var Price=document.tribe.tribename.value;
  if(Price=="")
  {
    alert("部落名不能为空！");
	document.tribe.tribename.focus();
	return false;
	}
	
 var ProUrl=document.tribe.pic.value;
  if(ProUrl=="")
  {
    alert("部落徽章不能为空！");
	document.tribe.pic.focus();
	return false;
	}
	 var triberad=document.tribe.triberad.value;
  if(triberad=="")
  {
    alert("部落激活码不能为空！");
	document.tribe.triberad.focus();
	return false;
	}
	
 var ProUrl=document.tribe.tribeinfo.value;
  if(ProUrl=="")
  {
    alert("部落简介不能为空！");
	document.tribe.tribeinfo.focus();
	return false;
	}
  return true;  
}
</SCRIPT>
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
    <!--#include file="../user/userleft.asp"--></TD>
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;建立部落 &gt;&gt; </div>
                    <div class=pp8><strong>建立部落</strong></div>
                    <div style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%"><span style="CLEAR: both; MARGIN-TOP: 20px; MARGIN-LEFT: 20px; LINE-HEIGHT: 150%; MARGIN-RIGHT: 20px">* 
建立部落后，你将是酋长，你将有权利管理域内事务。<A href="#" target=_blank><SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 14px; COLOR: red">部落功能介绍--》</SPAN></A><BR>
&nbsp;&nbsp;
* 
官方将根据各个部落的活跃情况，给与酋长与下属子民一定的奖励呵。<BR>
&nbsp;&nbsp;
* 所有会员均可向客服申请建立部落权限。 </span></div>
                    <br><br>
					  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;建立部落</STRONG></td>
                </tr>
              </table>
              <form action="" method="POST" name="tribe" id="tribe" onSubmit="return save_onclick()">
			  <%
			    sql = "select * from "&jieducm&"_tribe where username='"&session("useridname")&"'"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
				 tribename2=rs("tribename")
				 pic=rs("pic")
				 tribeinfo=rs("tribeinfo")
				 startid=rs("id")
				 triberad=rs("triberad")
				end if
			  %>
			  <input name="action" type="hidden" <%if startid<>0 then response.Write("value='edit'") else response.Write("value='ok'") end if%>>
                      <table width="639" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="35" align="right"><div align="right">部落名称：</div></td>
                          <td><INPUT id=tribename name=tribename value="<%=tribename2%>">
                          *名字好听，更能吸引人加入哦！10字以内！</td>
                        </tr>
                        <tr>
                          <td width="134" height="35" align="right"><div align="right">部落徽章：</div></td>
                          <td width="505"><input type="text" name="pic" id="pic" value="<%=pic%>"> <input type="button" name="Submit" value="上传图片"  <%if ( (jifei<5000) and  (vip<>1)) then%>  disabled <%end if%> onClick="window.open('../upload_flash.asp?formname=tribe&editname=pic&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">*建议宽度大小为160*120的jpg|gif|bmp</td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">激活码：</div></td>
                          <td><label>
                            <input type="password" name="triberad" value="<%=triberad%>">
                          </label></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">部落简介：</div></td>
                          <td><TEXTAREA id=tribeinfo name=tribeinfo rows=4 cols=30><%=tribeinfo%></TEXTAREA>
                          *部落介绍，200字以内！</td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                          <td>
			<INPUT id=Button1  type=submit value=建立我的部落，我当酋长了 name=Button1 <%if ( (jifei<5000) and  (vip<>1)) then%>  disabled <%end if%>>			 
			积分达到5000以上或尊贵vip会员才有权限建部落                              </td>
                        </tr>
                      </table>
                    </form>
              </div>
                </div></td>
	          </tr>
	        </table>	   
			 </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close()
set rs=nothing
call closeconn()%>
</BODY></HTML>
