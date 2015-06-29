<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
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
action=request.Form("action")
if action="ok" then
call login()
end if
 %>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE><%=stupname%>-<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<LINK href="css/login.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
</HEAD>
<BODY  onload=document.form.name.focus()> 
<!--#include file=jieducm_top.asp-->
<TABLE cellSpacing=0 cellPadding=0 width=930 align=center border=0>
  <TBODY>
  <TR>
    <TD width=20 height=32><IMG src="images/Top_10.gif"></TD>
    <TD align=right width=21 background=images/Top_11.gif><IMG height=32 
      src="images/Top_9.gif" width=10></TD>
    <TD class=K_mttitle align=middle width=120 
      background=images/Top_12.gif>操作教程</TD>
    <TD align=left width=47 background=images/Top_11.gif><IMG height=32 
      src="images/Top_13.gif" width=10></TD>
    <TD align=left width=728 background=images/Top_11.gif></TD>
	
    <TD align=right width=24 background=images/Top_11.gif><IMG height=32 
      src="images/Top_14.gif" width=6></TD></TR>
	  
  <TR>
    <TD class=K_mtcontent align=left colSpan=6>
	 <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <tr>
    <td><BR/>
	<div align="center">在开始您在人人来网任务之前请先了解一下<strong><a href="/news.asp?/1454.html" target="_blank"><span style="color:#3333FF;">《人人来刷钻平台规章制度》，点此在新页面打开</span></a>，</strong>看完后如有任何疑问或建议都可以<a href="/help/viewreturn.asp" target="_blank">在线留言</a>！</div>
	<BR/></td>
  </tr>
</table>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
		        
         <TR> 
          <TD align=middle width="48%" height=350>
			<table width="90%" border="0" align="right" cellpadding="8" cellspacing="0" >
              <tr>
                <td colspan="2"><div align="center"><img src="images/teach/vide.gif" width="80" height="66"></div></td>
                </tr>
              <tr>
                <td colspan="2"><div align="center"><a href="/newmore.asp?action=97" target="_blank"> 人人来网视频教程</a></div></td>
              </tr>
			   <tr>
           <td colspan="2"><div align="center"><a href="/news.asp?/1462.html" target="_blank"><img src="images/teach/curse.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  <tr>
                <td colspan="2">&nbsp;</td>
              </tr>
               <tr>
                <td colspan="2"><div align="center"><a href="/news.asp?/1461.html" target="_blank"><img src="images/teach/curse_1.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
              <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			   <tr>
                <td colspan="2"><div align="center"><a href="/news.asp?/1460.html" target="_blank"><img src="images/teach/curse_2.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			   <tr>
                <td colspan="2"><div align="center"><a href="/news.asp?/1459.html" target="_blank"><img src="images/teach/curse_3.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><a href="/news.asp?/1458.html" target="_blank"><img src="images/teach/curse_4.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			    <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><a href="/news.asp?/1457.html" target="_blank"><img src="images/teach/curse_5.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  
            </table>
			</TD>	
	
          <TD width="1%">
            <DIV 
            style="WIDTH: 2px; HEIGHT: 230px; BACKGROUND-COLOR: #3da3e7"></DIV></TD>
          <TD align=middle width="48%" height=350>
			<table width="90%" border="0" align="right" cellpadding="8" cellspacing="0" >
              <tr>
                <td colspan="2"><div align="center"><img src="images/teach/see.gif" width="80" height="66"></div></td>
                </tr>
              <tr>
                <td colspan="2"><div align="center"> 人人来网图文教程</div></td>
              </tr>
			   <tr>
           <td colspan="2"><div align="center"><a href="/news.asp?/1464.html" target="_blank"><img src="images/teach/curse.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center">&nbsp;</td>
              </tr> 
               <tr>
                <td colspan="2"><div align="center"><a href="/zcjc01.asp" target="_blank"><img src="images/teach/curse_1.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
              <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			   <tr>
                <td colspan="2"><div align="center"><a href="/jrjc01.asp" target="_blank"><img src="images/teach/curse_2.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			   <tr>
                 <td colspan="2"><div align="center"><a href="/jc01.asp" target="_blank"><img src="images/teach/curse_3.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><a href="/czjc01.asp" target="_blank"><img src="images/teach/curse_4.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><img src="images/teach/nest.jpg" width="22" height="30"></div></td>
              </tr>
			  <tr>
                <td colspan="2"><div align="center"><a href="/user/ReMoney.asp" target="_blank"><img src="images/teach/curse_5.gif" width="153" height="40" border="0" ></a></div></td>
              </tr>
            </table>
			</TD>
			
		  </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>

        <TR>
                <td colspan="2"><div align="center"> <img src="images/teachfooter.jpg" width="980" height="150"></div></td>
		  </TR>
		  
		  <TR>
                <td colspan="2"><div align="center"><a href="/fpsc01.asp" target="_blank"> <img src="/admin/UploadFiles/fpsc/fpnews.gif" width="980" height="80" border="0" ></a></div></td>
		  </TR>
		   <TR>
                <td colspan="2"><div align="center"><img src="/admin/UploadFiles/br.gif"></div></td>
		  </TR>
		
<!--#include file=jieducm_bottom.asp-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
