<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#INCLUDE FILE="background.asp"-->
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
Sqlc = "select classid,classname from "&jieducm&"_articleclass where ClassID="&action&" "
Set Rsc = Server.CreateObject("Adodb.RecordSet")
Rsc.Open Sqlc,conn,1,1
IF  not Rsc.Eof Then
classid=rsc("classid")
classname=rsc("classname")
else
		  response.write "页面参数有错，请检查！"	
		  response.end
end if
rsc.close
set rsc=nothing
%>

<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=classname%>-<%=stupname%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.5764" name=GENERATOR>
</HEAD>
<BODY>
<!--#include file=jieducm_top.asp-->
<TABLE cellSpacing=0 cellPadding=0 width=960 align=center border=0>
  

  <tr>
    <td valign="top"><table width="960" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:10px 0 10px 0;">
      <tr>
        <td colspan="2" valign="top">
		</td>
        </tr>
      <tr>
        <td width="271" valign="top"><table width="260" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><TABLE cellSpacing=0 cellPadding=0 width=260 border=0>
        <TBODY>
        <TR>
          <TD width=20><IMG height=32 
            src="images/Top_10.gif" 
            width=20></TD>
          <TD width=10><IMG height=32 
            src="images/Top_9.gif" 
            width=10></TD>
          <TD class=K_mttitle width=95 
          background=images/Top_12.gif>网站公告</TD>
          <TD width=10><IMG height=32 
            src="images/Top_13.gif" 
            width=10></TD>
          <TD width=119 
          background=images/Top_11.gif></TD>
          <TD width=6><IMG height=32 
            src="images/Top_14.gif" 
            width=6></TD></TR>
        <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(2,10,15)%></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
          </tr>
		  <tr height="10"><td></td></tr>
		  
          <tr>
            <td><TABLE cellSpacing=0 cellPadding=0 width=260 border=0>
        <TBODY>
        <TR>
          <TD width=20><IMG height=32 
            src="images/Top_10.gif" 
            width=20></TD>
          <TD width=10><IMG height=32 
            src="images/Top_9.gif" 
            width=10></TD>
          <TD class=K_mttitle width=95 
          background=images/Top_12.gif>推荐文章</TD>
          <TD width=10><IMG height=32 
            src="images/Top_13.gif" 
            width=10></TD>
          <TD width=119 
          background=images/Top_11.gif></TD>
          <TD width=6><IMG height=32 
            src="images/Top_14.gif" 
            width=6></TD></TR>
        <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call elitenew(10,15)%></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
          </tr>
        </table></td>
        <td width="699" valign="top"><table width="690" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="26" background="images/news1.gif"><table width="655" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td class="f12ls1">当前位置：<a href="new.asp">新闻中心</a> &gt; <%=classname%></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td valign="top" background="images/news2.gif"><table  width="669" border="0" align="center" cellpadding="0" cellspacing="0">
                <%				
	     sql="select * from "&jieducm&"_article where classid in ("&action&") order by Articleid desc"
		Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then			
	dim maxperpage,url,PageNo
	 url="newmore.asp?action="&action&""
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
	j=1
if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
DO WHILE NOT rs.EOF AND RowCount>0
URL1="news.asp?/"&rs("Articleid")&".html"
	   %>
			    <tr>
                  <td height="30" background="images/news4.gif"><table width="659" border="0" align="right" cellpadding="0" cellspacing="0" class=LeftNews >
                      <tr>
                        <td width="644" class="f12hs"> <%="<a href='"+URL1+"' target=""_blank"" title='"+rs("title")+"' >"%>
						 <font color="<%=rs("TitleFontColor")%>">
						 <%
						if rs("TitleFontType")="0" then
						 response.write rs("title")
						elseif rs("TitleFontType")="1" then 
						response.write "<strong>"&rs("title")&"</strong>"
						elseif rs("TitleFontType")="2" then 
						response.write "<em>"&rs("title")&"</em>"
						elseif rs("TitleFontType")="3" then 
						response.write "<strong><em>"&rs("title")&"</strong></em>"
						end if
						%></font>
</A></td>
                      </tr>
                  </table></td>
                </tr>
 	<%
	 j=j+1
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
            </table></td>
          </tr>
          <tr>
            <td height="26" background="images/news3.gif"><table width="655" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td><span class="hd"><% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></span></td>
                </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      
    </table></td>
  </tr>
  
 
</table>
<!--#include file=jieducm_bottom.asp-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
