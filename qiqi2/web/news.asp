<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
   id=request("id")
	 if request.QueryString("id")="" then 
		server_v40=Request.ServerVariables("QUERY_STRING") 
		id=Int(replace(replace(server_v40,"/",""),".html",""))
		where="ArticleId="&id
	elseif request.QueryString("Id")<>"" then
		where="ArticleId="&request.QueryString("Id")	
	else
		response.write "ҳ������д����飡"	
		response.end
	end if
	
sql="update "&jieducm&"_Article set hits=hits+1 where "&where&""
conn.execute sql

	   set rs=server.createobject("adodb.recordset")
		rs.open "select  * from "&jieducm&"_Article where  "&where&" ",conn,1,1
		if not rs.eof then
		title=rs("title")
		content=rs("content")
		copyfrom=rs("copyfrom")
		hits=rs("hits")
		classid=rs("classid")
		updatetime=rs("UpdateTime")
		ArticleId=rs("ArticleId")
		else
		  response.write "ҳ������д����飡"	
		  response.end
		end if
		

              Sqlc = "select * from "&jieducm&"_articleclass where ClassID="&classid&" "
				Set Rsc = Server.CreateObject("Adodb.RecordSet")
				Rsc.Open Sqlc,conn,1,1
				IF  not Rsc.Eof Then
				  classname=rsc("classname")
				 rsc.close
				 set rsc=nothing
				 else
				 	 response.write "ҳ������д����飡"	
		             response.end
                end if
%>

<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=title%>-<%=stupname%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.5764" name=GENERATOR></HEAD>
<style type="text/css">
.f12bt{
	font-size: 12px;
	color: #024E68;
	line-height: 20px;
	text-decoration: none;
	font-weight: bold;
	
}
.f12ls1 {
	font-size: 12px;
	color: #0365BE;
	line-height: 20px;
	text-decoration: none;
}
</style>
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
          background=images/Top_12.gif>��վ����</TD>
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
          background=images/Top_12.gif>�Ƽ�����</TD>
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
                  <td class="f12ls1">��ǰλ�ã�<a href="new.asp">��������</a> &gt; <a href="newmore.asp?action=<%=classid%>">
				  <%=classname%></a></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td valign="top" background="images/news2.gif"><table width="669" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="30"><table width="644" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="644" height="60" align="center" valign="bottom" class="f12bt STYLE1"><%=title%></td>
                      </tr>
                      <tr >
                        <td height="33" align="center" class="f12ls1" style="border-bottom:1px #0365be dashed; height:30px;">ʱ�䣺<%=UpdateTime%>    ��Դ��<%=copyfrom%>     �������<%=hits%> </td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30"><table width="644" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="644" height="145" ><br /><%=content%></td>
                      </tr>
                  </table></td>
                </tr>
            </table>
            `</td>
          </tr>
          <tr>
            <td height="26" background="images/news3.gif"><table width="655" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td><span class="hd">
				 ��һƪ��  <%
	  dim rsPrev
	  sql="Select Top 1 ArticleID,Title From "&jieducm&"_Article Where Deleted=0 and ArticleID<" & ArticleID & " order by ArticleID desc"
	  Set rsPrev= Server.CreateObject("ADODB.Recordset")
	  rsPrev.open sql,conn,1,1
	  if rsPrev.Eof then
	  	response.write "û����"
	  else
	    URL="news.asp?/"&rsPrev("ArticleID")&".html"
	  	response.write "<a href='"&url& "'>"&rsPrev("Title") & "</a>"
	  end if
	  rsPrev.close
	  set rsPrev=nothing
	  %>
				   ��һƪ��  <%
	  dim rsNext
	  sql="Select Top 1 ArticleID,Title From "&jieducm&"_Article Where Deleted=0 and ArticleID>" & ArticleID& " order by ArticleID asc"
	  Set rsNext= Server.CreateObject("ADODB.Recordset")
	  rsNext.open sql,conn,1,1
	  if rsNext.Eof then
	  	response.write "û����"
	  else
	   URL="news.asp?/"&rsNext("ArticleID")&".html"
	  	response.write "<a href='"&url& "'>"&rsNext("Title") & "</a>"
	  end if
	  rsNext.close
	  set rsNext=nothing
	  %> </span></td>
                </tr>
            </table>
			
			<table width="655" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td></td>
                </tr>
            </table>
			</td>
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
