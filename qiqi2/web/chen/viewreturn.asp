<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->

<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
action=request("action")
if action="ok" then
id=request("id")
		Set rsr=server.createobject("ADODB.RECORDSET")
	    rsr.open "Select * From "&jieducm&"_guestbook where id="&id&" " ,Conn,3,3  	
		rsr("sort")=1
    	rsr.update
    	rsr.close
		response.Redirect("viewreturn.asp")
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td class="forumRowHighlight" height="22" bgcolor="f1f1f1"> 
      <div align="center"><font color="#000000">�鿴��Ϣ����</font></div>
    </td>
  </tr>
  <tr> 
    <td class="forumRowHighlight" height="169" align="center" valign="top" bgcolor="#FFFFFF"> 
      <%
Function unHtml(content)
	ON ERROR RESUME NEXT
	unHtml=content
	IF content <> "" Then
		unHtml=Server.HTMLEncode(unHtml)
		unHtml=Replace(unHtml,vbcrlf,"<br>")
		unHtml=Replace(unHtml,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		unHtml=Replace(unHtml," ","&nbsp;")
	End IF
	IF Err.Number <>0 Then
		unHtml= "HTMLת���г�������ϵ����Ա<br>"
		Err.Clear
	End IF
End Function
%>
      <%
set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_guestbook order by id desc",conn,3,2
'dim page
'dim e_page
'e_page=5 'ÿҳ��ʾ������
'rs.pagesize=e_page
'if request.querystring("page")="" or request.querystring("page")=0 then
'page=1
'else
'page=request.querystring("page")
'rs.absolutepage=trim(request.querystring("page"))
'end if
%>
      <SCRIPT language=JavaScript>

function is_number(str)
{
	exp=/[^0-9()-]/g;
	if(str.search(exp) != -1)
	{
		return false;
	}
	return true;
}

function CheckInput(){

	if(form.name.value==''){
		alert("��û����д�ǳƣ�");
		form.name.focus();
		return false;
	}
	if(form.name.value.length>20){
		alert("�ǳƲ��ܳ���20���ַ���");
		form.name.focus();
		return false;
	}

	if(!is_number(document.form.qq.value)){
		alert("QQ�ű��������֣�");
		form.qq.focus();
		return false;
	}

	if(form.content.value==''){
		alert("��û����д�������ݣ�");
		form.content.focus();
		return false;
	}
	if(form.content.value.length>255){
		alert("�������ݲ��ܳ���255���ַ���");
		form.content.focus();
		return false;
	}
	
	return true;
}
</SCRIPT>
      <br> 
      <%
if rs.eof and rs.bof then
%>
      <table width="380" border="0" align="center" cellpadding="4" >
        <tr> 
          <td class="forumRowHighlight" height="40" align="center"> <p>��û���κ����ԣ�</p></td>
        </tr>
      </table>
      <%else%>
      <%
	  	rs.PageSize =5 'ÿҳ��¼����
		iCount=rs.RecordCount '��¼����
		iPageSize=rs.PageSize
    	maxpage=rs.PageCount 
    	page=request("page")
    	per_page=rs.PageSize
    
    if Not IsNumeric(page) or page="" then
        page=1
    else
        page=cint(page)
    end if
    
    if page<1 then
        page=1
    elseif  page>maxpage then
        page=maxpage
    end if
    
    rs.AbsolutePage=Page
	if page=maxpage then
		x=iCount-(maxpage-1)*iPageSize
	else
		x=iPageSize
	end if
%>
      <table width="100%" border="0">
        <tr> 
          <td class="forumRowHighlight" colspan="12" height="25" align="center" bgcolor="#FFFFFF" > 
            <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>
          </td>
        </tr>
      </table>
      <%
For i=1 To x
'do while not rs.eof and e_page>0
%>
      <TABLE width=98% border=0 align="center" cellPadding=0 cellSpacing=1 bgcolor="#CCCCCC">
        <TBODY>
          <TR> 
            <td class="forumRowHighlight" bgcolor="f1f1f1"><TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                <TBODY>
                  <TR bgcolor="f1f1f1"> 
                    <td class="forumRowHighlight" width="67" height="23" align="center" bgcolor="f1f1f1"><font color="#FF0000"><strong>��<%=rs("id")%>��</strong></font></TD>
                    <td class="forumRowHighlight" width="464" align="right"> 
                      <%if rs("qq")<>"" then%>
                      <a title=QQ��<%=rs("qq")%>><img src="../images/qq.gif" width="16" height="16" border="0"> 
                      <%=rs("qq")%></a> 
                      <%end if%>
                      &nbsp; 
                      <%if rs("email")<>"" then%>
                      <a href="mailto:<%=rs("email")%>" target="_blank" title=Email��<%=rs("email")%>><img src="../images/email.gif" width="14" height="15" border="0"></a> 
                      <a href="mailto:<%=rs("email")%>" target="_blank" title=Email��<%=rs("email")%>><%=rs("email")%></a> 
                      <%end if%>
                      <%if rs("url")<>"" then%>
                      <a href="<%=rs("url")%>" target="_blank" title=��ҳ��<%=rs("url")%>><img src="../images/url.gif" width="16" height="16" border="0"></a> 
                      <a href="<%=rs("url")%>" target="_blank" title=��ҳ��<%=rs("url")%>><%=rs("url")%></a> 
                      <%end if%>
                      &nbsp;</TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
          </TR>
          <TR> 
            <td class="forumRowHighlight" ><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY>
                  <TR> 
                    <td class="forumRowHighlight" bgcolor="#FFFFFF"> <TABLE width="100%" border=0 style="table-layout:fixed;word-break:break-all">
                        <TBODY>
                          <TR> 
                            <td class="forumRowHighlight" width="109"><img src="../images/<%=rs("sex")%>.gif"></TD>
                            <td class="forumRowHighlight" colspan="2"> <strong><font color="<%=flzhbt%>">������</font></strong><font color="<%=flzhzt%>"><%=rs("name")%></font><br> 
                              <strong><font color="<%=flzhbt%>">���ݣ�</font></strong><font color="<%=flzhzt%>"><%=unHtml(rs("content"))%></font></TD>
                          </TR>
                          <TR> 
                            <td class="forumRowHighlight" colspan="2">&nbsp; [<a href="fkdel.asp?id=<%=rs("id")%>">ɾ��</a>] 
                            [<a href="fkreplay.asp?id=<%=rs("id")%>">�ظ�</a>][<a href="?action=ok&id=<%=rs("id")%>">�ö�</a>]</TD>
                            <td class="forumRowHighlight" width="250" align="right"><%=rs("time")%></TD>
                          </TR>
                        </TBODY>
                      </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
          </TR>
          <%if rs("reply")<>"" then%>
          <TR> 
            <td class="forumRowHighlight" bgColor=#f2f2f2> <table width="100%" border="0" style="table-layout:fixed;word-break:break-all">
                <tr> 
                  <td class="forumRowHighlight" width="10">&nbsp;</td>
                  <td class="forumRowHighlight" width="567"><font color="#FF0000">����Ա�ظ���</font><br> <%=unHtml(rs("reply"))%></td>
                </tr>
              </table></TD>
          </TR>
          <%end if%>
        </TBODY>
      </TABLE>
      <br> 
      <%
'e_page=e_page-1
'rs.movenext
'loop
		RS.MoveNext
next
%>
      <table width="100%" border="0">
        <tr> 
          <td class="forumRowHighlight" > 
            <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>
          </td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
end if
%>
    </td>
  </tr>
</table>
</body>
</html><%
Sub PageControl(iCount,pagecount,page,table_style,font_style,per_page)
'������һҳ��һҳ����
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
	
    temp=""

    Response.Write("<table " & Table_style & ">" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<td  align=right>" & vbCrLf )
    Response.Write(font_style & vbCrLf )    
        
    if page<=1 then
        Response.Write ("��ҳ " & vbCrLf)        
        Response.Write ("��ҳ " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">��ҳ</A> " & vbCrLf)
    end if

    if page>=pagecount then
        Response.Write ("��ҳ " & vbCrLf)
        Response.Write ("βҳ " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">βҳ</A> " & vbCrLf)            
    end if

    Response.Write(" ҳ�Σ�" & page & "/" & pageCount & "ҳ" &  vbCrLf)
    Response.Write(" ����" & iCount & "��/ÿҳ"&per_page&"��" &  vbCrLf)
    Response.Write(" ת��" & "<INPUT TYEP=TEXT NAME=page SIZE=4 Maxlength=8 VALUE=" & page & ">" & "ҳ"  & vbCrLf & "<INPUT type=submit style=""font-size: 9pt"" value=GO class=b2>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>