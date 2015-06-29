<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
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
%>
<html>
<head>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<LINK 
href="help.css" type=text/css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<!--#include file="top.asp"-->
<table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="10" valign="top">&nbsp;</td>
    <td width="761" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td style="BORDER-bottom:#ccc 1px solid;BORDER-top:#ccc 1px solid;" height="28"  background="images/left_1.jpg"><div align="center">尊敬的客户，留下您宝贵的建议，我们愿为您提供更好的服务！</div></td>
      </tr>
      <tr>
        <td><span class="b">
          
        </span>
          <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
              <tr> 
                <td bordercolor="#FFFFFF" bgcolor="#FFFFFF"><br>
                  <table width="95%"  border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td valign="top"> <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td background="img/bg1_4color_.jpg"><img src=img/bg1_4color_.jpg" width="1" height="3"></td>
                          </tr>
                        </table>
                        <%if ly="0" then%>
                        <FORM action=fksave.asp method=post name="form" id="form">
                          <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
                            <tr> 
                              <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                            </tr>
                            <tr> 
                              <td height="25" bgcolor="efefef"> <strong>&nbsp;&nbsp; 
                                </strong><strong>用户留言</strong></td>
                            </tr>
                            <tr> 
                              <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                            </tr>
                          </table>
                          <br>
                          <TABLE width="90%" border=0 align="center" cellPadding=0 cellSpacing=1>
                            <TBODY>
                              <TR> 
                                <TD width="407" height="30">&nbsp;&nbsp; 姓名： 
                                  <input name=name class="wenbenkuang" id="name"></TD>
                              </TR>
                              
                              
                              <TR> 
                                <TD height="30">&nbsp;&nbsp; 留言内容：</TD>
                              </TR>
                              <TR> 
                                <TD height="130" noWrap> &nbsp;&nbsp; <textarea name=content cols=60 rows=8 id="content" class="wenbenkuang"></textarea>                                </TD>
                              <TR> 
                                <TD height="30" colspan="2"> &nbsp;&nbsp; <input name="submit2" type=submit class="go-wenbenkuang" id="submit4" value=我要留言> 
                                  &nbsp;&nbsp; <input name="submit32" type=reset class="go-wenbenkuang" id="submit32" value=清除重来> 
                                  <%
Randomize '初始代随机数种子
num1=rnd() '产生随机数num1
num1=int(26*num1)+65 '修改num1的范围以使其是A-Z范围的Ascii码，以防表单名出错
session("antry")="test"&chr(num1) '产生随机字符串
%>
                                  <input name="temp" type="hidden" id="temp" value="<%=session("antry")%>">                                </TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                        </FORM>
                        <p>&nbsp; </p>
                        <p> 
                          <%else%>
                        </p>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
                                <tr> 
                                  <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                </tr>
                                <tr> 
                                  <td height="25" bgcolor="efefef"> <strong>&nbsp;&nbsp; 
                                    </strong><b>留言信息</b></td>
                                </tr>
                                <tr> 
                                  <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                </tr>
                              </table>
                              <table width="98%" border="0" align="center">
                                <tr> 
                                  <td align="center"> 
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
		unHtml= "HTML转换中出错请联系管理员<br>"
		Err.Clear
	End IF
End Function
%>
                                    <%
set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_guestbook order by sort",conn,3,2
'dim page
'dim e_page
'e_page=5 '每页显示留言数
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
		alert("您没有填写昵称！");
		form.name.focus();
		return false;
	}
	if(form.name.value.length>20){
		alert("昵称不能超过20个字符！");
		form.name.focus();
		return false;
	}

	if(!is_number(document.form.qq.value)){
		alert("QQ号必须是数字！");
		form.qq.focus();
		return false;
	}

	if(form.content.value==''){
		alert("您没有填写留言内容！");
		form.content.focus();
		return false;
	}
	if(form.content.value.length>255){
		alert("留言内容不能超过255个字符！");
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
                                        <td height="40" align="center"> <p>还没有任何留言！</p></td>
                                      </tr>
                                    </table>
                                    <%else%>
                                    <%
	  	rs.PageSize =5 '每页记录条数
		iCount=rs.RecordCount '记录总数
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
                                        <td colspan="12" height="25" align="center" bgcolor="#FFFFFF" > 
                                          <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>                                        </td>
                                      </tr>
                                    </table>
                                    <%
For i=1 To x
'do while not rs.eof and e_page>0
%>
                                    <TABLE width=100% border=0 align="center" cellPadding=0 cellSpacing=1 bgcolor="#CCCCCC">
                                      <TBODY>
                                        <TR> 
                                          <TD bgcolor="f1f1f1"><TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                                              <TBODY>
                                                <TR bgcolor="f1f1f1"> 
                                                  <TD width="67" height="23" align="center" bgcolor="f1f1f1"><font color="#FF0000"><strong>第<%=rs("id")%>条</strong></font></TD>
                                                  <TD width="464" align="right"> 
                                                    <%if rs("qq")<>"" then%>
                                                    <img src="images/qq.gif" width="16" height="16" border="0"> 
                                                    <%=rs("qq")%> 
                                                    <%end if%>
                                                    &nbsp; 
                                                    <%if rs("email")<>"" then%>
                                                    <a href="mailto:<%=rs("email")%>" target="_blank" title=Email：<%=rs("email")%>><img src="images/email.gif" width="14" height="15" border="0"></a> 
                                                    <a href="mailto:<%=rs("email")%>" target="_blank" title=Email：<%=rs("email")%>><%=rs("email")%></a> 
                                                    <%end if%>
                                                    <%if rs("url")<>"" then%>
                                                    <a href="<%=rs("url")%>" target="_blank" title=主页：<%=rs("url")%>><img src="images/url.gif" width="16" height="16" border="0"></a> 
                                                    <a href="<%=rs("url")%>" target="_blank" title=主页：<%=rs("url")%>><%=rs("url")%></a> 
                                                    <%end if%>
                                                    &nbsp;</TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE></TD>
                                        </TR>
                                        <TR> 
                                          <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                                              <TBODY>
                                                <TR> 
                                                  <TD bgcolor="#FFFFFF"> <TABLE width="100%" border=0 style="table-layout:fixed;word-break:break-all">
                                                      <TBODY>
                                                        <TR> 
                                                          <TD width="109"><img src="images/<%=rs("sex")%>.gif"></TD>
                                                          <TD colspan="2"> <strong><font color="<%=flzhbt%>">姓名：</font></strong><font color="<%=flzhzt%>"><%=rs("name")%></font><br> 
                                                            <strong><font color="<%=flzhbt%>">内容：</font></strong><font color="<%=flzhzt%>"><%=unHtml(rs("content"))%></font></TD>
                                                        </TR>
                                                        <TR> 
                                                          <TD colspan="2">&nbsp;</TD>
                                                          <TD width="250" align="right"><%=rs("time")%></TD>
                                                        </TR>
                                                      </TBODY>
                                                    </TABLE></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE></TD>
                                        </TR>
                                        <%if rs("reply")<>"" then%>
                                        <TR> 
                                          <TD bgColor=#f2f2f2> <table width="100%" border="0" style="table-layout:fixed;word-break:break-all">
                                              <tr> 
                                                <td width="10">&nbsp;</td>
                                                <td width="567"><font color="#FF0000">管理员回复：</font><br> 
                                                  <%=unHtml(rs("reply"))%></td>
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
                                        <td> 
                                          <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>                                        </td>
                                      </tr>
                                    </table>
                                    <%
rs.close
set rs=nothing
end if
%>                                  </td>
                                </tr>
                                <tr> 
                                  <td align="center"><FORM action=fksave.asp method=post name="form" id="form">
                                      <br>
                                      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
                                        <tr> 
                                          <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                        </tr>
                                        <tr> 
                                          <td height="25" bgcolor="efefef"> <strong>&nbsp;&nbsp;<img src="img/dot_03.gif" width="9" height="9" border="0"> 
                                            </strong><strong>用户留言</strong></td>
                                        </tr>
                                        <tr> 
                                          <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                        </tr>
                                      </table>
                                      <TABLE width="100%" border=0 cellPadding=0 cellSpacing=1>
                                        <TBODY>
                                        <TD height="30">&nbsp;&nbsp; 姓名： 
                                          <input name=name class="wenbenkuang" id="name" value="<%=session("useridname")%>" readonly></TD>
                                        </TR>
                                        
                                        
                                        <TR> 
                                          <TD height="30">&nbsp;&nbsp; 留言内容：</TD>
                                        </TR>
                                        <TR> 
                                          <TD height="130" noWrap> &nbsp;&nbsp; 
                                            <textarea name=content cols=60 rows=8 id="content" class="wenbenkuang"></textarea>                                          </TD>
                                        <TR> 
                                          <TD height="30" colspan="2"> &nbsp;&nbsp; 
                                            <input name="submit" type=submit class="go-wenbenkuang" id="submit" value=我要留言> 
                                            &nbsp;&nbsp; <input name="submit3" type=reset class="go-wenbenkuang" id="submit3" value=清除重来> 
                                            <%
Randomize '初始代随机数种子
num1=rnd() '产生随机数num1
num1=int(26*num1)+65 '修改num1的范围以使其是A-Z范围的Ascii码，以防表单名出错
session("antry")="test"&chr(num1) '产生随机字符串
%>
                                            <input name="temp" type="hidden" id="temp" value="<%=session("antry")%>">                                          </TD>
                                        </TR>
                                      </TABLE>
                                    </FORM></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table>
                        <%end if%>                      </td>
                    </tr>
                  </table>
                  <br>                </td>
              </tr>
            </table></td>
      </tr>
    </table></td>
    <td width="10"></td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
<%
Sub PageControl(iCount,pagecount,page,table_style,font_style,per_page)
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
	
    temp=""

    Response.Write("<table " & Table_style & ">" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<td  align=right>" & vbCrLf )
    Response.Write(font_style & vbCrLf )    
        
    if page<=1 then
        Response.Write ("首页 " & vbCrLf)        
        Response.Write ("上页 " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>首页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">上页</A> " & vbCrLf)
    end if

    if page>=pagecount then
        Response.Write ("下页 " & vbCrLf)
        Response.Write ("尾页 " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">下页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">尾页</A> " & vbCrLf)            
    end if

    Response.Write(" 页次：" & page & "/" & pageCount & "页" &  vbCrLf)
    Response.Write(" 共有" & iCount & "条/每页"&per_page&"条" &  vbCrLf)
    Response.Write(" 转到" & "<INPUT TYEP=TEXT NAME=page SIZE=4 Maxlength=8 VALUE=" & page & ">" & "页"  & vbCrLf & "<INPUT type=submit style=""font-size: 9pt"" value=GO class=b2>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>

<SCRIPT LANGUAGE="JavaScript">
<!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function checkfk()
{
   if(checkspace(document.fkinfo.fksubject.value)) {
	document.fkinfo.fksubject.focus();
    alert("您没有填写主题！");
	return false;
  }
   if(checkspace(document.fkinfo.fkusername.value)) {
	document.fkinfo.fkusername.focus();
    alert("请填写您的姓名！");
	return false;
  }
   if(checkspace(document.fkinfo.fklaizi.value)) {
	document.fkinfo.fklaizi.focus();
    alert("请填写您来自哪里！");
	return false;
  }
     if(checkspace(document.fkinfo.fkcontent.value)) {
	document.fkinfo.fkcontent.focus();
    alert("请填写反馈信息内容！");
	return false;
  }
  if(document.fkinfo.fkemail.value.length!=0)
  {
    if (document.fkinfo.fkemail.value.charAt(0)=="." ||        
         document.fkinfo.fkemail.value.charAt(0)=="@"||       
         document.fkinfo.fkemail.value.indexOf('@', 0) == -1 || 
         document.fkinfo.fkemail.value.indexOf('.', 0) == -1 || 
         document.fkinfo.fkemail.value.lastIndexOf("@")==document.fkinfo.fkemail.value.length-1 || 
         document.fkinfo.fkemail.value.lastIndexOf(".")==document.fkinfo.fkemail.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.fkinfo.fkemail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
   document.fkinfo.fkemail.focus();
   return false;
   }
}
//-->
</script> 
