<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%
action=HtmlEncode(trim(request.form("action")))
if action="xie" then
if HtmlEncode(trim(request.form("code")))="" then
	Response.Write("<script>alert('��������֤��!');history.back();</script>")	
	response.End()
elseif codenum<>HtmlEncode(trim(request.form("code"))) then
    	num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�����ֻ���֤"
		rs("content")=session("useridname")&"�������֤������ֻ����յĲ�һ��,ʱ�䣺"&now()&"IP:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs("price")=""
		rs("jiegou")="��֤ʧ��"
    	rs.update
    	rs.close 
     Response.Write("<script>alert('�������֤������ֻ����յĲ�һ��!');history.back();</script>")	
	 response.End()
end if
 
       Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select dst From "&jieducm&"_register where username='"&session("useridname")&"' and dst=0" ,Conn,3,3     
	   if  (not rs.bof) and (not rs.bof) then
		rs("dst")=request("dst")		
    	rs.update
    	rs.close
	   else
	        Response.Write("<script>alert('�벻Ҫ�ظ���!');history.back();</script>")	
	        response.End()
	   end if
	   
if stupphonestart=1  then	
sqlr="update "&jieducm&"_register set  jifei=jifei+"&stuptjrzhu&" where username='"&tjr&"'"
conn.execute sqlr
else 

end if
	   
		num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�����ֻ���֤"
		rs("content")=session("useridname")&"���ֻ��ɹ�"&now()&""
		rs("price")=""
		rs("jiegou")="��֤�ɹ�"
    	rs.update
    	rs.close 
	 Response.Write("<script>alert('���ֻ��ɹ�!');window.location.href='Center_Userlock.asp';</script>")		
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>��Ա����-<%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
</head>
<SCRIPT LANGUAGE="JavaScript">
<!--
function save_onclick()
{
  var dst=document.form1.dst.value
  if(dst.length==0)
  {
     alert("�ֻ����������д��");
	 document.form1.dst.focus();
	 return false;
	 }
 if (dst.length!=11)
 {
      alert("����д��ȷ���ֻ����룡");
	 document.form1.dst.focus();
	 return false;
 }
 
 var code=document.form1.code.value
  if(code.length==0)
  {
     alert("�����������д��");
	 document.form1.code.focus();
	 return false;
	 }
  
  return true;
}

alert("aa");
//-->
</SCRIPT>
<body>
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
      <td height="5"> </td>
    </tr>
</table>
				<div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;���ֻ�&gt;&gt; </div>
                    <div class=pp8><strong>���ֻ�</strong></div>
                    
					  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
          
          <TR>
            <TD height=600 valign="top" class="border_hui3 wenzi" ><TABLE width="99%"
             border=0 align="center" cellPadding=0  cellSpacing=0  class="padding_left">
              <TBODY>
                <TR>
                  <TD vAlign=top><%if  dst=0 and dst1=0 then%>
					<FORM id=form1 name=form1 onSubmit="return save_onclick()"  method=post action="sendsms.asp"> 
					<input name="action" type="hidden" value="ok" />
                    <TABLE  style="MARGIN-TOP: 10px; WIDTH: 100%; BACKGROUND-COLOR: #f3f8fe">
                      <DIV id=Panel1></DIV>
                      <TBODY>
                        <TR>
                          <TD width="21%" 
                            style="WIDTH: 25%; TEXT-ALIGN: right">�����ֻ����룺</TD>
                          <TD width="79%" style="WIDTH: 75%; TEXT-ALIGN: left"><INPUT 
                              id=dst maxLength=20 name=dst>
                            �����κ�������</TD>
                        </TR>
                        <TR>
                          <TD style="WIDTH: 25%; TEXT-ALIGN: right"></TD>
                          <TD 
                              style="WIDTH: 75%; TEXT-ALIGN: left">*�ƶ�����ͨ��С��ͨ������+���룩</TD>
                        </TR>
                        <TR>
                          <TD style="WIDTH: 25%; TEXT-ALIGN: right">�����룺</TD>
                          <TD style="WIDTH: 75%; TEXT-ALIGN: left"><INPUT id=code type=password maxLength=20 name=code>
                              </SPAN></TD>
                        </TR>
                      <DIV></DIV>
                      <TR>
                        <TD style="WIDTH: 25%; TEXT-ALIGN: right"></TD>
                        <TD style="WIDTH: 75%; TEXT-ALIGN: left">
						<INPUT id=button1 type=submit value=��ʼ���ͻ�ȡ������ name=button1 ></TD>
                      </TR>
                      <TR>
                        <TD style="TEXT-ALIGN: center" colSpan=2>
	              <input name="msg" type="hidden"  value="�װ����û�������֤���ǣ�" />   </TD>
                      </TR>
                      <TR>
                        <TD colspan="2" ><table width="66%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="6%"><img src="../img/jieducm_union_01.gif" width="25" height="22"></td>
                            <td width="29%">��һ���������ֻ����룻</td>
                            <td width="6%"><img src="../img/jieducm_union_02.gif" width="25" height="22"></td>
                            <td width="59%">�ڶ����������ֻ����յ�����֤�� </td>
                          </tr>
                        </table></TD>
                        </TR>
                      </TABLE>
					  </form>
					  <%elseif dst=0 and dst1<>0 then%>
<FORM id=form1 name=form1   method=post action=""> 
<input name="action" type="hidden" value="xie" />
                    <TABLE  style="MARGIN-TOP: 10px; WIDTH: 100%; BACKGROUND-COLOR: #f3f8fe">
                      <DIV id=Panel1></DIV>
                      <TBODY>
               <TR>
    <TD style="WIDTH: 25%; TEXT-ALIGN: right">�����յ�����֤�룺</TD>
    <TD style="WIDTH: 75%; TEXT-ALIGN: left"><INPUT id="code" type=password  maxLength=20 name="code">
    </SPAN>���Ż���30���ڷ��������ֻ�����û���յ�����ϵ����Ա��</TD>
  </TR>
  <DIV></DIV>
  
  <TR>
    <TD></TD>
    <TD style="WIDTH: 75%; TEXT-ALIGN: left">
      <INPUT id=button1 type=submit value="���������ֻ��յ�����֤��" name=button1>
      <input name="dst" type="hidden" value="<%=dst1%>"  />   </TD></TR>
                      <DIV></DIV>
                      <TR>
                        <TD style="WIDTH: 25%; TEXT-ALIGN: right"></TD>
                        <TD style="WIDTH: 75%; TEXT-ALIGN: left">�벻Ҫ�رմ˴��ڣ��������յ�����֤�룡</TD>
                      </TR>
                      <TR>
                        <TD style="TEXT-ALIGN: center" colSpan=2>&nbsp;</TD>
                      </TR>
                      <TR>
                        <TD colspan="2" ><table width="66%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="6%"><img src="../img/jieducm_union_01.gif" width="25" height="22" /></td>
                            <td width="29%">��һ���������ֻ����룻</td>
                            <td width="6%"><img src="../img/jieducm_union_02.gif" width="25" height="22"></td>
                            <td width="59%">�ڶ����������ֻ����յ�����֤�� </td>
                          </tr>
                        </table></TD>
                        </TR>
                      </TABLE>
					  </form> 
					  <%else%>
					   <TABLE  style="MARGIN-TOP: 10px; WIDTH: 100%; BACKGROUND-COLOR: #f3f8fe">
					  <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">
			  �Ѱ󶨵��ֻ����ǣ�*******<%=right(dst,4)%>			</td>
              </tr></TABLE>
					  <%end if%>              
                      <!--2�� --></TD>
                </TR>
            </TABLE> </TD>
          </TR>
      </TABLE>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>
  <!--#include file=../inc/jieducm_bottom.asp-->
  <%call closeconn()%>
</body>
</html>