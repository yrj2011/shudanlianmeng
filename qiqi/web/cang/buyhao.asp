<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
On Error Resume Next 
Server.ScriptTimeOut=9999999 
action=request("action")
if action="ok" then

pingjia=Replace(Trim(Request.Form("pingjia")),"'","''")
url=pingjia
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=true
objRegExp.Pattern="<span class=""J_WangWang"" data-nick=""(.*?)""></span>"
set Matches=objRegExp.Execute(wstr)
jieducm_cang=Matches(0).SubMatches(0)
if jieducm_cang="" then
   Response.Write("<script>alert('С���ղؼе�ַ����ȷ!');history.back();</script>")
   response.End()
end if
	   
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_xinyu where shopname='"&jieducm_cang&"' and class=4" ,Conn,3,3  
	  if rs.eof then
	    rs.addnew
		rs("username")=session("useridname")
		rs("shopname")=jieducm_cang
		rs("prourl")=pingjia
		rs("num")=10
		rs("class")=4	
		rs("start")=0
    	rs.update
    	rs.close
		 Response.Write("<script>alert('�󶨳ɹ�!');window.location.href='buyhao.asp';</script>")
		 response.End()
	 else
	 	 Response.Write("<script>alert('�˺��ѱ������û���!');history.back();</script>")
		 response.End()
	 end if
elseif action="del" then
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_xinyu where id="&id&" and username='"&session("useridname")&"'",conn,3,3
rs.close
Response.Write("<script>window.location.href='buyhao.asp';</script>")
response.End()
end if
%>	
<HTML><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->
<SCRIPT language=javascript>
function  save_onclick12()
{
 var ProUrl=document.form1.pingjia.value;
  if(ProUrl=="")
  {
    alert("������С���ղؼе�ַ��");
	document.form1.pingjia.focus();
	return false;
	}

  return true;  
}
</SCRIPT>
<div align="center">
<div 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
    <div style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
      <div style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
        <ul id=task_nav>
          <li><a  href="index.asp">�Ա��ղ���</a> </li>
          <li><a  href="pushmission.asp">���ղ�����</a> </li>
          <li><a href="ReMission.asp">�ѽ�����</a> </li>
          <li><a  href="MyMission.asp">�ѷ�����</a> </li>
		  <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
          <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="Buyhao.asp">�����</a> </li>
          <li><a href="MySave.asp">����ֿ�</a> </li>
        </ul>
      </div>
    </div>
    <div style="CLEAR: both; WIDTH: 910px"><img 
src="../images/task_nav_line.gif"></div>
  </div>
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 15px; BACKGROUND-COLOR: #ffffff">

 <div class="pub_tip4">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="bottom" width="120" align="center">&nbsp;</td>
        <td><div align="left">ע�⣺<br>
          �����Ա��ղ�ʱ����ع�ѡ��<STRONG id="GongKaiFavSpan" title="<img src='images/GongKaiFav.gif' alt='�����ղ�Ҫ���Ŷ' />" jQuery1292576186015="8">�����ղ�</STRONG>��������ƽ̨��ⲻ�����Ƿ��ղ��˴���Ʒ(�����)�� <br>
          <BR>
          �ڳ�����Ʒ����(����Ʒ)����ֱ�ӹ�ѡ��<STRONG>�����ղ�</STRONG>���������ղأ��ղغ�ȥ�ղؼ��ҵ������ղؼ�¼���㡰�༭������ѡ��<STRONG>�����ղ�</STRONG>�����ɡ� </div></td>
      </tr>
    </table>   
  </div>
<DIV style="BACKGROUND: #ffffff; WIDTH: 900px">
<FORM id=form1 name=form1  action="" onSubmit="return save_onclick12()" method=post>
<input name="action" type="hidden" value="ok">
<DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="96%" 
align=center 
style="BORDER-TOP-WIDTH: 0px; MARGIN-TOP: 10px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; MARGIN-BOTTOM: 10px; WIDTH: 50%; BORDER-RIGHT-WIDTH: 0px">
  <TBODY>
  
  <TR>
    <TD height=30><div align="right">С���ղؼе�ַ��</div></TD>
    <TD><div align="left">
      <input name="pingjia" type="text" id="pingjia" size="60" style="WIDTH: 190px">
    </div></TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left"><a href="../news.asp?/1397.html" target="_blank">�ղؼе�ַ������</a></div></TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left">
      <INPUT id=Button1 style="WIDTH: 200px; HEIGHT: 30px"  type=submit value=���ҵ��Ա��ղ��ʺ� name=Button1>
    </div></TD></TR>
  <TR>
    <TD height=50></TD>
    <TD>&nbsp;</TD>
  </TR></TBODY></TABLE>
</DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_3.gif" 
          width=3></TD>
          <TD align=middle width=180 
            background=../images/mytaobaobj1_4.gif><FONT 
            color=#ffffff>�Ѱ󶨵��Ա��ղ�С��</FONT></TD>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_6.gif" 
          width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;</TD>
        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#E9A91E height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="100%" align="center" cellpadding="5" cellspacing="1" bgcolor="#D0DBE3" >
        <TR bgcolor="#E1F2FB">
          <TD height="30" style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">�Ա������ʺ� </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">�ղؼе�ַ</div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">���ʱ��</div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 5%"><div align="center">���� </div></TD>
        </TR>
        <TBODY>
		  <%
			 
			
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=4  order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("������Ϣ")
				Else
					Do While Not Rs.Eof
			  %>
          <TR>
              <TD bgcolor="#FFFFFF" class=centers><img src="../images/j_z.gif" width="13" height="16"><%=rs("shopname")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><label>
      <input name="textfield" type="text" size="50" value="<%=rs("prourl")%>">
    </label></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=rs("now")%></TD>
<TD bgcolor="#FFFFFF" class=centers><a href="#1" onClick="javascript:showDialog('��ȷ��Ҫɾ���������',1,7,'?action=del&id=<%=rs("id")%>')" title="��ȷ��Ҫɾ���������">ɾ��</A></TD>
          </TR>
		    <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
        </TBODY>
      </TABLE>
</DIV></DIV>
<DIV> </DIV>
</FORM>
</DIV>
</DIV></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%rs.close
set rs=nothing
call closeconn()
%>
</BODY></HTML>
