<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************����������ˢϵͳ***********************
'������רΪ������ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
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
if czm<>request("czm") then
   Response.Write("<script>alert('�������������������룡');history.back();</script>")
   response.End()
end if
hao=Replace(Trim(Request.Form("hao")),"'","''")
	Sql = "select count(*) as xinyu from "&jieducm&"_xinyu where username='"&session("useridname")&"' and class=2 and start=0 "
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF not Rs.Eof Then
	 ii=rs("xinyu")
	End IF
				
       if ii>=stupbuyhao then
	   	 Response.Write("<script>alert('һ����ֻ�ܰ�"&stupbuyhao&"�����!');history.back();</script>")
         response.End()
	   end if
	   
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_xinyu where shopname='"&request("hao")&"'" ,Conn,3,3  
	  if rs.eof then
	    rs.addnew
		rs("username")=session("useridname")
		rs("shopname")=hao
		rs("prourl")=pingjia
		rs("num")=10
		rs("class")=2
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
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->

<!-- �����ز˵����� --><!-- �����ز˵�BIG -->
<div align="center">
<DIV 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 700px">
<UL id=paipai_task_nav>
  <LI><A  href="index.asp">���Ļ�����</A> </LI>
  <LI><A   href="pushmission.asp">��������</A> </LI>
  <LI><A href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
   <LI><A href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/paipai_task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI></UL>
  </DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG 
src="../images/paipai_task_nav_line.gif"></DIV>
</DIV>
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 15px; BACKGROUND-COLOR: #ffffff">

<div class="pub_tip4">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="bottom" width="120" align="center"><strong class="f_b_org"><font color="#EC8C11">��ʾ˵��</font></strong></td>
        <td><div align="left">1.��ҳ����Ҫ�����󶨺�ά�����������񣬹���������Ʒ���Ա���ţ�<a href="http://www.778892.com/news.asp?/1415.html" target="_blank" class="link_o">ʲô����ţ�</a><br />
          2.Ϊ�˱�֤��������Լ�����������İ�ȫ���뾡���󶨶����ѡ��ţ���������ʱ�ֻ�ʹ�ã�<a href="http://www.778892.com/news.asp?/1480.html" target="_blank" class="link_o">ΪʲôҪ�󶨶����ţ�</a><br />
          3.Ϊ�˱�֤��������Լ�����������İ�ȫ��ƽ̨����ÿ���Ա����ÿ�ս�����������Ҫ����6����ÿ�ܲ�Ҫ����35��!<br />
          4.Ϊ�˱�֤ˢ�������������Ŷȣ�ϵͳΪÿ��������Զ�����һ��������ֵ�������������������ȴﵽ������ֵ���󽫲��ܽ����µ�����<br /><br /></div></td>
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
    <TD align=right height=30>���������ʺţ�</TD>
    <TD><div align="left">
      <INPUT id=hao style="WIDTH: 190px" maxLength=20 name=hao>
    ��QQ���룩</div></TD></TR>
  <TR>
    <TD height=30><div align="right">�����룺</div></TD>
    <TD><div align="left">
      <input name="czm" type="password" id="czm" size="60" style="WIDTH: 190px">
    </div></TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD>&nbsp;</TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left">
      <INPUT id=Button1 style="WIDTH: 200px; HEIGHT: 30px"  type=submit value=���ҵ����������ʺ� name=Button1>
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
            color=#ffffff>�Ѱ����������ʺ�</FONT></TD>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_6.gif" 
          width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
       </TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#E9A91E height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="100%" align="center" cellpadding="5" cellspacing="1" bgcolor="#D0DBE3" >
        <TR bgcolor="#E1F2FB">
          <TD height="30" style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">���������ʺ� </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">��ʼ�������� </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">����Ѿ����� </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">�ú�����껹�����</div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">״̬ </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 5%"><div align="center">���� </div></TD>
        </TR>
        <TBODY>
		  <%
	
			
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=2  order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("������Ϣ")
				Else
					Do While Not Rs.Eof
			  %>
          <TR>
              <TD bgcolor="#FFFFFF" class=centers><img src="../images/j_z.gif" width="13" height="16"><%=rs("shopname")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=rs("num")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%
	Sqlj = "select count(*) as jxinyu from "&jieducm&"_xinyu where username='"&session("useridname")&"' and classid=2 and shopname='"&rs("num")&"' "
	Set Rsj = Server.CreateObject("Adodb.RecordSet")
	Rsj.Open Sqlj,conn,1,1
	IF not Rsj.Eof Then
	 i=rsj("jxinyu")
	 else
	 i=0
	End IF
	response.Write(i)
	%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=200-(rs("num")+i)%>
                    <%xiaoyu=(rs("num")+i)
				  if xiaoyu>=200 then
				    sqlr="update "&jieducm&"_xinyu set start=1 where id="&rs("id")&""
                    conn.execute sqlr
				  end if
				  %></TD>
<TD bgcolor="#FFFFFF" class=centers><%if rs("start")="0" then%>���Լ���ʹ�� <%else%> �Ѵﵽ����<%end if%></TD>
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
</BODY></HTML>
<SCRIPT language=javascript>
function  save_onclick12()
{

  var Shopkeeper=document.form1.hao.value;
  if(Shopkeeper=="")
  {
    alert("�������ҵ�ˢ����ţ�");
	document.form1.hao.focus();
	return false;
	}

  return true;  
}
</SCRIPT>