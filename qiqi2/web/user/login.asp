<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ**************************************
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
action=request.form("action")
id=request.Form("id")
if action="ok" then	

	 Sql= "select * from "&jieducm&"_vipcar where id="&id&""
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,1,1
	 if not rs.eof then
	    car_name=rs("car_name")
		car_price=rs("car_price")
	    car_date=rs("car_date")
		car_fabudian=rs("car_fabudian")
		rs.close
	else
		Response.Write("<script>alert('������!');history.back();</script>")
       response.End()
	 end if 
	 
if int(cunkuan)<int(car_price) then
	Response.Write("<script>alert('���Ĵ�����!');history.back();</script>")
    response.End()
elseif vip=1 then
	Response.Write("<script>alert('���Ѿ���VIP�û���!');history.back();</script>")
	response.End()
end if

      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where vip=0 and username='"&session("useridname")&"'" ,Conn,3,3  
	  if not rs.eof then  
		rs("cunkuan")=rs("cunkuan")-car_price
		rs("fabudian")=rs("fabudian")+car_fabudian
    	rs("vip")=1
		rs("vipdel")=now()
		rs("vipdate")=car_date
    	rs.update
    	rs.close
	  end if
	  
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="����vip��"
		rs("username")=session("useridname")
    	rs("cunkuan")=car_price
    	rs.update
    	rs.close 
	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="����vip��"
		rs("username")=session("useridname")
    	rs("cunkuan")=car_price
		rs("fabudian")=car_fabudian
    	rs.update
    	rs.close 
	 if tjr<>"" then

	   sqlr="update "&jieducm&"_register set fabudian=fabudian+"&stupvip_car_jieducm&" where username='"&tjr&"'"
       conn.execute sqlr
	   
	   num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	   Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=tjr
    	rs("num")=num
		rs("class")="�ƹ㽱��"
		rs("content")="���Ƽ��Ļ�Ա"&session("useridname")&"ȷ������VIP:"&car_name&"ϵͳ������"&stupvip_car_jieducm&"��������"
		rs("price")=0
		rs("jiegou")="�����ɹ�"
    	rs.update
    	rs.close	
	 end if
	  
	  num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	  Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="����VIP�û�"
		rs("content")=session("useridname")&"ȷ������VIP:"&car_name&"�û���Ч����:"&car_date&"��,��"&car_price&"Ԫ���,�ͷ�����:"&car_fabudian&"��"
		rs("price")="-"&car_price
		rs("jiegou")="����ɹ�"
    	rs.update
    	rs.close		
Response.Write("<script>alert('��ϲ����Ϊ��VIP�û�!�����ڿ��Կ�ʼ����VIP֮����!');window.location.href='index.asp';</script>")	
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
      <td height="5"></td>
    </tr>
</table>
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; ����VIP &gt;&gt; </div>
                    <div class=pp8><strong>����VIP</strong></div>
                    <br>
                   
					 <DIV style="CLEAR: both; WIDTH: 710px;  BACKGROUND-COLOR: #f3f8fe">
<DIV 
style="CLEAR: both; PADDING-RIGHT: 30px; PADDING-LEFT: 30px; HEIGHT: 100%">

<DIV> </DIV>
<%
   	sql="SELECT * FROM "&jieducm&"_vipcar  order by car_sort"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	do while not rs.eof
%>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; PADDING-BOTTOM: 5px; PADDING-TOP: 15px">
<DIV style="FLOAT: left; WIDTH: 350px"><IMG src="../<%=rs("car_pic")%>"></DIV>
<form name="Form2" method="post" action="" id="Form2">
<input name="action" type="hidden" value="ok">
<input name="id" type="hidden" value="<%=rs("id")%>">
<DIV style="FLOAT: left; WIDTH: 200px; LINE-HEIGHT: 150%; PADDING-TOP: 20px">
<SPAN style="COLOR: red"><%=rs("car_content")%></SPAN><BR>
��Ч�ڣ�<%=rs("car_date")%>��
<br>
�ͷ����㣺<%=rs("car_fabudian")%>��<BR>
<%=rs("car_name")%>��<%=rs("car_price")%>Ԫ<BR>
<INPUT id=ImageButton2  onClick="return confirm('��ȷ��Ҫ����˿���');" type=image src="../images/buy1.png" name=ImageButton2>
</DIV>
</form>
</DIV>
<%
Rs.MoveNext
Loop
end if%>

<DIV></DIV>
</DIV></DIV>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
