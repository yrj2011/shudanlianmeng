<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
classid=request("classid")
id=request("id")
if classid="jieducm_p" then
   Sql= "select * from "&jieducm&"_pay where id="&id&"  and jieducm_start=0"
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,1,1
	 if not rs.eof then
	    car_username=rs("jieducm_username")
	    car_num=rs("jieducm_num")
		car_price=rs("jieducm_price")
	    car_maijia=rs("jieducm_maijia")
		car_start=rs("jieducm_start")
	else
		 Response.Write("<script>alert('�޴���Ϣ!');history.back();</script>")
        response.End()
	 end if 

sqlr="delete  from "&jieducm&"_pay where id='"&id&"'"
conn.execute sqlr
 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select fabudian From "&jieducm&"_register where username='"&car_username&"'" ,Conn,3,3 
rs("fabudian")=rs("fabudian")+car_num
rs.update
rs.close

 
	    nums=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=car_username
    	rs("num")=nums
		rs("class")="ȡ�����۷�����"
		rs("content")="����Ա���к�̨ȡ����������,�ѳɹ����ص���Ա��"&car_username&" �˺���"&car_num&"��������"
		rs("price")=0
		rs("jiegou")="ȡ���ɹ�"
    	rs.update
    	rs.close
		
	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")="ȡ�����۷�����"
		rs("content")="����Ա���к�̨ȡ����Ա"&car_username&"����"&car_num&"��������"
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close
		
Response.Write("<script>alert('ȡ���ɹ���');window.location.href='jieducm_pay.asp';</script>")
response.End()		
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>

<body>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel()
{
	if(document.myform.Action.value=="Del")
	{
		document.myform.action="jieducm_paydel.asp";
		if(confirm("ȷ��Ҫɾ��ѡ�еļ�¼�𣿱����������޷��ָ���"))
		    return true;
		else
			return false;
	}
	else
	{
		document.myform.action="jieducm_paydel.asp";
		if(document.myform.TargetClassID.value=="")
		{
			alert("���ܽ������ƶ�����������Ŀ����Ŀ���ⲿ��Ŀ�У�");
			return false;
		}
		if(confirm("ȷ��Ҫ��ѡ�е������ƶ���ָ������Ŀ��"))
		    return true;
		else
			return false;
	}
}

</SCRIPT>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong>�����㽻���г�</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>��������</strong></td>
    <td>&nbsp;&nbsp;<a href="?action=1"><strong> �ѳ���</strong></a>&nbsp; <a href="?action=2"><strong>������ </strong></a></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
    <form name="myform" method="Post" action="jieducm_paydel.asp" onSubmit="return ConfirmDel();">
      <table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
        <tr>
          <td colspan="2" align="left" class="title"></td>
        </tr>
        <tr valign="middle" class="tdbg">
          <td height="22"></td>
          <td width="200" height="22" align="right"></td>
        </tr>
      </table>
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22">
            <td width="158" align="center" ><strong>��ˮ��</strong></td> 
            <td width="158" height="22" align="center" ><strong>�û���</strong></td>
            <td width="129" align="center" ><strong> �۸�/����</strong></td>
            <td width="211" align="center" ><strong>ָ�����</strong></td>
            <td width="169" align="center" ><strong>���ʱ��</strong></td>
            <td width="193" align="center" ><strong>���</strong></td>
            <td width="193" align="center" ><strong>ƽ���۸�</strong></td>
            <td width="177" align="center" ><strong>״̬</strong></td>
            <td width="198" align="center" ><strong>����</strong></td>
            <!--<td width="40" align="center" ><strong>�����</strong></td>-->
          </tr>
<%	
action=request.QueryString("action")
if action="" then
sql="SELECT * FROM "&jieducm&"_pay  order by id desc"
elseif action=1 then
sql="SELECT * FROM "&jieducm&"_pay where jieducm_start=1  order by id desc"
elseif action=2 then
sql="SELECT * FROM "&jieducm&"_pay where jieducm_start=0 order by id desc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	if action="" then
	 url="jieducm_pay.asp"
	 else
	  url="jieducm_pay.asp?action="&action&""
	 end if
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
		  response.Write("�Բ���û������Ҫ��ҳ��")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("û����һҳ!")
		Response.End
	 End if
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	 DO WHILE NOT rs.EOF AND RowCount>0%>
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
		    <td width="158" align="center">
			<%if rs("jieducm_start") =0 then%>
			<%else%>
			<input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>>
			<%end if%>
			<%=rs("id")%></td>
            <td width="158" align="center"><%=rs("jieducm_username")%></td>
            <td width="129" align="center"><%=formatnumber(rs("jieducm_price"),2,-1)%>/(<%=rs("jieducm_num")%>)��</td>
            <td width="211" align="center"><%if rs("jieducm_maijia") ="" then%>��<%else%><%=rs("jieducm_maijia")%><%end if%></td>        
            <td width="169" align="center"><%=rs("jieducm_datatime")%></td>
            <td width="193" align="center"><%=rs("jieducm_maiuseranme")%></td>
            <td width="193" align="center"><%=formatnumber((rs("jieducm_price")/rs("jieducm_num")),2,-1)%>Ԫ/��</td>
            <td width="177" align="center"><%if rs("jieducm_start") =0 then%><font color="#0000FF">������</font><%else%><font color="#FF0000">�ѳ���</font><%end if%></td>          
            <td width="198" align="center">
			<%if rs("jieducm_start") =0 then%>
			<a href="jieducm_pay.asp?id=<%=rs("id")%>&classid=jieducm_p"  onClick="return confirm('ȷ��Ҫȡ��������');" title="ȡ����¼�����۵ķ����㷵��������">ȡ���˼�¼</a>
			<%else%>
			<a href="jieducm_paydel.asp?id=<%=rs("id")%>"  onClick="return confirm('ȷ��Ҫɾ��������');">ɾ��</a><%end if%></td>   
          </tr>
    <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          </table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td height="30" colspan="2"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
              ѡ�б�ҳ��ʾ�����м�¼
<input name="submit" type='submit' value='&nbsp;ɾ��ѡ���ļ�¼&nbsp;' onClick="document.myform.Action.value='Del'"  style="cursor: hand;background-color: #cccccc;">
              <input name="Action" type="hidden" id="Action" value="Del">
&nbsp;ֻ��ɾ���ѳ��۵ļ�¼&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td  ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %>
  </div></td></tr>
</table>
</td>
</form></tr>

</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>���ߴ�ý-���߻�ˢƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>


</body>
</html>
