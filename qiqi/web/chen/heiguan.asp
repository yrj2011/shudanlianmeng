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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ߴ�ý</title>
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
		document.myform.action="heidelg.asp";
		if(confirm("ȷ��Ҫɾ��ѡ�еļ�¼�𣿱����������޷��ָ���"))
		    return true;
		else
			return false;
	}
	else
	{
		document.myform.action="heidelg.asp";
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
	<td colspan="2" align="center" class="title"><strong>����������</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>��������</strong></td>
    <td><a href="heiadd.asp">��ӹٷ�������</a></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
    <form name="myform" method="Post" action="heidelg.asp" onSubmit="return ConfirmDel();">
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
            <td height="22" width="109" align="center"><strong>ѡ��</strong></td>
            <td width="151" align="center"  height="22"><strong>�������û���</strong></td>
            <td width="161" align="center" ><strong>������</strong></td>
            <td width="161" align="center" ><strong>����</strong></td>
            <!--<td width="40" align="center" ><strong>�����</strong></td>-->
          </tr>
       
			<%	
   	sql="SELECT * FROM  "&jieducm&"_hei  where class='2'  order by id desc"

			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	 url="hei.asp"
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
            <td width="109" align="center"><input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>></td>
            <td width="151" align="center"><%=rs("name")%></td>
            <td width="161" align="center"><%=rs("username")%></td>
            <td width="161" align="center">&nbsp; <a href="heidelg.asp?id=<%=rs("id")%>">ɾ��</a></td>
            <!--<td width="40" align="center"> 
			'if rsArticleList("Passed")=true then
				'response.write "��"
			'else
				'response.write "��"
			'end if%></td>-->
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
&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %>
  </div></td></tr>
</table>
</td>
</form></tr>

</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com/' target="_blank"><font color=#ff6600><strong>���ߴ�ý-���߻�ˢƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>

</body>
</html>
