<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

<!--#INCLUDE FILE="../background.asp"-->
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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
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
		document.myform.action="tidel.asp";
		if(confirm("ȷ��Ҫɾ��ѡ�еļ�¼�𣿱����������޷��ָ���"))
		    return true;
		else
			return false;
	}
	else
	{
		document.myform.action="chongdel.asp";
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
	<td colspan="2" align="center" class="title"><strong>�� �� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>��������</strong></td>
    <td><a href="?action=0">�ȴ�����</a> <a href="?action=2">�ѳ���</a> <a href="?action=3">��������</a>&nbsp;<a href="?action=1">�������</a></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
    <form name="myform" method="Post" action="tidel.asp" onSubmit="return ConfirmDel();">
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
            <td width="151" align="center"  height="22"><strong>��ˮ��</strong></td>
            <td width="163" align="center" ><strong>��������</strong></td>
            <td width="163" align="center" ><strong>�����û���</strong></td>
            <td width="123" align="center" ><strong>���ֽ��</strong></td>
            <td width="113" align="center" ><strong>��ʵ����</strong></td>
            <td width="139" align="center" ><strong>�˺�</strong></td>
            <td width="161" align="center" ><strong>״̬</strong></td>
            <td width="161" align="center" ><strong>����ʱ��</strong></td>
            <td width="161" align="center" ><strong>����</strong></td>
            <!--<td width="40" align="center" ><strong>�����</strong></td>-->
          </tr>
       
			<%	
action=request("action")
if action="" then
   	sql="SELECT * FROM "&jieducm&"_tixian  order by id desc"
elseif action="ok" then
  sql="SELECT * FROM "&jieducm&"_tixian where username ='"&request("text")&"'  order by id desc"
elseif action="0" then
   	sql="SELECT * FROM "&jieducm&"_tixian  where start='0' order by id desc"
elseif action="1" then
   	sql="SELECT * FROM "&jieducm&"_tixian  where start='1' order by id desc"
elseif action="2" then
    	sql="SELECT * FROM "&jieducm&"_tixian  where start='2' order by id desc"
elseif action="3" then
    	sql="SELECT * FROM "&jieducm&"_tixian  where start='3' order by id desc"

end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	dim maxperpage,url,PageNo
	if action="" then
	 url="admin_bankroll.asp"
	elseif action="ok" then
		 url="admin_bankroll.asp?action=ok&text="&request("text")&""
	else
		 url="admin_bankroll.asp?action="&action&""
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
            <td width="151" align="center"><%=rs("pid")%></td>
            <td width="163" align="center"><%if rs("class")=1 then%>�Ա���Ʒ��ַ����
			<%elseif rs("class")=2 then%>֧��������
			<%elseif rs("class")=3 then%>�Ƹ�ͨ����
			<%end if%></td>
            <td width="163" align="center"> <%=rs("username")%></td>
            <td width="123" align="center"><%=rs("ReNum")%>Ԫ</td>
            <td width="113" align="center"><%=rs("ReRName")%></td>        
            <td width="139" align="center"><%=rs("ReZfb")%></td>
            <td width="161" align="center">
			<%
			if(rs("start")=0) then
			   response.write("<font color=red>������</font>")
			elseif (rs("start")=1 )then
			   response.write("���ֳɹ�")
			   response.Write("<br>���׺ţ�")
			   response.Write(rs("zfbnum"))
			 elseif (rs("start")=2 )then
			   response.write("��������")
			 elseif (rs("start")=3 ) then
			    response.write("������")
			 end if
			%>            </td>
            <td width="161" align="center"><%=rs("now")%></td>
            <td width="161" align="center">
           <% if rs("start")<>2  and rs("start")<>1 then%><a href="tixianok.asp?id=<%=rs("id")%>&action=rem"> ��������</a> <% else%><font color="#999999">��������</font><% end if%>
           | <% if rs("start")<>3  and rs("start")<> 1 and rs("start")<>2 then%><a href="tixianok.asp?id=<%=rs("id")%>&action=shou">�����û�����</a><% else%><font color="#999999">�����û�����</font><% end if%> <br> 
           
           <% if rs("start")<>1 and rs("start")<>2 then%>
            <a href="tixianjiao.asp?id=<%=rs("id")%>">�������</a>
            <%else%>
            <font color="#999999">�������</font>
            <% end if%>          </td>
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
    <td height="30" colspan="2"><input name="Action" type="hidden" id="Action" value="Del">
&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %>
  </div></td></tr>
</table>
</td>
</form></tr>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="?action=ok">
                  <%  
	  set guest=nothing 
      set rs = nothing%>
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>�����û�����
                  <input 
      class=input1 size=25 name=text>
��
<input name="submit" type=submit class=input2 value=" �� �� ">
                </form> 
            </div></td>
    </tr>
          <tr>
            <td height=12>&nbsp;</td>
          </tr>
          </td>
</table>
</table>
<%
set rs=nothing
conn.close
set conn=nothing%>

</body>
</html>
