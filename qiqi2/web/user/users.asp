<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
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
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<style type="text/css" media="screen">
	
	@import url("../css/interiormail.css");
</style>
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; ˽��վ���� &gt;&gt; </div>
                    <div class=pp8><strong>˽��վ����</strong></div>
               
					  <FORM name=form method=post>
           
           
            <%

dim sql,rs
	
keyword=replace(trim(request.form("keyword")),"'","")
if keyword <>"" then
sql="select * from "&jieducm&"_china_message where (uid='"&session("useridname")&"' ) and zn='yes' and ( messagename like '%" & keyword & "%' or  messagelxfs like '%" & keyword & "%' or messagetext like '%" & keyword & "%' ) order by id desc"
else
sql="select * from "&jieducm&"_china_message where (uid='"&session("useridname")&"' ) and zn='yes' order by id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3%>
            <SCRIPT language=javascript>
function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
function Checked()
{
	var j = 0
	for(i=0;i < document.form.elements.length;i++){
		if(document.form.elements[i].name == "adid"){
			if(document.form.elements[i].checked){
				j++;
			}
		}
	}
	return j;
}

function DelAll()
{
	if(Checked()  <= 0){
		alert("������ѡ��1����Ϣ!");
	}	
	else{
		if(confirm("ȷ��Ҫɾ��ѡ�����Ϣ��\n�˲��������Իָ���")){
			form.action="delm.asp?del=message";
			form.submit();
		}
	}
}

</SCRIPT>
          <br>
				<a href="send_message.asp"><img src="../img/jieducm_xin_img.jpg" width="95" height="24" border="0"></a><br>
				<br>

              			  <div id="TopBox">�����ڲ鿴���ǣ� </div>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
              		<tr>
		<td colspan="8">
			<ul class="private-msg-tab">
				<li ><a href="user.asp">վ����</a></li>
                <li class="selected"><a href="users.asp ">˽��վ����</a></li>
				</ul>		</td>
               
	</tr>

                <tr>
                  <th scope="col" width="9%"align="center"><div align="left">
                    <input id=chkAll 
                  onClick=CheckAll(this.form) type=checkbox 
                  value=checkbox name=chkAll>
                  </div></th>
                  <th scope="col" width="10%"align="center"><div align="left">���</div></th>
                  <th scope="col" width="15%" align="center"><div align="left">������</div></th>
                  <th scope="col" width="31%" align="center"><div align="left">����</div></th>
                  <th scope="col" width="13%" ><div align="left">������</div></th>
                  <th scope="col" width="22%" align="center"><div align="left">����ʱ��</div></th>
                </tr>
                 <%if rs.eof or rs.bof then
response.write"<font color='red'>���������Ϣ��</font>"
'response.end
else
const maxperpage=20
dim currentpage
rs.pagesize=maxperpage
currentpage=request.querystring("pageid")
if isnumeric(currentpage)=false then
response.write "<script>alert('�������󣬹رմ��ڣ�');window.close();</script>"
response.end
end if
if currentpage="" then
currentpage=1
elseif currentpage<1 then
currentpage=1
else
currentpage=clng(currentpage)
	if currentpage > rs.pagecount then
	currentpage=rs.pagecount
	end if
end if

dim totalput,n
totalput=rs.recordcount
if totalput mod maxperpage=0 then
n=totalput\maxperpage
else
n=totalput\maxperpage+1
end if
if n=0 then
n=1
end if
rs.move(currentpage-1)*maxperpage
i=0
do while i< maxperpage and not rs.eof
messagetext=rs("messagetext")

%>
                	    				    			<tr>
                	    				    			  <td><%if rs("uid")="all" then%>
                <b><font color="#FF0000">!</font></b><%else%><input type="checkbox" name="adid" value="<%=rs("id")%>" onClick=Checked(form)><%end if%></td>
	    				<td><%=i+currentpage*maxperpage-maxperpage+1%></td>					
	    				<td><%=rs("messagename")%></td>
	    				<td>
                           <a href="zn_messagelook.asp?id=<%=rs("id")%>" ><% if rs("hits")=0 then%><font color="#000000" size="2 px;"><%else%> <font color="#0066CC"><%end if%><strong><%=rs("messagelxfs")%></strong></font></a>
                        </td>
	    				<td> <%if rs("uid")="all" then%>
                ȫ����Ա 
                <%else%>
                <%=rs("uid")%> 
                <%end if%></td>
	    				<td><%=rs("messagedate")%></td>
	    			</tr>
                    
						 <%i=i+1                     
    rs.movenext                                                               
    loop    
                                                                                       
    end if%>							               
                         <tr>
                           <td colspan="8" class="w">&nbsp; <div class="PageButton">
				  	<INPUT title=ɾ�� onclick=DelAll() type=button value=ɾ�� name=Submit>
&nbsp;&nbsp;				  </div>	</td>
                         </tr>
                <tr>
                  <td colspan="8" class="w">				 			  </td>
                </tr>
                
                              <tr>
                  <td colspan="8">
		  <!--AdForward Begin:-->
		  <!--AdForward End--></td> 
		                  </tr>
              </table>
      </form>
               <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="CAE2F8">  
   <tr>                                                     
    <td height="28" bgcolor="#FFFFFF">
      <p align="center">ҳ����<%=currentpage%>/<%=n%> 
        <%k=currentpage                                                                                         
   	if k<>1 then%>
        <a href="?pageid=1">��ҳ</a> <a href="?pageid=<%=k-1%>">��һ</a> 
        <%else%>
        ��ҳ&nbsp;��һҳ 
        <%end if%>
        <%if k<>n then%>
        <a href="?pageid=<%=k+1%>">��һ</a> <a href="?pageid=<%=n%>">βҳ</a> 
        <%else%>
        ��һҳ&nbsp;βҳ 
        <%end if%>
        ���� <%=totalput%> ����Ϣ 
    </td>                                                                                  
    <form action="" method="post" name="search"><td width="240" align="center" bgcolor="#FFFFFF">�ؼ���
      <input  maxLength="20" name="keyword" onFocus="this.value=''" size="18"  value="<%=keyword%>"> 
    <input  type="submit" value="����" style="font-size: 12px" name="search"></td>
    </form>
   </tr>                                                                
</table>   </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
