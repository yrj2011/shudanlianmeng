<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

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
dim sql,rs
keyword=replace(trim(request.form("keyword")),"'","")
if keyword <>"" then
sql="select * from "&jieducm&"_sms where ( messagename like '%" & keyword & "%' or  messagelxfs like '%" & keyword & "%' or messagetext like '%" & keyword & "%' ) and zn='yes'order by id desc"
else
sql="select * from "&jieducm&"_sms  order by id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ߴ�ý</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>
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
			form.action="delm_phone.asp?action=del";
			form.submit();
		}
	}
}

</SCRIPT>
<FORM name=form method=post>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td height="20" bgcolor="#799AE1" align="center">
      <table width="98%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center"><font color="#FFFFFF" style="font-size:14px">�� �� �� Ϣ �� ��</font></td>
          <td width="35"><INPUT title=ɾ�� onclick=DelAll() type=button value=ɾ�� name=Submit></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"> <br>
        <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#D6DFF7">
          <tr bgcolor="#FFFFFF">
            <td width="30" align="center">���</td>
            <td width="125" align="center">������</td>
            <td width="156" align="center">������</td>
            <td width="552" align="center">����</td>
            <td width="80" align="center">��������</td>
            <td width="30" align="center"><input id=chkAll 
                  onClick=CheckAll(this.form) type=checkbox 
                  value=checkbox name=chkAll></td>
          </tr>
          <%if rs.eof or rs.bof then
response.write"<tr bgcolor=#FFFFFF><td colspan='10'><p align='center'><font color='red'>������Ϣ��</font></td></tr></table><br>"
response.end
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
do while i< maxperpage and not rs.eof%>
          <tr bgcolor="#FFFFFF">
            <td><p align="center"><%=i+currentpage*maxperpage-maxperpage+1%></td>
            <td valign="top"><%=rs("username")%> </td>
            <td valign="top"><%=rs("susername")%></td>
            <td valign="top"><%=rs("content")%></td>
            <td><p align="center"><%=rs("now")%></td>
            <td align="center"><input type="checkbox" name="adid" value="<%=rs("id")%>" onClick=Checked(form)></td>
          </tr>
          <%i=i+1                     
    rs.movenext                                                               
    loop    
    rs.close  
    set rs=nothing                                                                                     
    conn.close   
    set conn=nothing                                                                                        
    end if%>
        </table>
        <br>
    </td>
  </tr>
</table></FORM>                                                             
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">  
   <tr>                                                     
    <td height="20" bgcolor="#FFFFFF"><p align="center">ҳ����<%=currentpage%>/<%=n%>                                                            
   	<%k=currentpage                                                                                         
   	if k<>1 then%>                                                                     
   	<a href="?pageid=1">��ҳ</a>                                                                                         
   	<a href="?pageid=<%=k-1%>">��һ</a>                                                                                         
   	<%else%>                                                                     
   	��ҳ&nbsp;��һҳ                                                                                         
   	<%end if%>                                                                     
   	<%if k<>n then%>                                                                                         
   	<a href="?pageid=<%=k+1%>">��һ</a>                                                                                         
   	<a href="?pageid=<%=n%>">βҳ</a>                                                                                         
   	<%else%>                                                                     
   	��һҳ&nbsp;βҳ                                                                                         
   	<%end if%>                                                                     
     ���� <%=totalput%> ����Ϣ     </td>                                                                                  
    
   </tr>                                                                
</table>                                                       
</body>                                                                        
</html>     