<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
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
<title>七七传媒</title>
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
		alert("您至少选择1条消息!");
	}	
	else{
		if(confirm("确定要删除选择的消息吗？\n此操作不可以恢复！")){
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
          <td align="center"><font color="#FFFFFF" style="font-size:14px">手 机 消 息 管 理</font></td>
          <td width="35"><INPUT title=删除 onclick=DelAll() type=button value=删除 name=Submit></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"> <br>
        <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#D6DFF7">
          <tr bgcolor="#FFFFFF">
            <td width="30" align="center">编号</td>
            <td width="125" align="center">发送人</td>
            <td width="156" align="center">接收人</td>
            <td width="552" align="center">内容</td>
            <td width="80" align="center">发送日期</td>
            <td width="30" align="center"><input id=chkAll 
                  onClick=CheckAll(this.form) type=checkbox 
                  value=checkbox name=chkAll></td>
          </tr>
          <%if rs.eof or rs.bof then
response.write"<tr bgcolor=#FFFFFF><td colspan='10'><p align='center'><font color='red'>暂无消息！</font></td></tr></table><br>"
response.end
else
const maxperpage=20
dim currentpage
rs.pagesize=maxperpage
currentpage=request.querystring("pageid")
if isnumeric(currentpage)=false then
response.write "<script>alert('参数错误，关闭窗口！');window.close();</script>"
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
    <td height="20" bgcolor="#FFFFFF"><p align="center">页数：<%=currentpage%>/<%=n%>                                                            
   	<%k=currentpage                                                                                         
   	if k<>1 then%>                                                                     
   	<a href="?pageid=1">首页</a>                                                                                         
   	<a href="?pageid=<%=k-1%>">上一</a>                                                                                         
   	<%else%>                                                                     
   	首页&nbsp;上一页                                                                                         
   	<%end if%>                                                                     
   	<%if k<>n then%>                                                                                         
   	<a href="?pageid=<%=k+1%>">下一</a>                                                                                         
   	<a href="?pageid=<%=n%>">尾页</a>                                                                                         
   	<%else%>                                                                     
   	下一页&nbsp;尾页                                                                                         
   	<%end if%>                                                                     
     共有 <%=totalput%> 条消息     </td>                                                                                  
    
   </tr>                                                                
</table>                                                       
</body>                                                                        
</html>     