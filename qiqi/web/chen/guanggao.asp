<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
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
if session("AdminName")="" then
   response.redirect("admin_login.asp")
end if%>
<%
If request("action")<>"" then
	
	title1=trim(replace(request.form("title1"),"'","''"))
	pic1=trim(replace(request.form("pic1"),"'","''"))
	link1=trim(replace(request.form("link1"),"'","''"))
	
	title2=trim(replace(request.form("title2"),"'","''"))
	pic2=trim(replace(request.form("pic2"),"'","''"))
	link2=trim(replace(request.form("link2"),"'","''"))
	
	title3=trim(replace(request.form("title3"),"'","''"))
	pic3=trim(replace(request.form("pic3"),"'","''"))
	link3=trim(replace(request.form("link3"),"'","''"))
	
	title4=trim(replace(request.form("title4"),"'","''"))
	pic4=trim(replace(request.form("pic4"),"'","''"))
	link4=trim(replace(request.form("link4"),"'","''"))
	
set rs=server.createobject("adodb.recordset")
sql="select * from "&jieducm&"_indexpic "
rs.open sql,conn,1,3
if request("action")="add" then 
rs.addnew
rs("title1") = title1
rs("pic1")	= pic1
rs("link1")	= link1
rs("title2") = title2
rs("pic2")	= pic2
rs("link2")	= link2
rs("title3") = title3
rs("pic3")	= pic3
rs("link3")	= link3
rs("title4") = title4
rs("pic4")	= pic4
rs("link4")	= link4
rs.update
response.write "<Script>alert('��淢���ɹ����뷵�ز鿴��');location.replace('guanggao.asp');</Script>"

elseif request("action")="modi" then
rs("title1") = title1
rs("pic1")	= pic1
rs("link1")	= link1
rs("title2") = title2
rs("pic2")	= pic2
rs("link2")	= link2
rs("title3") = title3
rs("pic3")	= pic3
rs("link3")	= link3
rs("title4") = title4
rs("pic4")	= pic4
rs("link4")	= link4
rs.update
response.write "<Script>alert('����޸ĳɹ����뷵�ز鿴��');location.replace('guanggao.asp');</Script>"
elseif request("action")="del" then
   sql3="delete * from "&jieducm&"_indexpic "
	conn.execute sql3
	response.write "<Script>alert('���ɾ���ɹ���');location.replace('guanggao.asp');</Script>"
end if

rs.close
set rs=nothing


end if
%>
<LINK href="css.css" rel=stylesheet>
<style type="text/css">
<!--
body {
	background-color: #CED7F7;
}
body,td,th {
	font-size: 12px;
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
-->
</style>
<table width="760" height="1" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="1" background="images/hline_bg.gif"></td>
  </tr>
</table>
<table width="760" height="932" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="571" height="932" align="center" valign="top"> 
	 <%
  set rs1=server.createobject("adodb.recordset")
sql1="select * from "&jieducm&"_indexpic"
rs1.open sql1,conn,1,3
if rs1.eof and rs1.bof then
%>
  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="ccccccc" >
  <form  method=post action="guanggao.asp?action=add" name=form1>
    <tr bgcolor="#FFFFFF">
      <td height="0" align="right" valign="bottom" bgcolor="#CED7F7">&nbsp;</td>
      <td valign="bottom" bgcolor="#CED7F7"><font color="#FF0000">(ͼƬ�ϴ��ߴ��556*��220����.֧��jpg,gif��ʽ��ͼƬ)</font> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ1˵��</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="title1" type="text" id="url1" style="font-size:9pt;" size="50">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7">�ϴ�ͼƬ1</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><input name="pic1" type="text"   size="35" >
	   
       
       <input type="button" name="Submit" value="�ϴ�ͼƬ1" onClick="window.open('../upload_flash.asp?formname=form1&editname=pic1&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	  &nbsp; 	  </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ1������ַ</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link1" type="text" id="link1" style="font-size:9pt;" size="40">
        <font size="2">(������http://��ͷ)      </font></font></td>
    </tr>
	
	    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ2˵��</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="title2" type="text" id="url2" style="font-size:9pt;" size="50">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7">�ϴ�ͼƬ2</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><input name="pic2" type="text"   size="35" >
	   <input type="button" name="Submit" value="�ϴ�ͼƬ2" onClick="window.open('../upload_flash.asp?formname=form1&editname=pic2&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
&nbsp;	  </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ2������ַ</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link2" type="text" id="link2" style="font-size:9pt;" size="40">
        <font size="2">(������http://��ͷ)       </font></font></td>
    </tr>
	
	    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ3˵��</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="title3" type="text" id="url3" style="font-size:9pt;" size="50">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7">�ϴ�ͼƬ3</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><input name="pic3" type="text"   size="35" >
	   <input type="button" name="Submit" value="�ϴ�ͼƬ3" onClick="window.open('../upload_flash.asp?formname=form1&editname=pic3&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	   &nbsp;&nbsp; </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ3������ַ</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link3" type="text" id="link3" style="font-size:9pt;" size="40">
        <font size="2">(������http://��ͷ)       </font></font></td>
    </tr>
	
	    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ4˵��</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="title4" type="text" id="url4" style="font-size:9pt;" size="50">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7">�ϴ�ͼƬ4</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><input name="pic4" type="text"   size="35" >
	   <input type="button" name="Submit" value="�ϴ�ͼƬ4" onClick="window.open('../upload_flash.asp?formname=form1&editname=pic4&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	   &nbsp;&nbsp; </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7">ͼƬ4������ַ</td>
      <td width="83%" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link4" type="text" id="link4" style="font-size:9pt;" size="40">
        <font size="2">(������http://��ͷ)       </font></font></td>
    </tr>
    <tr> 
      <td height="13" align="right" bgcolor="#CED7F7">&nbsp;</td>
      <td valign="top" bgcolor="#CED7F7"><input type="submit" name="Submit22" value="����ӡ�"  style="font-size:9pt;color:#333333;"></td>
    </tr>
  </form>
</table> 
<%else%>
  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="ccccccc" >
  <form  method=post action="guanggao.asp?action=modi" name=frm1>
    <tr bgcolor="#FFFFFF">
      <td height="0" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">&nbsp;</font></td>
      <td colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#FF0000">(ͼƬ�ϴ��ߴ��556*��220����.֧��jpg,gif��ʽ��ͼƬ)</font> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ1˵��</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="title1" type="text" id="url1" style="font-size:9pt;" size="50" value="<%=rs1("title1")%>">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">�ϴ�ͼƬ1</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><input name="pic1" type="text"   size="35" value="<%=rs1("pic1")%>">
	   <input type="button" name="Submit" value="�ϴ�ͼƬ1" onClick="window.open('../upload_flash.asp?formname=frm1&editname=pic1&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	   &nbsp;&nbsp; </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ1������ַ</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link1" type="text" id="link1" style="font-size:9pt;" size="40"  value="<%=rs1("link1")%>">
        <font size="2">(������http://��ͷ)       </font></font></td>
    </tr>
	
	    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ2˵��</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="title2" type="text" id="url2" style="font-size:9pt;" size="50" value="<%=rs1("title2")%>">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">�ϴ�ͼƬ2</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><input name="pic2" type="text"   size="35"  value="<%=rs1("pic2")%>">
	   <input type="button" name="Submit" value="�ϴ�ͼƬ2" onClick="window.open('../upload_flash.asp?formname=frm1&editname=pic2&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	   &nbsp;&nbsp; </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ2������ַ</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link2" type="text" id="link2" style="font-size:9pt;" size="40" value="<%=rs1("link2")%>">
        <font size="2">(������http://��ͷ)       </font></font></td>
    </tr>
	
	    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ3˵��</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="title3" type="text" id="url3" style="font-size:9pt;" size="50" value="<%=rs1("title3")%>">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">�ϴ�ͼƬ3</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><input name="pic3" type="text"   size="35" value="<%=rs1("pic3")%>">
	   <input type="button" name="Submit" value="�ϴ�ͼƬ3" onClick="window.open('../upload_flash.asp?formname=frm1&editname=pic3&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	   &nbsp;&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ3������ַ</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link3" type="text" id="link3" style="font-size:9pt;" size="40" value="<%=rs1("link3")%>">
        <font size="2">(������http://��ͷ)</font> </font></td>
    </tr>
	
	    <tr bgcolor="#FFFFFF"> 
      <td width="17%" height="0" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ4˵��</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
      <input name="title4" type="text" id="title4" style="font-size:9pt;" size="50" value="<%=rs1("title4")%>" >
</font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="1" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">�ϴ�ͼƬ4</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><input name="pic4" type="text"   size="35"  value="<%=rs1("pic4")%>">
	   <input type="button" name="Submit" value="�ϴ�ͼƬ4" onClick="window.open('../upload_flash.asp?formname=frm1&editname=pic4&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
	   &nbsp;&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="2" align="right" valign="bottom" bgcolor="#CED7F7"><font size="2">ͼƬ4������ַ</font></td>
      <td width="83%" colspan="2" valign="bottom" bgcolor="#CED7F7"><font color="#666666">
        <input name="link4" type="text" id="link4" style="font-size:9pt;" size="40" value="<%=rs1("link4")%>">
        <font size="2">(������http://��ͷ) </font> </font></td>
    </tr>
    <tr> 
      <td height="13" align="right" bgcolor="#CED7F7">&nbsp;</td>
      <td valign="top" bgcolor="#CED7F7"><input type="submit" name="Submit22" value="���޸ġ�"  style="font-size:9pt;color:#333333;"></td>
      <td valign="top" bgcolor="#CED7F7"></td>
    </tr>
  </form>
</table> 
<% end if
conn.close
set conn=nothing
%>

    </td>
  </tr>
</table>


