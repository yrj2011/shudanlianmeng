<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->
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
name1=request("username")
if request.Form<>"" then

 username=request("name")
 sf=request("sf")
 class1=request("class")
 fabudian=Replace(Trim(request.Form("fabudian")),"'","''")
 content=request("content")
 czm=request("czm")
 if sf=1 then
  leibie="�ٷ�����"
 else
  leibie="�ٷ�����"
 end if
 
if md5(czm)<>admin_czmpass then
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")=leibie
   rs("content")=session("AdminName")&"����Ա��"&username&"����"&leibie&"ʱ�������������"
   rs("jiegou")="ʧ��"
   rs.update
   rs.close
   Response.Write("<script>alert('����������!�벻Ҫ�ظ�������ƽ̨���¼�����Ϊ��');history.back();</script>")
   response.End()
end if

if sf=2 then
  if class1="1" then
   sqlr="update "&jieducm&"_register set cunkuan=cunkuan-"&fabudian&" where username='"&username&"'"
  elseif class1="2"  then
     sqlr="update "&jieducm&"_register set fabudian=fabudian-"&fabudian&" where username='"&username&"'"
  elseif class1="3"  then
       sqlr="update "&jieducm&"_register set jifei=jifei-"&fabudian&" where username='"&username&"'"
  end if
elseif sf=1 then
  if class1="1" then
   sqlr="update "&jieducm&"_register set cunkuan=cunkuan+"&fabudian&" where username='"&username&"'"
  elseif class1="2"  then
     sqlr="update "&jieducm&"_register set fabudian=fabudian+"&fabudian&" where username='"&username&"'"
  elseif class1="3"  then
       sqlr="update "&jieducm&"_register set jifei=jifei+"&fabudian&" where username='"&username&"'"
  end if
end if
conn.execute sqlr
  

if class1="1" then
classname="Ԫ���"
elseif class1="2" then
classname="��������"
elseif class1="3" then
classname="������" 
end if  

 	 Set rs=server.createobject("ADODB.RECORDSET")
     rs.open "Select * From "&jieducm&"_chengfa" ,Conn,3,3  
      rs.addnew
	  rs("username")=username
	  rs("class")=class1
	  rs("num")=fabudian
	  rs("now")=now()
	  rs("manage")=session("AdminName")
	  rs("content")=content
	  rs("leibie")=sf
	  rs.update
	  rs.close
	  
 	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")=leibie
		rs("content")=session("AdminName")&"����Ա"&leibie&""&username&"��"&fabudian&""&classname
		if class1=1 then
        if sf=1 then
         rs("price")="+"&fabudian
        elseif sf=2 then
		rs("price")="-"&fabudian
         end if
		else
		rs("price")=0
		end if
		rs("jiegou")="�����ɹ�"
    	rs.update
    	rs.close
		
	Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")=leibie
		rs("content")=session("AdminName")&"����Ա"&leibie&""&username&"��"&fabudian&""&classname
		rs("jiegou")=leibie&"�ͷ��ɹ�"
    	rs.update
    	rs.close		
   Response.Write("<script>alert('�����ɹ�!');window.location.href='usermanage.asp?action=sear&text="&username&"';</script>")
   response.End()
end if
%>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<script language=javascript>
<!--
function CheckForm()
{
	if(document.form.fabudian.value=="")
	{
		alert("�������");
		document.form.fabudian.focus();
		return false;
	}
  if(document.form.czm.value=="")
	{
		alert("����������룡");
		document.form.czm.focus();
		return false;
	}

}
//-->
</script>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<DIV> </DIV>

<DIV style="WIDTH: 95%; PADDING-TOP: 10px; HEIGHT: 160px">
<DIV style="PADDING-LEFT: 50px; WIDTH: 98%; HEIGHT: 20px">
  <div align="center"><SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 14px; COLOR: red">�ٷ��ͷ�����</SPAN><BR>
  </div>
</DIV>
<UL class=pushmission>
  <LI><FORM id=form name=form  action="" method="post" onSubmit="return CheckForm();">

    <table width="100%" border="0">
      <tr>
        <td width="47%"><div align="right">�ͷ��û�����</div></td>
        <td width="53%"><input type="text" name="name" id="name"  value="<%=name1%>" readonly></td>
      </tr>
      <tr>
        <td><div align="right">���</div></td>
        <td><label>
          <input name="sf" type="radio" value="1" checked="checked" />
          ����
          <input type="radio" name="sf" value="2" />����
        </label></td>
      </tr>
      <tr>
        <td><div align="right">�������:</div></td>
        <td><select name="class" id="class">
          <option value="1" selected>���</option>
          <option value="2">������</option>
          <option value="3">����</option>
          </select>          </td>
      </tr>
      <tr>
        <td><div align="right">��</div></td>
        <td><input name="fabudian" type="text" id="fabudian" onkeyup="if(isNaN(value))execCommand('undo')" maxlength="4"></td>
      </tr>
	       <tr>
        <td><div align="right">ԭ��</div></td>
        <td><label>
          <textarea name="content" rows="5"></textarea>
          </label></td>
      </tr>
	    <tr>
        <td><div align="right">�����룺</div></td>
        <td><label>
        <input type="password" name="czm" />
        </label></td>
      </tr>
      <tr>
        <td colspan="2"><div align="center">
          <input type="submit" name="button" id="button" value="�� ��">
          &nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="button" name="button2" id="button2" value="ȡ ��"  onClick="history.back();"/>
        </div></td>
        </tr>
    </table></FORM>
  </LI>
</UL>
</DIV>

</DIV>
<%
set rs=nothing
conn.close
set conn=nothing%>