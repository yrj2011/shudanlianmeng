<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../md5.asp"-->
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

if stupphonestart=1 and dst =0 then
	Response.Write("<script>alert('���Ȱ��ֻ����ܽ�����������!');window.location.href='../user/Center_Userlock.asp';</script>")
	response.End()
end if
				
				
action=HtmlEncode(trim(request.form("action")))
if action="1" then
  geshu=(trim(request.form("ToUser")))
if stupfhc="0" then
 	 Response.Write("<script>alert('������һ��ֽ��Ա�����Ա�رգ�\n����һ�������ϵ�ͷ�!');history.back();</script>")
	 response.End()
elseif  isdun=1 then
	Response.Write("<script>alert('���ѱ�����Ա��ֹ�˶һ�����!');history.back();</script>")
	response.End()
elseif fabudian<stupfhcshu then
 	Response.Write("<script>alert('���������"&stupfhcshu&"���ܶһ�!');history.back();</script>")
	response.End()		
elseif(int(fabudian)<int(geshu)) then
    Response.Write("<script>alert('�㷢���㲻��,���ܶһ�!');history.back();</script>")
	response.End()	
elseif czm<>request("czm") then
	Response.Write("<script>alert('�����벻��ȷ!');history.back();;</script>")
	response.End()
elseif int(geshu)<1 then
	Response.Write("<script>alert('��������!');history.back();;</script>")
	response.End() 
elseif not isNumeric(geshu) then
    Response.Write("<script>alert('�����������������룡');history.back();</script>")
    response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('�޴˻�Ա');window.location.href='moneyorpush.asp';</script>")
response.End()
else
rs("fabudian")=rs("fabudian")-geshu
rs("cunkuan")=rs("cunkuan")+(geshu*stupfa)
rs.update
rs.close
 end if	  

	  num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�����㻻���"
		rs("content")=session("useridname")&"��"&geshu&"�������㻻���"
		rs("price")=geshu*stupfa
		rs("jiegou")="�����㻻���ɹ�"
    	rs.update
    	rs.close		
		call check_jieducm_name(session("useridname"))   			
    Response.Write("<script>alert('�һ��ɹ�!');window.location.href='moneyorpush.asp';</script>")
	response.End()	
    set rs=nothing
elseif action="3" then
     give1=HtmlEncode(trim(request.form("givenum")))
	 touser=HtmlEncode(trim(request.form("touser")))
	 give=give1+give1*(stupshouxu/100)

	if isgive=1 then
		 Response.Write("<script>alert('���ѱ�����Ա��ֹ�����͹���!');history.back();;</script>")
	     response.End()
	elseif czm<>request("czm") then
		Response.Write("<script>alert('�����벻��ȷ!');history.back();;</script>")
	     response.End()
	elseif((fabudian)<=(give)) then
		 Response.Write("<script>alert('�����͵ķ����㲻�ܴ�����ķ�����!');history.back();</script>")
		response.End() 
	elseif int(give1)<1 then
	   Response.Write("<script>alert('��������!');history.back();</script>")
		response.End() 
	elseif not isNumeric(give1) then
          Response.Write("<script>alert('�����������������룡');history.back();</script>")
          response.End()
	 end if	
	 
	 Set rs=server.createobject("ADODB.RECORDSET")
	 rs.open "Select * From "&jieducm&"_register where username='"&touser&"'" ,Conn,3,3  
	   if rs.eof then
	       Response.Write("<script>alert('�޴˻�Ա');history.back();</script>")
		   response.End()
       elseif session("useridname")=touser then
	      Response.Write("<script>alert('�������͸��Լ�������!');history.back();</script>")
		  response.End()
	   else
		rs("fabudian")=rs("fabudian")+give1
    	rs.update
    	rs.close
	  end if
	  
if isgives=1 then
sqlr="update "&jieducm&"_register set fabudian=fabudian-"&give1&"  where username='"&session("useridname")&"'"
conn.execute sqlr
else	  
sqlr="update "&jieducm&"_register set fabudian=fabudian-"&give&"  where username='"&session("useridname")&"'"
conn.execute sqlr
end if		
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="����������"
		rs("content")=session("useridname")&"���͸�"&touser&"��Ա"&give1&"��������"
		rs("price")=0
		rs("jiegou")="���������ͳɹ�"
    	rs.update
    	rs.close
 				
	Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=touser
    	rs("num")=num
		rs("class")="����������"
		rs("content")=touser&"�õ�"&session("useridname")&"��Ա���͵�"&give1&"��������"
		rs("price")=0
		rs("jiegou")="���������ͳɹ�"
    	rs.update
    	rs.close
		call check_jieducm_name(session("useridname"))   
    Response.Write("<script>alert('���ͳɹ�!');window.location.href='moneyorpush.asp';</script>")	
   response.End()

elseif action="5" then
	 ReNum1=request.form("ReNum1")
	if isdun=1 then
	 Response.Write("<script>alert('���ѱ�����Ա��ֹ�˶һ�����!');history.back();;</script>")
	 response.End()
	elseif int(jifei)<int(ReNum1)*int(stupjifei) then
	 Response.Write("<script>alert('����ֲ���,���ܶһ�!');history.back();</script>")
	 response.End()
	elseif czm<>request("czm") then
	  Response.Write("<script>alert('�����벻��ȷ!');history.back();</script>")
	  response.End()
	 elseif int(ReNum1)<1 then
	 	  Response.Write("<script>alert('��������!');history.back();</script>")
	  response.End()
	 	elseif not isNumeric(ReNum1) then
          Response.Write("<script>alert('�����������������룡');history.back();</script>")
          response.End()
	end if
	
	  Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
	   if rs.eof then
	       Response.Write("<script>alert('�޴˻�Ա');history.back();</script>")
		   response.End()
	   else
	    rs("jifei")=rs("jifei")-ReNum1*stupjifei
		rs("fabudian")=rs("fabudian")+ReNum1
    	rs.update
    	rs.close
	  end if
	 
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="���ֻ�������"
		rs("content")=session("useridname")&"���ֻ�"&ReNum1&"����������ּ�����"&ReNum1*stupjifei
		rs("price")=0
		rs("jiegou")="���ֻ�������ɹ�"
    	rs.update
    	rs.close	
		call check_jieducm_name(session("useridname"))   	
    Response.Write("<script>alert('�һ��ɹ�!');window.location.href='moneyorpush.asp';</script>")
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
 <SCRIPT language=javascript>
function save_onclick()
{
    var Price=document.formt.ToUser.value;
  if(Price=="")
  {
    alert("������һ�������");
	document.formt.ToUser.focus();
	return false;
	}
if   ((document.formt.ToUser.value.indexOf("-")   ==   0)||!(document.formt.ToUser.value.indexOf(".")   ==   -1)){   
  alert("�һ���������ΪС������");   
  document.formt.ToUser.focus();   
  return   false;   
  } 
if(Price<10)
  {
    alert("10����������һ���");
	document.formt.ToUser.focus();
	return false;
	}

var code=document.formt.czm.value;
  if(code=="")
  {
    alert("�����벻��Ϊ�գ�");
	document.formt.czm.focus();
	return false;
	}
  return true;  
}


function  save_onclick3()
{

    var ToUser=document.form3.ToUser.value;
  if(ToUser=="")
  {
    alert("������Ҫ���͸���λ��Ա��");
	document.form3.ToUser.focus();
	return false;
	}	
	
  var GiveNum=document.form3.GiveNum.value;
  if(GiveNum=="")
  {
    alert("����������������");
	document.form3.GiveNum.focus();
	return false;
	}  


 if(!Number(GiveNum))
	  
  {
    alert("���ͽ��������0,����������!");
	document.form3.GiveNum.focus();
	return false;
	}
if   ((document.form3.GiveNum.value.indexOf("-")   ==   0)||!(document.form3.GiveNum.value.indexOf(".")   ==   -1)){   
  alert("���ͽ���ΪС������");   
  document.form3.GiveNum.focus();   
  return   false;   
  }  
	

  var code=document.form3.czm.value;	
	if(code=="")
  {
    alert("�����벻��Ϊ�գ�");
	document.form3.czm.focus();
	return false;
	}

  return true;  
}


function  save_onclick5()
{
    var Price=document.form5.ReNum1.value;
  if(Price=="")
  {
    alert("������һ�������");
	document.form5.ReNum1.focus();
	return false;
	}
 if(!Number(Price))
	  
  {
    alert("������������!");
	document.form5.ReNum1.focus();
	return false;
	}
if   ((document.form5.ReNum1.value.indexOf("-")   ==   0)||!(document.form5.ReNum1.value.indexOf(".")   ==   -1)){   
  alert("�һ���������ΪС������");   
  document.form5.ReNum1.focus();   
  return   false;   
  } 
 var czm=document.form5.czm.value;
  if(czm=="")
  {
    alert("���������Ĳ����룡");
	document.form5.czm.focus();
	return false;
	}
  
  return true;  
}

</script>   
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;��Ҫ�һ� &gt;&gt; </div>
                    <div class=pp8><strong>��Ҫ�һ�</strong></div>
                    <div style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
                      <ul>
                        <li>* ������Ƽ�����Խ�࣬����ø���Ļ��֣�<br>
* �һ��ɹ���ϵͳ�Զ���Ӧ�Ĳ�����
  <br>
  * ���Ĳ�����¼��Ҳ������Ӧ�ļ�¼��Ϣ </li>
                      </ul>
                    </div><br><br>
					  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;������һ����</STRONG></td>
                </tr>
              </table>
              <form action="" method="POST" name="formt" id="formt" onSubmit="return save_onclick()">
                      <table width="639" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="35" align="right"><div align="right">�һ��ˣ�</div></td>
                          <td><%=session("useridname")%> �����ڵķ������ǣ�<%=fabudian%> </td>
                        </tr>
                        <tr>
                          <td width="134" height="35" align="right"><div align="right">�����ã�</div></td>
                          <td width="505"><input id=ToUser  name=ToUser  onKeyUp="if(isNaN(value))execCommand('undo')">
                            �����������һ�(ÿ1����������Զһ�
							<%if stupfa<1 then 
							response.Write("0")
							response.Write(stupfa)
                            else
							response.Write(stupfa)
							end if%>Ԫ,<%=stupfhcshu%>����������һ�)</td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">�����룺</div></td>
                          <td><input type="password" name="czm" id="czm"></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                          <td><input type="submit" name="button" id="button" value="��ʼ�һ�" <%if stupfhc=0 then%> disabled="disabled"<%end if%>/>
                              <input type="hidden" name="action" value="1"></td>
                        </tr>
                      </table>
                    </form>

              <table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;���ֻ�������</STRONG></td>
                </tr>
              </table>
              <form name="form5" action="" method="post" onSubmit="return save_onclick5();">
                      <div>
                        <table width="614" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td height="35" align="right"><div align="right">�һ��ˣ�</div></td>
                            <td><%=session("useridname")%> �����ڵĻ����ǣ�<%=jifei%> </td>
                          </tr>
                          <tr>
                            <td width="134" height="35" align="right"><div align="right">�����ã�</div></td>
                            <td width="480"><input id=ReNum1  maxlength=4 name=ReNum1 onKeyUp="if(isNaN(value))execCommand('undo')">
                              ��������( ÿ<%=stupjifei%>���ֿɶһ�1��������) </td>
                          </tr>
                          <tr>
                            <td height="35"><div align="right">�����룺</div></td>
                            <td><input type="password" name="czm" id="czm"></td>
                          </tr>
                          <tr>
                            <td height="35">&nbsp;</td>
                            <td><input type="submit" name="button" id="button" value="��ʼ�һ�" />
                                <input type="hidden" name="action" value="5"></td>
                          </tr>
                        </table>
                      </div>
                    </form>
					
					<table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;�����㻥��</STRONG></td>
                </tr>
              </table>
              <form name="form3" action="" method="post" onSubmit="return save_onclick3();">
                      <div>
                        <table width="614" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td height="35" align="right"><div align="right">�����ˣ�</div></td>
                            <td><%=session("useridname")%> �����ڵķ������ǣ�<%=fabudian%> </td>
                          </tr>
                          <tr>
                            <td width="134" height="35" align="right"><div align="right">��Ҫ���͸���</div></td>
                            <td width="480"><input id=ToUser  maxlength=16 name=ToUser >
                              �Է��ʺ� </td>
                          </tr>
                          <tr>
                            <td height="35"><div align="right">����������</div></td>
                            <td><input id=GiveNum  maxlength=4  name=GiveNum  onKeyUp="if(isNaN(value))execCommand('undo')"></td>
                          </tr>
                          <tr>
                            <td height="35"><div align="right">�����룺</div></td>
                            <td><input type="password" name="czm" id="czm"></td>
                          </tr>
                          <tr>
                            <td height="35">&nbsp;</td>
                            <td><input type="submit" name="button" id="button" value="��ʼ����"  <%if isgive=1 then%> disabled="disabled" <%end if%> />
                                <input type="hidden" name="action" value="3"></td>
                          </tr>
                        </table>
                      </div>
                    </form>
				  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close()
set rs=nothing
call closeconn()%>
</BODY></HTML>
