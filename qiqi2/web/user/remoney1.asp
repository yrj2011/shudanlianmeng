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
if stupphonestart=1 and dst =0 then
	Response.Write("<script>alert('���Ȱ��ֻ����ܽ�����������!');window.location.href='../user/Center_Userlock.asp';</script>")
	response.End()
end if
 action=HtmlEncode(trim(request.form("action")))
if action="ok" then    
	 ReNum=HtmlEncode(trim(request.form("ReNum1")))
	 ReRName=HtmlEncode(trim(request.form("ReRName1")))
	 ReZfb=HtmlEncode(trim(request.form("ReZfb1")))
	 czm1=HtmlEncode(trim(request.form("czm")))
 
 cunkuan2=cunkuan*100
 renum2=renum*100
if czm1<>czm then
	Response.Write("<script>alert('������Ĳ����벻��ȷ!');history.back();</script>")
	 response.End()
elseif int(cunkuan2)<int(renum2) then
	Response.Write("<script>alert('�����ֽ��ܴ�����Ĵ��!');history.back();</script>")
    response.End()
elseif ReRName="" then
     Response.Write("<script>alert('�����������Ϊ��!');history.back();</script>")
	 response.End()
elseif ReZfb="" then
     Response.Write("<script>alert('���֧�����˺Ų���Ϊ��!');history.back();</script>")
	 response.End()
elseif int(ReNum)<10 then
	Response.Write("<script>alert('����10Ԫ�����֣����Է����������ʽ���ñ��˽���������񣬴���ת�����֧������!');history.back();</script>")
    response.End()
elseif not isNumeric(ReNum) then
    Response.Write("<script>alert('�����������������룡');history.back();</script>")
    response.End()
end if
  
			
Sql = "select * from "&jieducm&"_tixian where username='"&session("useridname")&"' and class='2'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,3,3
if not rs.eof then
	if rs("ReZfb")<>rezfb then
		 Response.Write("<script>alert('�Բ������������һ�����ְ󶨵�֧������!');history.back();</script>")
		 response.End()
	end if
end if
		
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)	
	    Set rs=server.createobject("ADODB.RECORDSET") 
	 	rs.open "Select * From "&jieducm&"_tixian where username='"&session("useridname")&"' and (start='0' or start='3') " ,Conn,3,3  
		if not rs.eof then
			Response.Write("<script>alert('��һ�����ֻ�δ���!');history.back();</script>")
            response.End()
		else
	    rs.addnew
		rs("class")=2
		rs("pid")=num
		rs("ReNum")=ReNum
		rs("start")=0
    	rs("ReRName")=ReRName
		rs("ReZfb")=ReZfb
		rs("username")=session("useridname")		
    	rs.update
    	rs.close
		end if
		
sqlr="update "&jieducm&"_register set cunkuan=cunkuan-"&renum&" where username='"&session("useridname")&"'"
conn.execute sqlr
call check_jieducm_name(session("useridname"))				
				
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="��������"
		rs("content")=session("useridname")&"������������"&ReNum&"Ԫ"
		rs("price")="-"&renum
		rs("jiegou")="������"
    	rs.update
    	rs.close
    Response.Write("<script>alert('���ѳɹ���������,\n���ǽ��ڵڶ��������ս�����ת������֧����!');window.location.href='ReMoneyList.asp';</script>")
	response.End()	
set rs=nothing
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
function  save_onclick22()
{

    var Price=document.formt.ReNum1.value;
  if(Price=="")
  {
    alert("���������ֽ�");
	document.formt.ReNum1.focus();
	return false;
	}
 if(!Number(Price))
	  
  {
    alert("������������!");
	document.formt.ReNum1.focus();
	return false;
	}
	if(Price<10)
	{
	alert("����10Ԫ�����֣����Է����������ʽ���ñ��˽���������񣬴���ת�����֧������!");
	document.formt.ReNum1.focus();
	return false;
	}
if   ((document.formt.ReNum1.value.indexOf("-")   ==   0)||!(document.formt.ReNum1.value.indexOf(".")   ==   -1)){   
  alert("���ֽ���ΪС������");   
  document.formt.ToUser.focus();   
  return   false;   
  } 
  var ReRName1=document.formt.ReRName1.value;
  if(ReRName1=="")
  {
    alert("������������ʵ���֣�");
	document.formt.ReRName1.focus();
	return false;
	}

 var ReZfb1=document.formt.ReZfb1.value;
  if(ReZfb1=="")
  {
    alert("����������֧�����˺ţ�");
	document.formt.ReZfb1.focus();
	return false;
	}
 var czm=document.formt.czm.value;
  if(czm=="")
  {
    alert("���������Ĳ����룡");
	document.formt.czm.focus();
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
               
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; �������� &gt;&gt; </div>
                    <div class=pp8><strong>��������</strong></div>
                    <div style="MARGIN-TOP: 10px; LINE-HEIGHT: 350%">
                      <ul>
                        <li>* ƽ̨�ṩ�������ַ�ʽ���Ա���Ʒ���ֺ�֧�������ֺͲƸ�ͨ��<br>
                          * ��Ҳ�����������б��ڳ������֣����Զ�Ϊ���������ؽ�������û�����<br>
                          *   ���Ĳ�����¼��Ҳ������Ӧ�ļ�¼��Ϣ��<br>
                          *   ��������������Ʒ���ӵ���10Ԫ��֧��������10Ԫ�����֣����Է����������ʽ���ñ��˽����������񣬴���ת�����ĵ���֧�����ˣ� </li>
                      </ul>
                    </div><br> 
					 
 <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
              <form name="formt" action="" method="post" onSubmit="return save_onclick22();">
					  <input name="action" type="hidden" value="ok">   <tr>
                  <td height="35" colspan="2" align="center"> ÿ��ֻ�ܽ���һ�����֣�����������ɺ���ܽ�����һ�����֣� </td>
                  </tr>
                <tr>
                  <td height="35" align="right">&nbsp;</td>
                  <td>
				  <input type="radio" name="radiobutton" value="1" onClick="javascript:location.href='remoney.asp';"/>
				  �Ա���Ʒ��ַ����                  
                  <input name="radiobutton" type="radio"  value="2" checked/>
                  ֧��������               <input type="radio" name="class" value="3" onClick="javascript:location.href='remoney2.asp';"/>
                  �Ƹ�ͨ����        </td>
                </tr>
				 <tr>
                  <td height="35" align="right"><font color="#FF0000">��һ����</font></td>
                  <td>�����������ύ�տ�ף����ύ����ĵڶ�����<font color="#FF0000">400Ԫ��������</font></td>
                </tr>
				 <tr>
                  <td height="35" align="right">�������ֽ�</td>
				  <td width="80%" align="left"><input name="CZMoney" id="CZMoney" type="text" class="textbox" size="25" maxlength="8" onKeyUp="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" onBlur="this.value=this.value.replace(/\D/g,'')" value="" style="color:#060; font-weight:bold; font-size:14px;" /></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="left" height=30><A href="https://shenghuo.alipay.com/transfer/gathering/create.htm?optEmail=<%=stupzfb%>&totleFee=&title=����:<%=session("useridname")%>" target=_blank> <img src="/user/images/tx.jpg" ></A></td>
                </tr>
				<tr>
                  <td height="35" align="right">&nbsp;</td>
                </tr>
				<tr>
                  <td height="35" align="right"><font color="#FF0000">�ڶ�����</font></td>
                </tr>
                <tr>
                  <td height="35" align="right">�����ˣ�</td>
                  <td><%=session("useridname")%> </td>
                </tr>
                <tr>
                  <td width="134" height="35" align="right">���ֽ�</td>
                  <td width="366"><INPUT id=ReNum1  maxLength=4 name=ReNum1 onKeyUp="if(isNaN(value))execCommand('undo')"></td>
                </tr>
                <tr>
                  <td height="35" align="right">���������</td>
                  <td><INPUT id=ReRName1 maxLength=50 name=ReRName1></td>
                </tr>
                <tr>
                  <td height="35"  align="right">���֧�����˺ţ�</td>
                  <td><INPUT name=ReZfb1 id=ReZfb1 size="25" maxLength=255></td>
                </tr>
                <tr>
                  <td height="35"><div align="right">�����룺</div></td>
                  <td><input type="password" name="czm" id="czm"></td>
                </tr>
                <tr>
                  <td height="35">&nbsp;</td>
                  <td><input type="submit" name="button" id="button" value="��ʼ����" /></td>
                </tr></FORM>
              </table>
  </td>
	          </tr>
	        </table>	  </td>
	  </tr>
</table>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
