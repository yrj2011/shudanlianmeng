<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->

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
if vip<>1 then
	Response.Write("<script>alert('�㻹����vip��Ա!���ܷ���vip����');window.location.href='../user/login.asp';</script>")
    response.End()
elseif session("czm")="" then
 response.Redirect("code.asp?id=fa")
 response.End()
end if
call myww(1)
response.charset="gb2312"
action=request.QueryString("go")
if action="check" then
url = request("url")
nickname = trim(request("nickname"))
content = GetNickName(url)
if nickname<>content then
echo "no"
else
echo "yes"
end if
response.End()
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh" lang="zh" dir="ltr">
<head>
<title><%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/style.css" type=text/css rel=stylesheet>
<LINK href="../css/base.css" type=text/css rel=stylesheet>
<LINK href="../css/layout.css" type=text/css rel=stylesheet>
 <script type="text/javascript" src="../js/common5.js"></script>
 <script type="text/javascript" src="../js/jieducm_jsdate.js"></script>
</head>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->
<div align="center">  
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
<UL id=task_nav>
 <li><a  href="index.asp">�Ա�������</a> </li>
    <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="pushmissionvip.asp">��������</a> </li>
  <LI><A href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
  <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI>
   <li><a href="../user/Express.asp">���ɿ�ݵ�</a> </li>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>
</div>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr height="1">
                  <td height="5"></td>
      </tr>
  </table>
              <table width="910"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr height="1">
                  <td height="5"></td>
      </tr>
  </table>
              <DIV class=page>
<DIV class=addtask-wrap>
<DIV class=inner>
��������ǰ��֪�� 
<UL>
  <LI>1. ����һ�������۳����Ĵ��뱣֤����ƽ̨�����㣻 
  <LI>2. �뱣֤���ڱ�վ���㹻�Ĵ�������������۸���Ӵ������۳������˽ӵ����������Ϊ��Ʒ��ظ��㡣 </LI>
    
  <li>3. (����۸��ǰ����ʷѵ��ܼ۸�)1-40Ԫ����һ�������㣻41-100Ԫ�����������㣻101-200Ԫ�������������㣻201-500Ԫ�����ĸ������㣻501-1000Ԫ�����������</li>
  </UL></DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
     <FORM name=form id="form" method=post action="jieducm_faokvip.asp" >
           <input name="fa" type="hidden" value="ok">
  <TR>
    <TH><div align="right">����ѡ��������������</div></TH>
    <TD><LABEL><INPUT id="num" onclick=javascript:gouwuce(1); type=radio CHECKED value="1" name=num>
                                    һ��ˢ</LABEL>
                                     <LABEL><INPUT id="num"  onclick=javascript:gouwuce(2); type=radio  value="2" name=num>
                                    ����ˢ</LABEL>
                                    <LABEL><INPUT id="num" onclick=javascript:gouwuce(3); type=radio  value="3" name=num>
                                    ����ˢ</LABEL>
                                   <LABEL > <INPUT id="num" onclick=javascript:gouwuce(4); type=radio  value="4" name=num>
                                    ����ˢ</LABEL>
                                    <LABEL><INPUT id="num" onclick=javascript:gouwuce(5); type=radio  value="5" name=num>
                                    ����ˢ</LABEL></TD></TR>
  <TR>
    <TH><div align="right">��&nbsp; ��&nbsp; ����</div></TH>
    <TD ><%call shopname(1)%></TD>
  </TR>
  
  <TR>
    <TH><div align="right">����1��</div></TH>
    <TD >��Ʒ����ۣ�<input name=Price1 id=Price1 size="10" maxlength=6  onKeyUp="if(isNaN(value))execCommand('undo')">
	��Ʒ��ַ��  <input id=ProUrl1 maxlength=255 name=ProUrl1 onBlur="check(this)">	</TD></TR>
  <TR id=renwu2 style="DISPLAY: none">
    <TH><div align="right">����2��</div></TH>
    <TD >��Ʒ����ۣ�<input name=Price2 id=Price2 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	��Ʒ��ַ��  <input id=ProUrl2 maxlength=255 name=ProUrl2 onBlur="check(this)" >	</TD></TR>
	<TR id=renwu3 style="DISPLAY: none">
    <TH><div align="right">����3��</div></TH>
    <TD >��Ʒ����ۣ�<input name=Price3 id=Price3 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	��Ʒ��ַ��  <input id=ProUrl3 maxlength=255 name=ProUrl3  onBlur="check(this)">	</TD></TR>
	<TR id=renwu4 style="DISPLAY: none">
    <TH><div align="right">����4��</div></TH>
    <TD >��Ʒ����ۣ�<input name=Price4 id=Price4 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	��Ʒ��ַ��  <input id=ProUrl4 maxlength=255 name=ProUrl4  onBlur="check(this)">	</TD></TR>
	<TR id=renwu5 style="DISPLAY: none">
    <TH><div align="right">����5��</div></TH>
    <TD >��Ʒ����ۣ�<input name=Price5 id=Price5 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	��Ʒ��ַ��  <input id=ProUrl5 maxlength=255 name=ProUrl5  onBlur="check(this)">	</TD></TR>
 		<TR id=renwu6 style="DISPLAY: none">
    <TH><div align="right">����6��</div></TH>
    <TD >��Ʒ����ۣ�<input name=Price6 id=Price6 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	��Ʒ��ַ��  <input id=ProUrl6 maxlength=255 name=ProUrl6  onBlur="check(this)">	</TD></TR>
         <TR>
    <TH><div align="right">�Ƿ���·��������</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><label>
      <input type="checkbox" name="jieducm_IfComeSet" id="jieducm_IfComeSet" value="1" onClick="javascript:SwitchDivDisplay('showinfo_c1')" >
      ��·��������
    </label>
    (ѡ���������1��������)</TD>
  </TR>
      <tr id=showinfo_c1 style="DISPLAY: none">
      <TH><div align="right">��·����ѡ��</div></TH>
      <TD colspan="3" 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><label><INPUT 
      onclick=setTaskid(1) value=1 CHECKED type=radio name=IfComeSet>�Ա���������</label><SPAN 
      style="MARGIN-LEFT: 20px"> �����Ա��������̽�������������� </SPAN></SPAN><BR><SPAN 
      style="MARGIN-RIGHT: 10px"><label><INPUT onclick=setTaskid(2) value=2 type=radio 
      name=IfComeSet>�Ա���������</label><SPAN style="MARGIN-LEFT: 20px"> �����Ա�ֱ��������������� </SPAN></SPAN><BR></TD>
    </TR>
	 <tr id=showinfo_c2 style="DISPLAY: none">
      <TH><div align="right"><SPAN id=tt1>�������̹ؼ���</SPAN>��</div></TH>
      <TD colspan="3" 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<INPUT style="WIDTH: 290px" id=ComeKey class=fa-input maxLength=50  type=text name=ComeKey><SPAN id=ttt1 
      class=fa-righttext>�˴���������Ҫ���߽�������Ʒ���Լ������ڵ�ʲô�ط������ҵ������磬���������������100Ԫ���ӿ�����Ʒ����һ�����ǡ� 
      </SPAN></TD>
    </TR>
    <tr id=showinfo_c3 style="DISPLAY: none">
      <TH><div align="right"><SPAN id=tt2>������ʾ</SPAN>��</div></TH>
      <TD colspan="3" 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"> 
	<INPUT style="WIDTH: 290px" id=ComeNote class=fa-input maxLength=50  type=text name=ComeNote><SPAN id=ttt2 
      class=fa-righttext>�˴���д��ʾ��Ϣ��˵�������ڹؼ�����������б��е�λ�ã���Ʒ�ڵ�����ҳ�д��λ�õ���Ϣ�����磺�����ڽ���б�ڶ�������Ʒ����Ʒ��ҳ�ڶ����м�һ���ȵȡ�</SPAN></TD>
    </TR>
		
		<tr>
          <TH><div align="right">���ӷ����㣺</div></TH>
          <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
            <label>
              <select name="addfabu">
                <option value="0">������</option>
                <option value="1">����1��</option>
                <option value="2">����2��</option>
                <option value="3">����3��</option>
                <option value="4">����4��</option>
                <option value="5">����5��</option>
              </select>
              </label>
          
         &nbsp; (�ײ������ѡ)</TD>
        </TR>
        <TH><div align="right">�۸��Ƿ���ȣ�</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="isprice" type="radio" id="isprice" value="������" checked>
                              <span class="font12l"> ������</span> <input type="radio" name="isprice" id="isprice" value="���޸ļ۸�">
                               <span class="font12l">���޸ļ۸�</span>&nbsp; ����۸�Ͱ����˷ѵ���Ʒ�ܼ۸��Ƿ���ȣ�</TD>
  </TR>
  <TR>
    <TH><div align="right">��̬���֣�</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="play" type="radio" value="ȫ����5��" checked>
                              <span class="font12l">ȫ����5��</span> <input type="radio" name="play" value="ȫ�������">
                             <span class="font12l"> ȫ�������</span></TD>
  </TR>
   <TR>
    <TH></TH>
    <TD><FONT color=#ff0000><STRONG>* 
      ������ʱ�ջ�������ƽ̨����ṩ�������ţ���ǿ�������ʱ�ջ�</STRONG></FONT></TD></TR>
	  <TR>
    <TH><div align="right">������ʽ��</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
      <label>
        <input name="ping" type="radio" value="���ֺ���" checked onclick=javascript:jieducm_ping(1);> ���ֺ���</label>
      <label>
        <input type="radio" name="ping" value="�����ֺ���" onclick=javascript:jieducm_ping(2);>�����ֺ��� </label>
		    <label>
        <input type="radio" name="ping" value="�Զ�������" onclick=javascript:jieducm_ping(3);>�Զ������� </label>
			    <label>
        <input type="radio" name="ping" value="Ĭ�Ϻ���" onclick=javascript:jieducm_ping(4);>Ĭ�Ϻ���(������)</label>    </TD>
  </TR>
  <TR id=pinghistory style="DISPLAY: none">
    <TH><div align="right">�Զ������</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
        <INPUT   name=pingtxt class=txtinput id=pingtxt style="WIDTH: 290px" maxlength="100">
    <select name="lsx"  style="width:150px;" onChange="jieducm_pingtxt(this.options[this.options.selectedIndex].value);"> 
								<option value="">��ѡ����ʷ�Զ�������</option>
<%Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 10 pingtxt from "&jieducm&"_zhongxin where username='"&session("useridname")&"' and classid='1'  and pingtxt<>'' order by  id desc",conn,1,1
if not (rs.eof) then
Do While Not Rs.Eof
%>
<option value="<%=rs("pingtxt")%>"><%=rs("pingtxt")%></option>
<%
   Rs.MoveNext
   Loop
   End IF
%>
						    </select>    </TD>
  </TR>
  <TR>
    <TH><div align="right">��ע��</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<LABEL><INPUT name="bei" type="radio" id="bei"  value="�����ջ�����" checked> 
	<SPAN class=font12l>�����ջ�����</SPAN></LABEL>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;	<LABEL><INPUT type="radio" name="bei" id="bei"  value="һ����ջ�����"> <SPAN class=font12l>һ����ջ�����</SPAN></LABEL>(��x*2��������)<BR>
    <LABEL><input type="radio" name="bei" id="bei" value="������ջ�����"> <span class="font12l">������ջ�����</span></LABEL>(��x*2+1��������)
  <LABEL><input type="radio" name="bei" id="bei" value="������ջ�����"> <span class="font12l">������ջ�����</span></LABEL>(��x*2+2��������)<BR>	</TD></TR>
  <TR>
    <TH><div align="right">���������ƣ�</div></TH>
    <TD><select name="Limit" >
                                <option value="1" selected>������</option>
                                <option value="2">����100��������</option>
                                <option value="3">����300��������</option>
                                <option value="4">����ֻ����VIP��Ա</option>
                                                                                          </select></TD>
  </TR>
  <TR>
    <TH><div align="right">���������</div></TH>
    <TD><INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100">
	<select name="lsx"  style="width:120px;" onChange="jieducm_tips(this.options[this.options.selectedIndex].value);"> 
								<option value="">��ѡ����ʷ������</option>
<%Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 10 tips from "&jieducm&"_zhongxin where username='"&session("useridname")&"' and classid='1'  and tips<>'' order by  id desc",conn,1,1
if not (rs.eof) then
Do While Not Rs.Eof
%>
									<option value="<%=rs("tips")%>"><%=rs("tips")%></option>
<%
   Rs.MoveNext
   Loop
   End IF
%>
						    </select></TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> ����·���� </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> �����걣��</LABEL>	 <LABEL> <input id=baohu32 disabled type=checkbox checked value=1  name=baohu32>�Զ���ⱦ����ַ���ƹ���</LABEL></TD></TR>
  
 <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT name="Package"   type=checkbox id="Package"  disabled value=1 checked>
      �ײ����� </LABEL>
      (�����ײ����������ӷ�����)</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><label><input  name="baohu2" type="checkbox" id="baohu2" value="1"  />  
                  ���񱣻��������߽��������Ҫ����˲ſ��Կ�����Ʒ��ַ��</label></TD>
  </TR>
  <TR>
    <TH><div align="right">��ʱ������</div></TH>
    <TD><input name="Timing" type="text"  onClick="SelectDate(this,'yyyy-MM-dd hh:mm:ss')"  readonly/>
                              ��������Ч</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><input  name="depot" type="checkbox" id="depot" value="1" />˳����ӵ��ҵ�����ֿ�</LABEL>&nbsp; �ֿⱸע(�������Լ��ֱ���Ʒ)�� 
                               <label>
                               <input name="title" type="text" maxlength="20">
                               </label></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px" 
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button" <% if cunkuan<1 or fabudian<1 then%> disabled <%end if%> value="��������"  onClick="ajax()"><% if cunkuan<1 or fabudian<1 then%>���򷢲������1ʱ�����ٷ�����<%end if%></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>    </td>
              </tr>
      </table>	    </td>
	    </tr>
  </table>
  <script>
function ajax(){
document.form.button.disabled="disabled";
document.form.button.value="ϵͳ���ڴ�����";
document.form.submit();
}

  function jieducm_tips(s)
 {
var ss=s;
document.getElementById("tips").value=ss;

 }
   function jieducm_pingtxt(s)
 {
var ss=s;
document.getElementById("pingtxt").value=ss;
 }
 function jieducm_ping(num)
 {
 if (num==3)
	{document.getElementById("pinghistory").style.display="";}
	else
	{
		{document.getElementById("pinghistory").style.display="none";}

	}
 }
 
 function gouwuce(num)
 {	
 //alert(num);
	atuo(num);
 }
function atuo(num)
 {
	if (num==1)
	{
		document.getElementById("num").value=1;
		document.getElementById("jieducm_IfComeSet").disabled="";
		allnone();
		
	}
	else
	{
		//alert(num);
		document.getElementById("num").value=num;
		document.getElementById("jieducm_IfComeSet").disabled="disabled";
		document.getElementById("jieducm_IfComeSet").checked="";
		document.getElementById("showinfo_c1").style.display="none";
		document.getElementById("showinfo_c2").style.display="none";
		document.getElementById("showinfo_c3").style.display="none";
		allnone();
		for(var i=2;i<=num;i++)
		{
			document.getElementById("renwu"+i).style.display="";
		}
	}
 }
function allnone()
 {
	
	for( var i = 2;i <= 5;i++)
		{
			document.getElementById("renwu"+i).style.display="none";
		}
 }
 onload=function(){
// alert('ss');
 var nums=document.getElementById("num").value;
 //alert(nums);
 gouwuce(nums);

 }
 
 function setTaskid(num)
	  {
		  var str1="";
		  var str2="";
		  var tstr1="";
		  var tstr2="";
		  var ttstr="";
		  if(num==1)
		  {
			  str1="�������̹ؼ���";
			  str2="�˴���������Ҫ���߽�������Ʒ���Լ������ڵ�ʲô�ط������ҵ������磬���������������100Ԫ���ӿ�����Ʒ����һ�����ǡ�";
			  tstr1="������ʾ";
			  tstr2="�˴���д��ʾ��Ϣ��˵�������ڹؼ�����������б��е�λ�ã���Ʒ�ڵ�����ҳ�д��λ�õ���Ϣ�����磺�����ڽ���б�ڶ�������Ʒ����Ʒ��ҳ�ڶ����м�һ���ȵȡ�";
			  //ttstr="�����Ա��������̽��������������";
		   }
		   else if(num==2)
		   {
			  str1="���������ؼ���";
			  str2="�Ƽ�ʹ�����ı������ƣ�������ı����������Ա���������Ʒ���࣬���ǽ������޸ı������ƺ������ô�����·�������ʹ�õ�����������·��ʽ��";
			  tstr1="������ʾ";
			  tstr2="�˴���д��ʾ��Ϣ��˵��������Ʒ�ڹؼ�����������б��е�λ�ã������һҳ��7����";
			  //ttstr="�����Ա�ֱ���������������";
			}
		   else if(num==3)
		   {
			  str1="����������۵�ַ";
			  str2="�˴������Ѿ���������ĸ����񱦱���ҵ������������ҳ���ַ�����������ı������ۼ�¼�е������û����Ҳ�����õȼ�ͼ�꣨���ģ�����ȣ����롣";
			  tstr1="��·��ʾ";
			  tstr2="�˴���д��ʾ��Ϣ��˵���˱��������¼�ڹ����б��е�ҳ����λ�õ���Ϣ�����磺�ڶ�ҳ������һ����";
			  //ttstr="�Ǵ��Ա���������ҳ�����ģ������������۵ı�����Ϣ��½���ı���";
		   }
		   window.tt1.innerHTML=str1;
		   window.ttt1.innerHTML=str2;
		   window.tt2.innerHTML=tstr1;
		   window.ttt2.innerHTML=tstr2;
		   //window.th1.innerHTML=ttstr;
		  
	  }
	  function SwitchDivDisplay(divName) {
		var obj_reg_info = document.getElementById(divName);
	   	if(obj_reg_info.style.display=="none") {
	       //	obj_reg_info.style.display="inline";
			document.getElementById("showinfo_c1").style.display="";
			document.getElementById("showinfo_c2").style.display="";
			document.getElementById("showinfo_c3").style.display="";
	   	} else {
	   		//obj_reg_info.style.display="block";
	   		//obj_reg_info.style.display="none";
			document.getElementById("showinfo_c1").style.display="none";
			document.getElementById("showinfo_c2").style.display="none";
			document.getElementById("showinfo_c3").style.display="none";
		}
	}
</script>
  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
   call CloseConn()%>
</body>
</html>